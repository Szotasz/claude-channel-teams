/**
 * In-memory (and, with a file, persisted) record of the ids of inbound messages
 * already received. A repeated delivery of the same message is dropped; the id
 * is noted on arrival and forgotten if the turn fails, so a retry still goes
 * through. Bounded.
 *
 * With `file` (the server passes `<stateDir>/inbound-ids.json` when the Graph
 * backfill is on, TEAMS_BACKFILL) the ids survive a restart: a message the live
 * path delivered before the process exited is not delivered a second time by the
 * backfill, and the other way round. The file is JSON
 * `{ "version": 1, "conversations": [[key, [id, ...]], ...] }`, oldest first,
 * mode 0600, written atomically (`.tmp` + rename) on every change — the same
 * pattern as the conversation reference store. A file that is unreadable,
 * not ours, writable by others, or not exactly that shape is ignored with ONE
 * stderr line and the store starts empty; a failing save logs one line per
 * process and never throws into the turn handler.
 */

import {
  closeSync,
  constants,
  existsSync,
  fstatSync,
  mkdirSync,
  openSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'fs'
import { dirname } from 'path'

export const MAX_IDS_PER_CONVERSATION = 500
export const MAX_ID_CONVERSATIONS = 1000

export interface InboundIds {
  /** Records the id and returns true the first time; returns false when it was already recorded. */
  firstSeen(conversationKey: string, messageId: string): boolean
  /** Removes the id again (used when the turn failed, so a retry is processed). */
  forget(conversationKey: string, messageId: string): void
}

export interface InboundIdsOptions {
  /** When set, the ids are loaded from and saved to this JSON file (mode 0600, written atomically). */
  file?: string
  /** Owner check override for tests (default: process.getuid). */
  uid?: () => number | undefined
}

function errMsg(err: unknown): string {
  return err instanceof Error ? err.message : String(err)
}

function ownerUid(uid?: () => number | undefined): number | undefined {
  if (uid) return uid()
  return typeof process.getuid === 'function' ? process.getuid() : undefined
}

/** Parses the file body; throws a message that carries nothing of the file's content. */
function parseIds(text: string): Array<[string, string[]]> {
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch {
    throw new Error('unexpected format')
  }
  if (!parsed || typeof parsed !== 'object' || (parsed as { version?: unknown }).version !== 1 || !Array.isArray((parsed as { conversations?: unknown }).conversations)) {
    throw new Error('unexpected format')
  }
  const out: Array<[string, string[]]> = []
  for (const entry of (parsed as { conversations: unknown[] }).conversations) {
    if (!Array.isArray(entry) || entry.length !== 2) throw new Error('unexpected format')
    const [key, ids] = entry as [unknown, unknown]
    if (typeof key !== 'string' || key.length === 0 || key.length > 512) throw new Error('unexpected format')
    if (!Array.isArray(ids)) throw new Error('unexpected format')
    for (const id of ids) {
      if (typeof id !== 'string' || id.length === 0 || id.length > 128) throw new Error('unexpected format')
    }
    out.push([key, ids as string[]])
  }
  return out
}

export function createInboundIds(opts: InboundIdsOptions = {}): InboundIds {
  const file = opts.file
  // Conversation key -> ids, oldest first in both places (the Map's first key is the longest-seen conversation).
  const conversations = new Map<string, string[]>()
  let warned = false

  const warn = (line: string): void => {
    if (warned) return
    warned = true
    process.stderr.write(`${line}\n`)
  }

  /** The ids of a conversation, moved to the newest position (an unseen conversation starts empty). */
  function idsOf(conversationKey: string): string[] {
    const ids = conversations.get(conversationKey) ?? []
    // Delete + set moves an existing key to the newest position.
    conversations.delete(conversationKey)
    conversations.set(conversationKey, ids)
    return ids
  }

  /** Only the conversation just written to can be over its id bound; the oldest conversation is dropped first. */
  function bound(conversationKey: string): void {
    const ids = conversations.get(conversationKey)
    if (ids) while (ids.length > MAX_IDS_PER_CONVERSATION) ids.shift()
    while (conversations.size > MAX_ID_CONVERSATIONS) {
      conversations.delete(conversations.keys().next().value as string)
    }
  }

  function save(): void {
    if (!file) return
    const tmp = `${file}.tmp`
    try {
      const dir = dirname(file)
      if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 })
      writeFileSync(tmp, JSON.stringify({ version: 1, conversations: [...conversations] }), { mode: 0o600 })
      renameSync(tmp, file)
    } catch (err) {
      try { rmSync(tmp, { force: true }) } catch {}
      warn(`teams channel: could not save inbound ids (${errMsg(err)})`)
    }
  }

  if (file) {
    let fd: number
    try {
      // O_NONBLOCK: opening a FIFO here must not hang startup before the regular-file check below.
      fd = openSync(file, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK)
    } catch (err: any) {
      if (err?.code !== 'ENOENT') {
        warn(`teams channel: inbound-ids file cannot be opened (${err?.code === 'ELOOP' ? 'a symlink' : errMsg(err)}), starting empty`)
      }
      fd = -1
    }
    if (fd >= 0) {
      try {
        const st = fstatSync(fd)
        if (!st.isFile()) warn('teams channel: inbound-ids file is not a regular file, starting empty')
        else {
          const me = ownerUid(opts.uid)
          if (me !== undefined && st.uid !== me) warn('teams channel: inbound-ids file belongs to another user, starting empty')
          else if (st.mode & 0o022) warn('teams channel: inbound-ids file is writable by group or others, starting empty')
          else {
            const loaded = parseIds(readFileSync(fd, 'utf8'))
            for (const [key, ids] of loaded) {
              // A conversation listed twice keeps its last entry, and lands in the newest position.
              const kept = [...ids].slice(-MAX_IDS_PER_CONVERSATION)
              conversations.delete(key)
              conversations.set(key, kept)
            }
            while (conversations.size > MAX_ID_CONVERSATIONS) {
              conversations.delete(conversations.keys().next().value as string)
            }
          }
        }
      } catch (err) {
        warn(`teams channel: inbound-ids file unreadable, starting empty (${errMsg(err)})`)
      } finally {
        try { closeSync(fd) } catch {}
      }
    }
  }

  return {
    firstSeen(conversationKey, messageId) {
      const ids = idsOf(conversationKey)
      if (ids.includes(messageId)) return false
      ids.push(messageId)
      bound(conversationKey)
      save()
      return true
    },
    forget(conversationKey, messageId) {
      const ids = conversations.get(conversationKey)
      if (!ids) return
      const at = ids.indexOf(messageId)
      if (at < 0) return
      ids.splice(at, 1)
      if (ids.length === 0) conversations.delete(conversationKey)
      save()
    },
  }
}
