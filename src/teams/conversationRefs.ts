/**
 * In-memory store of conversation references.
 *
 * Bot Framework's outbound path (`adapter.continueConversationAsync`) requires
 * a previously-captured `ConversationReference`. We record one for every
 * activity that clears the allowlist gate, then look it up when Claude calls
 * the `reply` tool.
 *
 * With a `file` option the references are persisted: they are what lets the
 * bot answer after a restart. They are re-checked against the allowlist on
 * load (and the existing outbound gate still checks every send), and a
 * reference whose serviceUrl is not the Teams service host is never loaded.
 * Without `file` the store is in-memory only.
 *
 * We also track which AAD ObjectID originally seeded each reference. The
 * outbound gate (`assertAllowedConversation`) uses this mapping to refuse
 * replies to conversations whose original sender is no longer on the
 * allowlist — defends against a prompt-injected Claude calling `reply` with
 * a conversation_id captured from elsewhere.
 */

import type { ConversationReference } from 'botbuilder'
import { existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'fs'
import { randomBytes } from 'crypto'
import { dirname } from 'path'
import { isAllowedInlineImageHost } from './attachments.js'

interface StoredRef {
  /** The Bot Framework reference used by `continueConversationAsync`. */
  ref: Partial<ConversationReference>
  /** AAD Object ID of the sender at the time of capture. Lowercase. */
  aadObjectId: string
  /** Last time we saw an inbound from this conversation (ms epoch). */
  lastSeen: number
}

/** A 1:1 chat with the bot. Shared conversations (team channels, group chats) are not personal. */
export function isPersonalReference(ref: Partial<ConversationReference> | undefined): boolean {
  const conv = ref?.conversation
  if (!conv) return false
  if (conv.conversationType === 'personal') return true
  // Any other type ('channel', 'groupChat', or one this code does not know) is not 1:1: fail closed.
  if (conv.conversationType !== undefined) return false
  if ((conv as { isGroup?: boolean }).isGroup === true) return false
  if (typeof conv.id === 'string' && conv.id.includes('@thread.')) return false
  return true
}

export interface ConversationRefStore {
  put(conversationId: string, ref: Partial<ConversationReference>, aadObjectId: string): void
  get(conversationId: string): StoredRef | undefined
  size(): number
  /**
   * Return the conversation_id of the most-recently-active conversation.
   * Used by the permission relay to pick the "primary operator" target in v1
   * (single-operator scope; see docs/security.md).
   */
  mostRecentConversationId(): string | undefined
  /** Like mostRecentConversationId(), but only over 1:1 chats (isPersonalReference). */
  mostRecentPersonalConversationId(): string | undefined
}

export interface RefStoreOptions {
  /** When set, references are loaded from and saved to this JSON file (mode 0600, written atomically). */
  file?: string
  /** Re-checked on load: an entry whose sender is no longer allowed is dropped. */
  isAllowed?: (aadObjectId: string) => boolean
  /** Clock override for tests. */
  now?: () => number
  /** Keep at most this many references, the most recently seen ones (default 1000). */
  maxEntries?: number
}

function errMsg(err: unknown): string {
  return err instanceof Error ? err.message : String(err)
}

function isValidEntry(e: any): boolean {
  if (!e || typeof e !== 'object') return false
  const { conversationId, ref, aadObjectId, lastSeen } = e
  if (typeof conversationId !== 'string' || conversationId.length === 0 || conversationId.length > 512) return false
  if (!ref || typeof ref !== 'object' || Array.isArray(ref)) return false
  if (!ref.conversation || typeof ref.conversation !== 'object' || ref.conversation.id !== conversationId) return false
  if (typeof ref.serviceUrl !== 'string' || !isAllowedInlineImageHost(ref.serviceUrl)) return false
  if (typeof aadObjectId !== 'string' || aadObjectId.length === 0) return false
  if (typeof lastSeen !== 'number' || !Number.isFinite(lastSeen)) return false
  return true
}

export function createConversationRefStore(opts: RefStoreOptions = {}): ConversationRefStore {
  const store = new Map<string, StoredRef>()
  const mostRecent = (accept: (stored: StoredRef) => boolean): string | undefined => {
    let bestId: string | undefined
    let bestTs = -Infinity
    for (const [id, stored] of store) {
      if (accept(stored) && stored.lastSeen > bestTs) {
        bestTs = stored.lastSeen
        bestId = id
      }
    }
    return bestId
  }
  const now = opts.now ?? Date.now
  const maxEntries = opts.maxEntries ?? 1000
  const file = opts.file

  function load(path: string): void {
    if (!existsSync(path)) return
    let entries: unknown[]
    try {
      const parsed = JSON.parse(readFileSync(path, 'utf8'))
      if (!parsed || typeof parsed !== 'object' || parsed.version !== 1 || !Array.isArray(parsed.refs)) {
        throw new Error('unexpected format')
      }
      entries = parsed.refs
    } catch (err) {
      process.stderr.write(
        `teams channel: conversation-refs file unreadable, starting empty (${errMsg(err)})\n`,
      )
      return
    }
    let dropped = 0
    const kept: Array<[string, StoredRef]> = []
    for (const e of entries as any[]) {
      let ok = isValidEntry(e)
      if (ok && opts.isAllowed) {
        try {
          ok = opts.isAllowed(e.aadObjectId)
        } catch {
          ok = false
        }
      }
      if (!ok) {
        dropped++
        continue
      }
      kept.push([
        e.conversationId,
        { ref: e.ref, aadObjectId: String(e.aadObjectId).toLowerCase(), lastSeen: e.lastSeen },
      ])
    }
    // Newest first; a conversation listed twice keeps its most recent entry.
    kept.sort((a, b) => b[1].lastSeen - a[1].lastSeen)
    for (const [id, stored] of kept) {
      if (store.size >= maxEntries) break
      if (!store.has(id)) store.set(id, stored)
    }
    if (dropped > 0) {
      process.stderr.write(`teams channel: dropped ${dropped} stored conversation reference(s) on load\n`)
    }
  }

  function save(path: string): void {
    let tmp: string | undefined
    try {
      const dir = dirname(path)
      if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 })
      tmp = `${path}.tmp-${process.pid}-${randomBytes(6).toString('hex')}`
      const refs = [...store].map(([conversationId, s]) => ({ conversationId, ...s }))
      writeFileSync(tmp, JSON.stringify({ version: 1, refs }), { mode: 0o600 })
      renameSync(tmp, path)
    } catch (err) {
      if (tmp) {
        try { rmSync(tmp, { force: true }) } catch {}
      }
      process.stderr.write(`teams channel: could not save conversation references (${errMsg(err)})\n`)
    }
  }

  if (file) load(file)

  return {
    put(conversationId, ref, aadObjectId) {
      // We index by conversation.id because that's what Claude will pass
      // back via the reply tool. Lowercase the AAD ID at storage time so
      // the outbound gate's compare is symmetric with the inbound store.
      store.set(conversationId, {
        ref,
        aadObjectId: aadObjectId.toLowerCase(),
        lastSeen: now(),
      })
      while (store.size > maxEntries) {
        let oldestId: string | undefined
        let oldestTs = Infinity
        for (const [id, stored] of store) {
          if (stored.lastSeen < oldestTs) {
            oldestTs = stored.lastSeen
            oldestId = id
          }
        }
        if (oldestId === undefined) break
        store.delete(oldestId)
      }
      if (file) save(file)
    },
    get(conversationId) {
      return store.get(conversationId)
    },
    size() {
      return store.size
    },
    mostRecentConversationId(): string | undefined {
      return mostRecent(() => true)
    },
    mostRecentPersonalConversationId(): string | undefined {
      return mostRecent(stored => isPersonalReference(stored.ref))
    },
  }
}
