import { readFileSync, writeFileSync, renameSync, chmodSync } from 'fs'
import type { ConversationReference } from 'botbuilder'

export interface StoredChannelRef {
  ref: Partial<ConversationReference>
  aadObjectId: string
  savedAt: string
}

export interface ChannelRefStore {
  put(channelId: string, ref: Partial<ConversationReference>, aadObjectId: string): void
  get(channelId: string): StoredChannelRef | undefined
}

type FileShape = { version: 1; entries: Record<string, StoredChannelRef> }

function readEntries(path: string): Record<string, StoredChannelRef> {
  try {
    const parsed = JSON.parse(readFileSync(path, 'utf8')) as Partial<FileShape>
    if (parsed?.version !== 1 || !parsed.entries || typeof parsed.entries !== 'object') return {}
    return parsed.entries
  } catch {
    return {}
  }
}

export function createChannelRefStore(path: string): ChannelRefStore {
  return {
    get(channelId) {
      const entries = readEntries(path)
      return Object.prototype.hasOwnProperty.call(entries, channelId) ? entries[channelId] : undefined
    },
    put(channelId, ref, aadObjectId) {
      try {
        const entries = readEntries(path)
        const aad = aadObjectId.toLowerCase()
        const prev = entries[channelId]
        // Every message in an enabled channel lands here: rewrite the file only
        // when the route or its sender actually changed.
        if (prev && prev.aadObjectId === aad && JSON.stringify(prev.ref) === JSON.stringify(ref)) return
        entries[channelId] = { ref, aadObjectId: aad, savedAt: new Date().toISOString() }
        const tmp = `${path}.tmp`
        const body: FileShape = { version: 1, entries }
        writeFileSync(tmp, JSON.stringify(body, null, 2), { mode: 0o600 })
        chmodSync(tmp, 0o600)
        renameSync(tmp, path)
      } catch (err) {
        process.stderr.write(`teams channel: channel ref save failed: ${err}\n`)
      }
    },
  }
}
