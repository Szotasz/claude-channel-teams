import type { ConversationReference } from 'botbuilder'

/** A targeted message expires after 24 hours. */
export const TARGETED_TTL_MS = 24 * 60 * 60 * 1000
export const MAX_ENTRIES_PER_CONVERSATION = 200
export const MAX_CONVERSATIONS = 1000
export const MAX_SENT_IDS = 2000

export interface TargetedEntry {
  messageId: string
  ref: Partial<ConversationReference>
  aadObjectId: string
  at: number
}

/**
 * In-memory record of inbound targeted messages (so a reply can go back to the
 * sender privately), of whether the latest delivered message of a conversation
 * was targeted, and of targeted messages the bot sent. Nothing is persisted.
 */
export function createTargetedInbound() {
  const entries = new Map<string, Map<string, TargetedEntry>>()
  const latest = new Map<string, { targeted: boolean; at: number }>()
  const sent = new Set<string>()
  // Set when the bound forced out a fresh private marker: until then an untracked conversation counts as private.
  let unknownPrivateUntil = 0

  function dropOldestKey<K>(m: Map<K, unknown> | Set<K>): void {
    const first = m.keys().next()
    if (!first.done) m.delete(first.value)
  }

  function remember(
    conversationId: string,
    messageId: string,
    ref: Partial<ConversationReference>,
    aadObjectId: string,
    nowMs: number,
  ): void {
    let conv = entries.get(conversationId)
    if (!conv) {
      conv = new Map()
      entries.set(conversationId, conv)
      while (entries.size > MAX_CONVERSATIONS) dropOldestKey(entries)
    }
    conv.delete(messageId)
    conv.set(messageId, { messageId, ref, aadObjectId, at: nowMs })
    while (conv.size > MAX_ENTRIES_PER_CONVERSATION) dropOldestKey(conv)
  }

  function noteDelivered(conversationId: string, targeted: boolean, nowMs: number): void {
    if (!latest.has(conversationId)) {
      // Make room with public or expired markers first. A fresh private marker is dropped only when nothing
      // else is left, and then it fails closed: every untracked conversation counts as private until it expires.
      for (const [key, value] of latest) {
        if (latest.size < MAX_CONVERSATIONS) break
        if (!value.targeted || nowMs - value.at > TARGETED_TTL_MS) latest.delete(key)
      }
      while (latest.size >= MAX_CONVERSATIONS) {
        const [key, value] = latest.entries().next().value as [string, { targeted: boolean; at: number }]
        latest.delete(key)
        unknownPrivateUntil = Math.max(unknownPrivateUntil, value.at + TARGETED_TTL_MS)
      }
    }
    latest.delete(conversationId)
    latest.set(conversationId, { targeted, at: nowMs })
  }

  function get(conversationId: string, messageId: string, nowMs: number): TargetedEntry | undefined {
    const e = entries.get(conversationId)?.get(messageId)
    if (!e || nowMs - e.at > TARGETED_TTL_MS) return undefined
    return e
  }

  function latestIsTargeted(conversationId: string, nowMs: number): boolean {
    const l = latest.get(conversationId)
    if (!l) return nowMs <= unknownPrivateUntil
    return l.targeted && nowMs - l.at <= TARGETED_TTL_MS
  }

  function rememberSent(messageId: string): void {
    sent.delete(messageId)
    sent.add(messageId)
    while (sent.size > MAX_SENT_IDS) dropOldestKey(sent)
  }

  function isSentTargeted(messageId: string): boolean {
    return sent.has(messageId)
  }

  return { remember, noteDelivered, get, latestIsTargeted, rememberSent, isSentTargeted }
}

export type TargetedInbound = ReturnType<typeof createTargetedInbound>
