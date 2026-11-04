/**
 * In-memory record of the ids of inbound messages already received. A repeated
 * delivery of the same message is dropped; the id is noted on arrival and
 * forgotten if the turn fails, so a retry still goes through. Bounded.
 */

export const MAX_IDS_PER_CONVERSATION = 500
export const MAX_ID_CONVERSATIONS = 1000

export interface InboundIds {
  /** Records the id and returns true the first time; returns false when it was already recorded. */
  firstSeen(conversationKey: string, messageId: string): boolean
  /** Removes the id again (used when the turn failed, so a retry is processed). */
  forget(conversationKey: string, messageId: string): void
}

export function createInboundIds(): InboundIds {
  const conversations = new Map<string, Set<string>>()
  return {
    firstSeen(conversationKey, messageId) {
      const ids = conversations.get(conversationKey) ?? new Set<string>()
      // Delete + set moves an existing key to the newest position.
      conversations.delete(conversationKey)
      conversations.set(conversationKey, ids)
      if (ids.has(messageId)) return false
      ids.add(messageId)
      while (ids.size > MAX_IDS_PER_CONVERSATION) {
        ids.delete(ids.values().next().value as string)
      }
      while (conversations.size > MAX_ID_CONVERSATIONS) {
        conversations.delete(conversations.keys().next().value as string)
      }
      return true
    },
    forget(conversationKey, messageId) {
      const ids = conversations.get(conversationKey)
      if (!ids) return
      ids.delete(messageId)
      if (ids.size === 0) conversations.delete(conversationKey)
    },
  }
}
