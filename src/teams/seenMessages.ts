/** Message ids seen per conversation, so a quote can only point inside the conversation it is sent to. In memory only:
 * after a restart an older message can no longer be quoted. Bounded: at most `perConversation` ids per conversation
 * (oldest dropped first) and at most `maxConversations` conversations (least recently used dropped first). */
export interface SeenMessages { remember(conversationId: string, messageId: string): void; has(conversationId: string, messageId: string): boolean }

export function createSeenMessages(perConversation = 500, maxConversations = 1000): SeenMessages {
  const conversations = new Map<string, Map<string, true>>()
  return {
    remember(conversationId, messageId) {
      if (conversationId === '' || messageId === '') return
      const ids = conversations.get(conversationId) ?? new Map<string, true>()
      conversations.delete(conversationId)
      conversations.set(conversationId, ids)
      ids.delete(messageId)
      ids.set(messageId, true)
      while (ids.size > perConversation) ids.delete(ids.keys().next().value as string)
      while (conversations.size > maxConversations) conversations.delete(conversations.keys().next().value as string)
    },
    has(conversationId, messageId) {
      return conversations.get(conversationId)?.has(messageId) ?? false
    },
  }
}
