/**
 * In-memory record of which messages the plugin delivered to Claude, and who
 * sent them. Edit / delete / restore events are only forwarded for messages
 * found here. Bounded; the oldest entries are dropped first.
 */

export const MAX_MESSAGES_PER_CONVERSATION = 500
export const MAX_CONVERSATIONS = 1000

export interface DeliveredMessages {
  remember(conversationKey: string, messageId: string, aadObjectId: string): void
  get(conversationKey: string, messageId: string): { aadObjectId: string } | undefined
}

export function createDeliveredMessages(): DeliveredMessages {
  const conversations = new Map<string, Map<string, string>>()
  return {
    remember(conversationKey, messageId, aadObjectId) {
      let msgs = conversations.get(conversationKey)
      // Delete + set moves an existing key to the newest position.
      conversations.delete(conversationKey)
      if (!msgs) msgs = new Map()
      conversations.set(conversationKey, msgs)
      msgs.delete(messageId)
      msgs.set(messageId, aadObjectId)
      while (msgs.size > MAX_MESSAGES_PER_CONVERSATION) {
        msgs.delete(msgs.keys().next().value as string)
      }
      while (conversations.size > MAX_CONVERSATIONS) {
        conversations.delete(conversations.keys().next().value as string)
      }
    },
    get(conversationKey, messageId) {
      const aadObjectId = conversations.get(conversationKey)?.get(messageId)
      return aadObjectId === undefined ? undefined : { aadObjectId }
    },
  }
}
