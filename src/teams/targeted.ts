import type { ConversationReference, TurnContext } from 'botbuilder'
import { createHttpHeaders, createPipelineRequest } from '@azure/core-rest-pipeline'
import { isPersonalReference, type ConversationRefStore } from './conversationRefs.js'

/** Bot Connector URL that creates a message only its recipient sees (Teams targeted message). */
export function targetedActivityUrl(serviceUrl: string, conversationId: string): string {
  return (
    serviceUrl.replace(/\/+$/, '') +
    '/v3/conversations/' +
    encodeURIComponent(conversationId) +
    '/activities?isTargetedActivity=true'
  )
}

/** The activity body: the message goes into ref.conversation, visible only to ref.user. */
export function buildTargetedActivity(
  ref: Partial<ConversationReference>,
  text: string,
): Record<string, unknown> {
  if (!ref.user?.id) throw new Error('targeted message needs a user id')
  if (!ref.conversation?.id) throw new Error('targeted message needs a conversation id')
  return {
    type: 'message',
    text,
    channelId: ref.channelId,
    serviceUrl: ref.serviceUrl,
    from: ref.bot,
    conversation: ref.conversation,
    recipient: { ...ref.user, isTargeted: true },
  }
}

export type TargetedSender = (ref: Partial<ConversationReference>, text: string) => Promise<void>

export interface ContinueAdapter {
  continueConversationAsync(
    appId: string,
    ref: Partial<ConversationReference>,
    logic: (ctx: TurnContext) => Promise<void>,
  ): Promise<void>
}

/** Send a targeted message through the turn's connector client, so the request is signed with the bot's credentials. */
export function createTargetedSender(adapter: ContinueAdapter, appId: string): TargetedSender {
  return async (ref, text) => {
    const body = buildTargetedActivity(ref, text)
    await adapter.continueConversationAsync(appId, ref, async ctx => {
      const client = ctx.turnState.get((ctx.adapter as any).ConnectorClientKey) as
        | { sendRequest(req: unknown): Promise<{ status: number }> }
        | undefined
      if (!client) throw new Error('no connector client for this turn')
      const serviceUrl = String(ref.serviceUrl ?? '')
      if (!serviceUrl) throw new Error('no service URL')
      const res = await client.sendRequest(
        createPipelineRequest({
          url: targetedActivityUrl(serviceUrl, String(ref.conversation!.id)),
          method: 'POST',
          headers: createHttpHeaders({ 'Content-Type': 'application/json' }),
          body: JSON.stringify(body),
        }),
      )
      if (res.status < 200 || res.status > 299) {
        throw new Error(`targeted send failed (HTTP ${res.status})`)
      }
    })
  }
}

export const TARGETED_PROMPT_HINT = ' In a channel or group chat, @mention me in your reply.'

export interface PromptDelivery {
  refs: ConversationRefStore
  isAllowed: (aadObjectId: string) => boolean
  /** Plain message into a conversation (everyone in it sees it). */
  sendPlain: (ref: Partial<ConversationReference>, text: string) => Promise<void>
  /** Undefined when targeted messages are off. */
  sendTargeted?: TargetedSender
}

/** The conversation the permission prompt goes to, or undefined. */
export function selectPromptConversation(d: PromptDelivery): string | undefined {
  const id = d.sendTargeted
    ? d.refs.mostRecentConversationId()
    : d.refs.mostRecentPersonalConversationId()
  if (!id) return undefined
  const stored = d.refs.get(id)
  if (!stored) return undefined
  // An operator revoked since their last message is not prompted.
  if (!d.isAllowed(stored.aadObjectId)) return undefined
  return id
}

/** Deliver the prompt. A shared conversation never gets a plain message. */
export async function deliverPrompt(
  d: PromptDelivery,
  conversationId: string,
  text: string,
): Promise<void> {
  const stored = d.refs.get(conversationId)
  if (!stored) throw new Error(`no conversation reference for ${conversationId}`)
  if (isPersonalReference(stored.ref)) {
    await d.sendPlain(stored.ref, text)
    return
  }
  if (!d.sendTargeted) {
    throw new Error('refusing to post the permission prompt into a shared conversation')
  }
  try {
    await d.sendTargeted(stored.ref, text + TARGETED_PROMPT_HINT)
    return
  } catch (err) {
    process.stderr.write(
      `teams channel: targeted permission prompt failed (${err}); falling back to the 1:1 chat\n`,
    )
  }
  const fid = d.refs.mostRecentPersonalConversationId()
  const fb = fid ? d.refs.get(fid) : undefined
  if (!fb || !d.isAllowed(fb.aadObjectId)) {
    throw new Error('targeted permission prompt failed and there is no 1:1 chat to fall back to')
  }
  await d.sendPlain(fb.ref, text)
}
