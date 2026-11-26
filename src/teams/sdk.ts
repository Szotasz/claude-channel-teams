/**
 * The one place the plugin takes its Bot Framework protocol types and helpers from (TEAMSSDK1126).
 *
 * The plugin ran on botbuilder 4.x, whose repository Microsoft archived (support tickets end 2025-12-31). It now runs on
 * the Microsoft 365 Agents SDK (@microsoft/agents-hosting, -activity, -hosting-extensions-teams), which keeps the same
 * shapes: CloudAdapter, TurnContext, TeamsInfo, ConversationReference.
 *
 * Two botbuilder statics have no static form there (TurnContext.getConversationReference and
 * TurnContext.removeRecipientMention exist only as methods of the Activity class). The turn handler and the tests pass
 * plain activity objects, so both are kept here as plain functions with botbuilder's exact behaviour.
 */

import type { Activity, ConversationReference } from '@microsoft/agents-activity'

export type { Activity, Attachment, ConversationParameters, ConversationReference } from '@microsoft/agents-activity'
export { CloudAdapter, TurnContext } from '@microsoft/agents-hosting'
export { TeamsInfo } from '@microsoft/agents-hosting-extensions-teams'

/**
 * The plugin builds outgoing activities as plain objects. The Agents SDK types its send/update methods with the
 * Activity class, but at run time it turns any plain object into one itself (TurnContext.sendActivities and
 * updateActivity call Activity.fromObject, which validates the shape). This is the one cast at that boundary.
 */
export function asActivity(a: Partial<Activity> | Record<string, unknown>): Activity {
  return a as Activity
}

/**
 * A stored reference (conversationRefs.ts keeps only the fields needed to send) as the SDK's ConversationReference.
 * continueConversation itself refuses one without a service URL or a conversation id.
 */
export function asReference(ref: Partial<ConversationReference>): ConversationReference {
  return ref as ConversationReference
}

/**
 * A raw Bot Connector request (a reaction, a targeted message) through the turn's connector client, so it carries the
 * bot's token like every other send. Returns the HTTP status: the Agents SDK's HttpClient throws on a non-2xx answer
 * (an HttpError with `status`), botframework-connector's sendRequest returned it, and the callers compare it.
 */
export async function connectorRequest(
  ctx: { adapter: unknown; turnState: { get(key: unknown): unknown } },
  req: { method: string; url: string; data?: unknown },
): Promise<number> {
  const key = (ctx.adapter as { ConnectorClientKey?: symbol } | undefined)?.ConnectorClientKey
  const client = key ? (ctx.turnState.get(key) as { httpClient?: { baseURL?: string; request(c: unknown): Promise<{ status: number }> } } | undefined) : undefined
  if (!client?.httpClient) throw new Error('no connector client for this turn')
  // The SDK's HttpClient puts the bot's Bearer token on an absolute URL too, whatever its host. Only the origin the
  // client was built for (the turn's service URL) may get it (Dani's review on #60, S1).
  let sameOrigin = false
  try {
    sameOrigin = typeof client.httpClient.baseURL === 'string' && new URL(req.url).origin === new URL(client.httpClient.baseURL).origin
  } catch {
    sameOrigin = false
  }
  if (!sameOrigin) throw new Error('connector request refused: not the origin of the turn\'s service URL')
  try {
    const res = await client.httpClient.request({
      method: req.method,
      url: req.url,
      ...(req.data !== undefined ? { data: req.data, headers: { 'Content-Type': 'application/json' } } : {}),
    })
    return res.status
  } catch (err) {
    const status = (err as { status?: unknown }).status
    if (typeof status === 'number') return status
    throw err
  }
}

type ActivityLike = Pick<Activity, 'type' | 'id' | 'from' | 'recipient' | 'conversation' | 'channelId' | 'locale' | 'serviceUrl'>

const shallowCopy = <T>(v: T): T => (v && typeof v === 'object' ? ({ ...(v as object) } as T) : v)

/**
 * The reference a later proactive send needs, from an incoming activity: botbuilder's
 * TurnContext.getConversationReference (botbuilder-core 4.23.3). The activity id is left out only for a
 * conversationUpdate from Direct Line or Web Chat, as there.
 */
export function getConversationReference(activity: ActivityLike): ConversationReference {
  const ch = String(activity.channelId ?? '')
  const replyTo = activity.type !== 'conversationUpdate' || (ch !== 'directline' && ch !== 'webchat') ? activity.id : undefined
  return {
    activityId: replyTo,
    user: shallowCopy(activity.from),
    agent: shallowCopy(activity.recipient),
    conversation: shallowCopy(activity.conversation),
    channelId: activity.channelId,
    locale: activity.locale,
    serviceUrl: activity.serviceUrl,
  } as ConversationReference
}

/**
 * Removes the first at-mention of the bot itself from the activity's text and returns the text: botbuilder's
 * TurnContext.removeRecipientMention (botbuilder-core 4.23.3). It changes activity.text, as there.
 */
export function removeRecipientMention(activity: { text?: string; entities?: unknown[]; recipient?: { id?: string } }): string | undefined {
  const id = activity.recipient?.id
  const mention = (activity.entities ?? []).find((e): e is { type: string; text?: string; mentioned?: { id?: string } } =>
    !!e && typeof e === 'object' && typeof (e as { type?: unknown }).type === 'string'
    && (e as { type: string }).type.toLowerCase() === 'mention'
    && (e as { mentioned?: { id?: unknown } }).mentioned?.id === id)
  if (mention && typeof activity.text === 'string') activity.text = activity.text.replace(mention.text ?? '', '').trim()
  return activity.text
}
