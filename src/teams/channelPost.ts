import { TeamsInfo, type CloudAdapter, type ConversationReference, type TurnContext, type Activity } from 'botbuilder'
import type { Allowlist } from '../pairing/allowlist.js'
import type { Config } from '../config.js'
import type { ConversationRefStore } from './conversationRefs.js'
import type { ChannelRefStore } from './channelRefs.js'
import { ConversationNotAllowedError } from './reply.js'

export class ChannelNotEnabledError extends Error {
  constructor(id: string) {
    super(`channel ${id} is not enabled for post_to_channel (TEAMS_PROACTIVE_CHANNELS)`)
    this.name = 'ChannelNotEnabledError'
  }
}

export class UnknownChannelError extends Error {
  constructor(id: string) {
    super(`no reference for channel ${id} yet: an allowlisted user must send one message in that channel after it was enabled`)
    this.name = 'UnknownChannelError'
  }
}

export type ChannelSender = (
  ctx: TurnContext,
  activity: Partial<Activity>,
  channelId: string,
  appId: string,
) => Promise<[Partial<ConversationReference>, string]>

export interface ChannelPostDeps {
  config: Config
  adapter: Pick<CloudAdapter, 'continueConversationAsync'>
  allowlist: Allowlist
  refs: ConversationRefStore
  channelRefs: ChannelRefStore
  /** Injected for tests; defaults to TeamsInfo.sendMessageToTeamsChannel. */
  sendToChannel?: ChannelSender
}

export function createChannelPoster(deps: ChannelPostDeps) {
  async function postToChannel(
    channelId: string,
    text: string,
  ): Promise<{ conversationId: string; messageId?: string }> {
    if (!deps.config.proactiveChannels?.has(channelId)) throw new ChannelNotEnabledError(channelId)
    const stored = deps.channelRefs.get(channelId)
    if (!stored) throw new UnknownChannelError(channelId)
    if (!deps.allowlist.isAllowed(stored.aadObjectId)) throw new ConversationNotAllowedError(channelId)

    const send: ChannelSender =
      deps.sendToChannel ??
      ((ctx, activity, id, appId) => TeamsInfo.sendMessageToTeamsChannel(ctx, activity as Activity, id, appId))
    let result: [Partial<ConversationReference>, string] | undefined
    await deps.adapter.continueConversationAsync(deps.config.appId, stored.ref, async ctx => {
      result = await send(ctx, { type: 'message', text, textFormat: 'markdown' }, channelId, deps.config.appId)
    })

    const newRef = result?.[0]
    const conversationId = newRef?.conversation?.id
    if (!conversationId) {
      throw new Error(`post_to_channel: Teams returned no conversation for channel ${channelId}`)
    }
    // The root post's activity id is the ';messageid=' part of the new thread id.
    // The second value TeamsInfo returns is NOT it: through CloudAdapter it is the
    // conversation id again (createCreateActivity sets activity.id to it). Without
    // the suffix we return no message id rather than a wrong one -- and do not
    // throw, because the post has already been made.
    const messageId = /;messageid=([^;]+)/.exec(conversationId)?.[1]
    deps.refs.put(conversationId, newRef as ConversationReference, stored.aadObjectId)
    process.stderr.write(
      `teams channel: channel post sent channel=${channelId} conv=${conversationId} len=${text.length}\n`,
    )
    return messageId ? { conversationId, messageId } : { conversationId }
  }
  return { postToChannel }
}
