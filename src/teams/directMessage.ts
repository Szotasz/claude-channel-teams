/**
 * Outbound 1:1 path — the `dm_member` tool implementation.
 *
 * Unlike `reply`, this addresses a PERSON, not a conversation: Claude passes an
 * Entra object id (the `aad_object_id` from `list_access` or an inbound
 * message) and the plugin works out the chat. That is what makes it proactive:
 * the recipient may never have written to the bot in a 1:1 chat, so there may
 * be no stored conversation reference for them yet.
 *
 * The gate is the same allowlist as everywhere else — the id has to be on it
 * or nothing is sent. When a personal reference already exists we simply reply
 * into it (which re-checks the allowlist at send time); otherwise we ask the
 * Bot Connector to open a fresh 1:1 conversation on the Teams service URL we
 * have seen on any inbound activity.
 *
 * Opening a chat only works when the app is installed for that user in
 * personal scope. Installing it proactively would need
 * TeamsAppInstallation.ReadWriteSelfForUser.All with admin consent and a
 * catalog-published app, which is out of scope by design — a 403 is translated
 * into an error that tells Claude what to ask the user for instead.
 */

import { type CloudAdapter, type ConversationParameters, type ConversationReference } from 'botbuilder'
import type { Allowlist } from '../pairing/allowlist.js'
import type { Config } from '../config.js'
import type { ConversationRefStore } from './conversationRefs.js'
import type { createReplySender } from './reply.js'
import { withSendRetry } from './sendRetry.js'

export class RecipientNotAllowedError extends Error {
  constructor(id: string) {
    super(`aad_object_id ${id} is not on the allowlist`)
    this.name = 'RecipientNotAllowedError'
  }
}

export class NoServiceUrlError extends Error {
  constructor() {
    super(
      'no Teams service URL is known yet: the bot must receive at least one message ' +
        '(from anyone, in any conversation) before it can open a 1:1 chat',
    )
    this.name = 'NoServiceUrlError'
  }
}

export class PersonalInstallMissingError extends Error {
  constructor(cause: unknown) {
    super(
      'Teams refused to open a 1:1 chat with this user (403): the app is most likely not ' +
        'installed for them in personal scope. Ask them to open the bot in Teams once.',
      { cause },
    )
    this.name = 'PersonalInstallMissingError'
  }
}

export class BotBlockedError extends Error {
  constructor(cause: unknown) {
    super(
      'Teams refused the 1:1 chat (403 MessageWritesBlocked): the user blocked or removed the bot.',
      { cause },
    )
    this.name = 'BotBlockedError'
  }
}

/**
 * Does a 403 payload mention MessageWritesBlocked anywhere we can read it?
 * The subcode hides in nested error bodies, so stringify each carrier; a
 * carrier whose getter throws just does not contribute.
 */
function mentionsMessageWritesBlocked(err: any): boolean {
  const parts: string[] = [String(err?.message ?? '')]
  for (const carrier of [err?.details, err?.body, err?.response?.bodyAsText]) {
    if (carrier === undefined || carrier === null) continue
    try {
      parts.push(JSON.stringify(carrier))
    } catch {}
  }
  return parts.some(p => p.includes('MessageWritesBlocked'))
}

export interface DirectMessageDeps {
  config: Pick<Config, 'appId' | 'tenantId'>
  adapter: Pick<CloudAdapter, 'createConversationAsync'>
  allowlist: Pick<Allowlist, 'isAllowed'>
  refs: ConversationRefStore
  /** Reused for an existing 1:1 chat, so its send-time allowlist re-check applies there too. */
  replier: Pick<ReturnType<typeof createReplySender>, 'sendReply'>
}

export function createDirectMessenger(deps: DirectMessageDeps) {
  async function dmMember(
    aadObjectId: string,
    text: string,
  ): Promise<{ conversationId: string; messageId?: string; created: boolean }> {
    // Same normalization as the allowlist itself, so whitespace or an
    // upper-case GUID cannot slip past the gate.
    const id = aadObjectId.trim().toLowerCase()
    if (!deps.allowlist.isAllowed(id)) throw new RecipientNotAllowedError(id)

    // An existing 1:1 chat is always preferable to a new one: the connector
    // only allows one personal conversation per bot/user pair anyway.
    const existing = deps.refs.personalConversationIdFor(id)
    if (existing !== undefined) {
      const messageId = await deps.replier.sendReply(existing, text)
      process.stderr.write(
        `teams channel: dm sent conv=${existing} created=false len=${text.length}\n`,
      )
      return { conversationId: existing, messageId, created: false }
    }

    const serviceUrl = deps.refs.serviceUrlHint()
    if (serviceUrl === undefined) throw new NoServiceUrlError()

    let newId: string | undefined
    let messageId: string | undefined
    let callbackStarted = false
    try {
      await deps.adapter.createConversationAsync(
        deps.config.appId,
        'msteams',
        serviceUrl,
        // null/undefined audience = the connector's default Bot Framework scope
        undefined as unknown as string,
        {
          isGroup: false,
          bot: { id: deps.config.appId },
          // Teams accepts the Entra object id here, personal scope only
          members: [{ id }],
          tenantId: deps.config.tenantId,
          channelData: { tenant: { id: deps.config.tenantId } },
        } as ConversationParameters,
        async ctx => {
          callbackStarted = true
          const convId = ctx.activity.conversation?.id
          if (typeof convId !== 'string' || convId.length === 0) {
            throw new Error('dm_member: Teams returned no conversation id for the new 1:1 chat')
          }
          newId = convId
          // Like storedReference(): only the fields needed to send, no display
          // names. The schema types mark `name` as required; the protocol does
          // not, hence the assertion.
          const ref = {
            channelId: 'msteams',
            serviceUrl,
            bot: { id: deps.config.appId },
            conversation: {
              id: convId,
              conversationType: 'personal',
              isGroup: false,
              tenantId: deps.config.tenantId,
            },
            user: { id, aadObjectId: id },
          } as Partial<ConversationReference>
          // Store the route BEFORE sending: a failed first send must not leave
          // the chat unaddressable, or the next dm_member would try to open a
          // second one (and `reply` could never reach the user either).
          // botOpened: the user was not active in this chat. Without the mark it would be the most recent 1:1 chat,
          // and the permission relay would send the next tool approval to this user. The recency stamp stays
          // normal, so a full store does not evict the chat it just opened.
          deps.refs.put(convId, ref, id, { botOpened: true })
          const res = await withSendRetry('dm', () =>
            ctx.sendActivity({ type: 'message', text, textFormat: 'markdown' }),
          )
          messageId = typeof res?.id === 'string' && res.id.length > 0 ? res.id : undefined
        },
      )
    } catch (err) {
      // Only a 403 from OPENING the chat is translated: Teams refused the chat itself, and user-side state
      // (blocked bot, missing install) does not clear on a retry. botbuilder runs the callback only after
      // createConversation succeeded, so an error from inside it (the first send) passes through unchanged.
      if (!callbackStarted && (err as { statusCode?: unknown })?.statusCode === 403) {
        if (mentionsMessageWritesBlocked(err)) throw new BotBlockedError(err)
        throw new PersonalInstallMissingError(err)
      }
      throw err
    }
    if (newId === undefined) {
      // Unreachable when the callback above ran to completion; guards the type
      // system's blind spot (assignments inside a closure are not tracked).
      throw new Error('dm_member: no conversation id returned for the new 1:1 chat')
    }
    process.stderr.write(
      `teams channel: dm sent conv=${newId} created=true len=${text.length}\n`,
    )
    return { conversationId: newId, messageId, created: true }
  }

  return { dmMember }
}
