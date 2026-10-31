/**
 * Outbound reply path — the `reply` tool implementation.
 *
 * Phase 2: text replies rendered as markdown (Teams' supported subset:
 * bold, italic, headers, bullet/numbered lists, links, inline code, code
 * blocks, blockquotes — tables and HTML are not reliably supported).
 * Adaptive Cards, attachments, and the `react` / `edit_message` tools are
 * out of scope (design decisions F.4 / Phase 5).
 *
 * The outbound side mirrors the inbound gate (`assertAllowedConversation`).
 * A compromised or prompt-injected Claude calling `reply` with a
 * `conversation_id` it pulled from elsewhere must be refused — the conversation
 * has to have been seeded by an allowlisted inbound activity at receive time,
 * AND the original sender must still be on the allowlist when the reply fires.
 */

import { type CloudAdapter, type ConversationReference } from 'botbuilder'
import type { Allowlist } from '../pairing/allowlist.js'
import type { Config } from '../config.js'
import { isPersonalReference, type ConversationRefStore } from './conversationRefs.js'
import type { TypingPump } from './typingPump.js'
import type { AckPlaceholders } from './ackPlaceholder.js'
import { withSendRetry } from './sendRetry.js'
import { isStreamClosedError, type StatusStreams } from './statusStream.js'
import { ADAPTIVE_CARD_CONTENT_TYPE, validateAdaptiveCard } from './cards.js'
import { createPipelineRequest } from '@azure/core-rest-pipeline'
import { USER_VIEW_VERB, usesReservedVerb, validateUserCards, type UserCard } from './userViews.js'
import { buildTargetedActivity, postTargetedActivity } from './targeted.js'
import type { TargetedInbound } from './targetedInbound.js'

export const REACTION_ID_RE = /^[a-z0-9_]{1,40}$/

export const LATEST_PRIVATE_MESSAGE =
  'the latest message in this conversation was private (targeted): answer it with targeted_message_id. ' +
  'Use public: true only for a message meant for everyone in the conversation.'

/**
 * The reply tool's targeted_message_id: absent (undefined or null) or a non-empty string. Anything else is
 * refused, so a private answer with a malformed id can never fall through to a public post.
 */
export function parseTargetedMessageId(raw: unknown): string | undefined {
  if (raw === undefined || raw === null) return undefined
  if (typeof raw !== 'string' || raw.length === 0) {
    throw new Error('targeted_message_id must be the message_id string of the inbound tag')
  }
  return raw
}

export type CitationInput = {
  name: string
  abstract: string
  url?: string
  keywords?: string[]
  position?: number
}

const MAX_CITATIONS = 20
const CITATION_KEYS = new Set(['name', 'abstract', 'url', 'keywords', 'position'])
const cpLen = (s: string) => [...s].length

function checkText(v: unknown, label: string, max: number): string {
  if (typeof v !== 'string' || v.length === 0) {
    throw new Error(`${label} must be a non-empty string`)
  }
  if (cpLen(v) > max) throw new Error(`${label} must be at most ${max} characters`)
  return v
}

/**
 * Build the single root-level message entity carrying the "AI generated"
 * label and/or citations. Returns undefined when neither is requested.
 */
export function buildMessageEntity(opts: {
  aiGenerated?: boolean
  citations?: CitationInput[]
}): Record<string, unknown> | undefined {
  const ai = opts.aiGenerated === true
  const citations = opts.citations ?? []
  if (!Array.isArray(citations)) throw new Error('citations must be an array')
  if (!ai && citations.length === 0) return undefined
  if (citations.length > MAX_CITATIONS) {
    throw new Error(`too many citations: ${citations.length} (at most ${MAX_CITATIONS})`)
  }
  const seen = new Set<number>()
  const claims = citations.map((c, i) => {
    const label = `citation ${i}`
    if (typeof c !== 'object' || c === null || Array.isArray(c)) {
      throw new Error(`${label} must be an object`)
    }
    for (const k of Object.keys(c)) {
      if (!CITATION_KEYS.has(k)) throw new Error(`${label} has unknown key '${k}'`)
    }
    const name = checkText(c.name, `${label} name`, 80)
    const abstract = checkText(c.abstract, `${label} abstract`, 160)
    // null in an optional field means absent, as for the top-level card (some models send null for it).
    let position = i + 1
    if (c.position !== undefined && c.position !== null) {
      if (typeof c.position !== 'number' || !Number.isInteger(c.position) || c.position < 1) {
        throw new Error(`${label} position must be an integer >= 1`)
      }
      position = c.position
    }
    if (seen.has(position)) throw new Error(`${label} has duplicate position ${position}`)
    seen.add(position)
    const appearance: Record<string, unknown> = { '@type': 'DigitalDocument', name, abstract }
    if (c.url !== undefined && c.url !== null) {
      let href: string | undefined
      if (typeof c.url === 'string') {
        try {
          const parsed = new URL(c.url)
          if (parsed.protocol === 'https:' || parsed.protocol === 'http:') href = parsed.href
        } catch {}
      }
      if (href === undefined) throw new Error(`${label} url must be a valid http(s) URL`)
      // Send what was checked: the parser drops stray spaces, tabs and newlines that the raw string still has.
      appearance.url = href
    }
    if (c.keywords !== undefined && c.keywords !== null) {
      if (!Array.isArray(c.keywords) || c.keywords.length > 3) {
        throw new Error(`${label} keywords must be an array of at most 3 strings`)
      }
      c.keywords.forEach((k, j) => checkText(k, `${label} keyword ${j}`, 28))
      appearance.keywords = [...c.keywords]
    }
    return { '@type': 'Claim', position, appearance }
  })
  return {
    type: 'https://schema.org/Message',
    '@type': 'Message',
    '@context': 'https://schema.org',
    '@id': '',
    ...(ai ? { additionalType: ['AIGeneratedContent'] } : {}),
    ...(claims.length > 0 ? { citation: claims } : {}),
  }
}

export class UnknownConversationError extends Error {
  constructor(conversationId: string) {
    super(
      `unknown conversation_id: ${conversationId}. ` +
        `The plugin doesn't have a conversation reference for this thread — ` +
        `likely after a plugin restart. Ask the user to send a fresh Teams message to re-seed.`,
    )
    this.name = 'UnknownConversationError'
  }
}

export class ConversationNotAllowedError extends Error {
  constructor(conversationId: string) {
    super(
      `conversation_id ${conversationId} is not allowlisted — refusing to send. ` +
        `The originating sender must be on the allowlist at the time of reply.`,
    )
    this.name = 'ConversationNotAllowedError'
  }
}

export interface ReplyDeps {
  config: Config
  adapter: CloudAdapter
  allowlist: Allowlist
  refs: ConversationRefStore
  /** Optional typing-indicator pump — stopped when a reply lands. */
  typingPump?: TypingPump
  /** Open status streams (TEAMS_STATUS_UPDATES). Optional so existing tests still wire. */
  statusStreams?: StatusStreams
  /** Optional acknowledgement store: a reply overwrites a pending acknowledgement. */
  acks?: AckPlaceholders
  /** Person -> user MRI, learned from inbound activities (user-specific views). */
  knownUsers?: { get(aadObjectId: string): string | undefined }
  /** Stored user-specific views (user_cards). */
  userViews?: { create(baseCard: Record<string, unknown>, userCards: UserCard[], userMris: string[]): { id: string; base: Record<string, unknown> } }
  /** Memory of inbound targeted messages and targeted messages sent. */
  targetedInbound?: TargetedInbound
  /** Clock, for tests. */
  now?: () => number
}

export interface ReplyOptions {
  /** Show the 'AI generated' label. */
  aiGenerated?: boolean
  /** Numbered source citations. */
  citations?: CitationInput[]
  /** message_id of an inbound targeted message: the reply is then shown only to its sender. */
  targetedMessageId?: string
  /** Post publicly even though the latest message in the conversation was targeted. */
  public?: boolean
  /** Per-user views of the card (user_cards), checked in sendReply. */
  userCards?: unknown
}

export function createReplySender(deps: ReplyDeps) {
  /**
   * The outbound gate.
   *
   * A conversation reference is only stored after an inbound activity passes
   * the allowlist (see adapter.ts). But the allowlist can be mutated between
   * receive and reply — an operator can remove an ID while a Claude turn is
   * in flight. We re-check on the way out so a freshly-revoked user can't
   * receive a trailing reply.
   *
   * This is mirrored on Telegram as `assertAllowedChat`.
   */
  function assertAllowedConversation(conversationId: string): {
    ref: Partial<ConversationReference>
  } {
    const stored = deps.refs.get(conversationId)
    if (!stored) throw new UnknownConversationError(conversationId)
    // Re-check the allowlist at send time. Defends against revocations.
    if (!deps.allowlist.isAllowed(stored.aadObjectId)) {
      throw new ConversationNotAllowedError(conversationId)
    }
    return { ref: stored.ref }
  }

  const nowMs = () => (deps.now ?? Date.now)()

  /** Refuse a public post right after a private (targeted) inbound unless explicitly asked for. */
  function assertPublicAllowed(conversationId: string, isPublic: boolean | undefined): void {
    if (isPublic !== true && deps.targetedInbound?.latestIsTargeted(conversationId, nowMs())) {
      throw new Error(LATEST_PRIVATE_MESSAGE)
    }
  }

  async function sendTargetedReply(
    conversationId: string,
    text: string,
    card: unknown,
    targetedMessageId: string,
    entity?: Record<string, unknown>,
  ): Promise<string | undefined> {
    const entry = deps.targetedInbound?.get(conversationId, targetedMessageId, nowMs())
    if (!entry) throw new Error('unknown or expired targeted_message_id for this conversation')
    if (!deps.allowlist.isAllowed(entry.aadObjectId)) {
      throw new ConversationNotAllowedError(conversationId)
    }
    const body = {
      ...buildTargetedActivity(entry.ref, text),
      textFormat: 'markdown',
      entities: [{ type: 'targetedMessageInfo', messageId: targetedMessageId }, ...(entity ? [entity as { type: string }] : [])],
      ...(card !== undefined
        ? { attachments: [{ contentType: ADAPTIVE_CARD_CONTENT_TYPE, content: card }] }
        : {}),
    }
    deps.typingPump?.stop(conversationId)
    // No fallback: a failure here must never turn into a public post.
    const bodyText = await postTargetedActivity(deps.adapter, deps.config.appId, entry.ref, body)
    let sentId: string | undefined
    try {
      const id = (JSON.parse(bodyText ?? '') as { id?: unknown } | null)?.id
      if (typeof id === 'string' && id.length > 0) sentId = id
    } catch {
      // No JSON body: no id to report.
    }
    if (sentId) deps.targetedInbound?.rememberSent(sentId)
    process.stderr.write(
      `teams channel: targeted reply sent conv=${conversationId} len=${text.length}${card !== undefined ? ' card=yes' : ''}\n`,
    )
    return sentId
  }

  async function sendReply(
    conversationId: string,
    text: string,
    card?: unknown,
    opts: ReplyOptions = {},
  ): Promise<string | undefined> {
    if (opts.targetedMessageId !== undefined && opts.public === true) {
      throw new Error('targeted_message_id and public cannot be combined')
    }
    // The private path posts one activity to the sender only; per-user card views would be dropped silently.
    if (opts.targetedMessageId !== undefined && opts.userCards !== undefined) {
      throw new Error('targeted_message_id and user_cards cannot be combined')
    }
    if (card !== undefined) {
      const checked = validateAdaptiveCard(card)
      if (!checked.ok) throw new Error(`invalid card: ${checked.error}`)
      // A click on a button with this verb would be taken for a view refresh and never reach Claude.
      if (usesReservedVerb(card)) throw new Error(`card uses the reserved verb ${USER_VIEW_VERB}`)
    }
    let outCard = card
    let views = 0
    let storeViews: (() => Record<string, unknown>) | undefined
    if (opts?.userCards !== undefined) {
      const userCards = validateUserCards(opts.userCards, card)
      if (!deps.knownUsers || !deps.userViews) throw new Error('user_cards is not available in this server')
      const unknown = userCards.filter(u => !deps.knownUsers!.get(u.aadObjectId)).map(u => u.aadObjectId)
      if (unknown.length > 0) {
        throw new Error(
          `user_cards: no Teams user id is known yet for ${unknown.join(', ')} (they have not written to the bot since it started)`,
        )
      }
      const mris = userCards.map(u => deps.knownUsers!.get(u.aadObjectId)!)
      storeViews = () => deps.userViews!.create(card as Record<string, unknown>, userCards, mris).base
      views = userCards.length
    }
    const entity = buildMessageEntity(opts)
    if (opts.targetedMessageId !== undefined) {
      return sendTargetedReply(conversationId, text, card, opts.targetedMessageId, entity)
    }
    assertPublicAllowed(conversationId, opts.public)
    const { ref } = assertAllowedConversation(conversationId)
    // Stop the typing pump first. We do this before the network send so the
    // indicator clears even if the connector call throws — flapping
    // "is typing…" forever after a failure is worse than the failure itself.
    deps.typingPump?.stop(conversationId)
    const ackId = await deps.acks?.take(conversationId)
    // `continueConversationAsync` rehydrates the conversation context using
    // the stored reference and our pinned credentials. The SDK takes care of
    // grabbing a fresh service-to-service token (cached by app id).
    let sentId: string | undefined
    let closedStream = false
    const send = async (stream: { streamId: string } | undefined) => {
      // The reply may have waited in the status queue: check the allowlist again now.
      const { ref: current } = assertAllowedConversation(conversationId)
      // The view is stored only after every check has passed, right before the send.
      if (storeViews) outCard = storeViews()
      await deps.adapter.continueConversationAsync(
        deps.config.appId,
        current,
        async turnContext => {
          const activity = {
            type: 'message',
            text,
            textFormat: 'markdown',
            ...(outCard !== undefined
              ? { attachments: [{ contentType: ADAPTIVE_CARD_CONTENT_TYPE, content: outCard }] }
              : {}),
            ...(entity ? { entities: [entity as { type: string }] } : {}),
          }
          // Streaming is 1:1 only: close the stream only if the conversation is still 1:1 now.
          const closes = stream && isPersonalReference(current) ? stream : undefined
          const streamEntity = closes ? { type: 'streaminfo', streamId: closes.streamId, streamType: 'final' } : undefined
          // A reply that closes the status stream is that stream's final message: it cannot also overwrite the
          // acknowledgement, which then stays as sent.
          if (ackId && !streamEntity) {
            try {
              await withSendRetry('reply', () => turnContext.updateActivity({ ...activity, id: ackId }))
              sentId = ackId
              return
            } catch (err) {
              process.stderr.write(
                `teams channel: overwriting the acknowledgement failed (${err instanceof Error ? err.message : String(err)}); sending a new message\n`,
              )
            }
          }
          const closing = streamEntity
            ? { ...activity, entities: [...((activity as { entities?: unknown[] }).entities ?? []), streamEntity] }
            : activity
          const deliver = (a: unknown) => withSendRetry('reply', () => turnContext.sendActivity(a as any))
          let res
          try {
            res = await deliver(closing)
            closedStream = streamEntity !== undefined
          } catch (err) {
            if (!streamEntity || !isStreamClosedError(err)) throw err
            process.stderr.write(`teams channel: status stream already closed conv=${conversationId}; reply sent as a new message\n`)
            res = await deliver(activity)
          }
          sentId = typeof res?.id === 'string' && res.id.length > 0 ? res.id : undefined
        },
      )
    }
    // With status updates on, a 1:1 reply runs in the conversation's status queue: it closes the open stream.
    // Other conversations never have a stream, so their replies do not wait.
    if (deps.statusStreams && isPersonalReference(ref)) await deps.statusStreams.withStream(conversationId, send)
    else await send(undefined)
    process.stderr.write(
      `teams channel: reply sent conv=${conversationId} len=${text.length}${card !== undefined ? ' card=yes' : ''}${opts?.aiGenerated === true ? ' ai=yes' : ''}${opts?.citations?.length ? ` cites=${opts.citations.length}` : ''}${closedStream ? ' stream=closed' : ''}${views > 0 ? ` views=${views}` : ''}\n`,
    )
    return sentId
  }

  /**
   * Edit a message the bot previously sent. Mirrors sendReply's gate exactly
   * (re-check the allowlist on the way out), then updates the activity in place
   * via the same conversation reference. Used for interim progress updates.
   * An edit is a public post too: right after a private (targeted) inbound it
   * is refused unless isPublic, like sendReply.
   */
  async function editReply(
    conversationId: string,
    messageId: string,
    text: string,
    isPublic?: boolean,
  ): Promise<void> {
    if (deps.targetedInbound?.isSentTargeted(messageId)) {
      throw new Error('editing a targeted message is not supported')
    }
    assertPublicAllowed(conversationId, isPublic)
    const { ref } = assertAllowedConversation(conversationId)
    await deps.adapter.continueConversationAsync(
      deps.config.appId,
      ref,
      async turnContext => {
        await withSendRetry('edit', () => turnContext.updateActivity({ type: 'message', id: messageId, text, textFormat: 'markdown' }))
      },
    )
    process.stderr.write(
      `teams channel: message edited conv=${conversationId} id=${messageId}\n`,
    )
  }

  /**
   * Add (or remove) the bot's own emoji reaction on a message. Teams exposes
   * this as PUT/DELETE .../activities/{id}/reactions/{type} on the Bot
   * Connector; we send it through the turn's connector client so it is signed
   * with the bot's credentials.
   */
  async function react(
    conversationId: string,
    messageId: string,
    reaction: string,
    remove = false,
  ): Promise<void> {
    if (typeof reaction !== 'string' || !REACTION_ID_RE.test(reaction)) {
      throw new Error('invalid reaction id')
    }
    if (typeof messageId !== 'string' || messageId.length === 0 || messageId.length > 512) {
      throw new Error('invalid message id')
    }
    const { ref } = assertAllowedConversation(conversationId)
    await deps.adapter.continueConversationAsync(
      deps.config.appId,
      ref,
      async ctx => {
        const client = ctx.turnState.get((ctx.adapter as any).ConnectorClientKey) as
          | { sendRequest(req: unknown): Promise<{ status: number }> }
          | undefined
        if (!client) throw new Error('no connector client for this turn')
        const serviceUrl = String(ref.serviceUrl ?? '').replace(/\/+$/, '')
        if (!serviceUrl) throw new Error('no service URL')
        const url =
          `${serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}` +
          `/activities/${encodeURIComponent(messageId)}/reactions/${reaction}`
        const res = await client.sendRequest(
          createPipelineRequest({ url, method: remove ? 'DELETE' : 'PUT' }),
        )
        if (res.status < 200 || res.status >= 300) {
          throw new Error(`react failed (HTTP ${res.status})`)
        }
      },
    )
    process.stderr.write(
      `teams channel: reaction ${remove ? 'removed' : 'added'} conv=${conversationId} id=${messageId} type=${reaction}\n`,
    )
  }

  async function sendStatus(conversationId: string, text: string): Promise<void> {
    if (!deps.statusStreams) throw new Error('status updates are not enabled (set TEAMS_STATUS_UPDATES)')
    const { ref } = assertAllowedConversation(conversationId)
    if (!isPersonalReference(ref)) throw new Error('status updates work only in a 1:1 chat')
    await deps.statusStreams.update(conversationId, text, async activity => {
      // The call may have waited for its slot: check the allowlist again now.
      const { ref: current } = assertAllowedConversation(conversationId)
      if (!isPersonalReference(current)) throw new Error('status updates work only in a 1:1 chat')
      // Teams shows no typing animation while a stream is open.
      deps.typingPump?.stop(conversationId)
      let id: string | undefined
      await deps.adapter.continueConversationAsync(deps.config.appId, current, async turnContext => {
        const res = await turnContext.sendActivity(activity as any)
        id = typeof res?.id === 'string' && res.id.length > 0 ? res.id : undefined
      })
      return id
    })
    process.stderr.write(`teams channel: status update conv=${conversationId} len=${text.length}\n`)
  }

  /** Send the acknowledgement message and return its id (undefined when Teams gave none). */
  async function sendAck(conversationId: string, text: string): Promise<string | undefined> {
    const { ref } = assertAllowedConversation(conversationId)
    let id: string | undefined
    await deps.adapter.continueConversationAsync(
      deps.config.appId,
      ref,
      async turnContext => {
        const res = await withSendRetry('ack', () => turnContext.sendActivity({ type: 'message', text }))
        id = typeof res?.id === 'string' && res.id.length > 0 ? res.id : undefined
      },
    )
    return id
  }

  return { sendReply, editReply, react, assertAllowedConversation, assertPublicAllowed, sendStatus, sendAck }
}
