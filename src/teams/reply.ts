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
import type { ConversationRefStore } from './conversationRefs.js'
import type { TypingPump } from './typingPump.js'
import { ADAPTIVE_CARD_CONTENT_TYPE, validateAdaptiveCard } from './cards.js'

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

  async function sendReply(
    conversationId: string,
    text: string,
    card?: unknown,
    opts?: { aiGenerated?: boolean; citations?: CitationInput[] },
  ): Promise<string | undefined> {
    if (card !== undefined) {
      const checked = validateAdaptiveCard(card)
      if (!checked.ok) throw new Error(`invalid card: ${checked.error}`)
    }
    const entity = buildMessageEntity(opts ?? {})
    const { ref } = assertAllowedConversation(conversationId)
    // Stop the typing pump first. We do this before the network send so the
    // indicator clears even if the connector call throws — flapping
    // "is typing…" forever after a failure is worse than the failure itself.
    deps.typingPump?.stop(conversationId)
    // `continueConversationAsync` rehydrates the conversation context using
    // the stored reference and our pinned credentials. The SDK takes care of
    // grabbing a fresh service-to-service token (cached by app id).
    let sentId: string | undefined
    await deps.adapter.continueConversationAsync(
      deps.config.appId,
      ref,
      async turnContext => {
        const res = await turnContext.sendActivity({
          type: 'message',
          text,
          textFormat: 'markdown',
          ...(card !== undefined
            ? { attachments: [{ contentType: ADAPTIVE_CARD_CONTENT_TYPE, content: card }] }
            : {}),
          ...(entity ? { entities: [entity as { type: string }] } : {}),
        })
        sentId = typeof res?.id === 'string' && res.id.length > 0 ? res.id : undefined
      },
    )
    process.stderr.write(
      `teams channel: reply sent conv=${conversationId} len=${text.length}${card !== undefined ? ' card=yes' : ''}${opts?.aiGenerated === true ? ' ai=yes' : ''}${opts?.citations?.length ? ` cites=${opts.citations.length}` : ''}\n`,
    )
    return sentId
  }

  /**
   * Edit a message the bot previously sent. Mirrors sendReply's gate exactly
   * (re-check the allowlist on the way out), then updates the activity in place
   * via the same conversation reference. Used for interim progress updates.
   */
  async function editReply(
    conversationId: string,
    messageId: string,
    text: string,
  ): Promise<void> {
    const { ref } = assertAllowedConversation(conversationId)
    await deps.adapter.continueConversationAsync(
      deps.config.appId,
      ref,
      async turnContext => {
        await turnContext.updateActivity({ type: 'message', id: messageId, text, textFormat: 'markdown' })
      },
    )
    process.stderr.write(
      `teams channel: message edited conv=${conversationId} id=${messageId}\n`,
    )
  }

  return { sendReply, editReply, assertAllowedConversation }
}
