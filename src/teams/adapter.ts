/**
 * Bot Framework HTTP listener — the inbound side.
 *
 * Phase 2 implementation:
 *
 *   - Spins up a `Bun.serve` listener on TEAMS_PLUGIN_BIND_HOST:TEAMS_PLUGIN_PORT
 *     (loopback by default).
 *   - Routes POST /api/messages through CloudAdapter.process so the SDK's JWT
 *     validator runs before anything else. Failure → 401.
 *   - For each authenticated message activity:
 *       1. Pull `aadObjectId`, `conversation.id`, `text`, `from.name`,
 *          `conversation.tenantId`.
 *       2. Allowlist check — drop silently if not allowed (with one stderr
 *          line per unique rejected ID, matching the Telegram pattern).
 *       3. Tenant defence-in-depth: compare conversation.tenantId against the
 *          configured tenant. Drop mismatches.
 *       4. Capture the conversation reference so the outbound side can find
 *          its way back.
 *       5. Hand the payload to the channel-event sink.
 *   - Also exposes GET /health — unauthenticated, returns 200, lets the
 *     reverse proxy / Caddy / monitoring confirm reachability without an
 *     activity round-trip.
 *
 * We deliberately use `Bun.serve` instead of restify here. CloudAdapter's
 * `Request` and `Response` interfaces are minimal (see node_modules/botbuilder/
 * lib/interfaces/) — we adapt a Bun Request to that shape with a few-line shim.
 * Dropping restify saves two transitive trees and removes a dep we don't need.
 * Documented in docs/dependencies.md.
 */

import {
  CloudAdapter,
  TurnContext,
  type Activity,
  type ConversationReference,
} from 'botbuilder'
import type { Attachment } from 'botframework-schema'
import type { Config } from '../config.js'
import type { Allowlist } from '../pairing/allowlist.js'
import type { PendingStore } from '../pairing/pair.js'
import type { PermissionRelay } from '../permission/relay.js'
import { PERMISSION_REPLY_RE } from '../permission/relay.js'
import { isPersonalReference, type ConversationRefStore } from './conversationRefs.js'
import type { TypingPump } from './typingPump.js'
import type { Outbox } from './outbox.js'
import { handleFilesRequest } from './filesRoute.js'
import type { DeliveredMessages } from './deliveredMessages.js'

/**
 * Optional pluggable attachment handler. The server wires the production
 * implementation (using `processAttachments` from ./attachments). Tests can
 * pass a fake to assert composition without touching the filesystem.
 */
export type AttachmentHandler = (
  attachments: Attachment[],
  fromName: string,
  messageText?: string,
) => Promise<{ annotations: string[] }>

/**
 * Sink type — what the adapter pushes when an inbound message clears the gate.
 * The server wires this to `mcp.notification(...)`. Keeping the interface
 * narrow means we can unit-test the adapter without spinning up MCP.
 */
export type ChannelEventSink = (event: {
  text: string
  conversationId: string
  aadObjectId: string
  tenantId: string
  fromName: string
  messageId?: string
  ts?: string
  /** Group context only: whether the bot was @mentioned. Undefined in 1:1 chats. */
  mentioned?: boolean
  /** Set when the event is about an earlier, already delivered message. */
  messageEvent?: 'edited' | 'restored' | 'deleted'
  /** true for a team channel or a group chat (the same isGroupContext the mention gate uses). */
  isGroup?: boolean
  teamId?: string
  teamName?: string
  /** The Teams channel (channelData.channel), not the Bot Framework channelId 'msteams'. */
  teamsChannelId?: string
  teamsChannelName?: string
  /** The root message id of a channel thread (the `;messageid=` part of the conversation id). */
  threadId?: string
  replyToId?: string
  /** Display names of the people @mentioned in the message, the bot itself excluded. */
  mentions?: string[]
}) => void

// Meta values end up as attributes of the <channel> tag. Team, channel and user names are set by people
// in the tenant, so they are untrusted: no double quotes, angle brackets, control characters (C0, C1) or
// bidi controls, bounded length.
export function metaText(value: unknown, max = 200): string | undefined {
  if (typeof value !== 'string') return undefined
  const cleaned = value
    .replace(/[\u0000-\u001f\u007f-\u009f\u202a-\u202e\u2066-\u2069"<>]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, max)
    .trim()
  return cleaned === '' ? undefined : cleaned
}

export function threadIdFromConversationId(conversationId: string): string | undefined {
  const marker = ';messageid='
  const i = conversationId.indexOf(marker)
  if (i < 0) return undefined
  const rest = conversationId.slice(i + marker.length)
  const end = rest.indexOf(';')
  return metaText(end < 0 ? rest : rest.slice(0, end))
}

/** The meta of the channel notification for one event. Keys are snake_case: Claude Code drops a key that is not [A-Za-z0-9_]+. */
export function channelEventMeta(event: Parameters<ChannelEventSink>[0]): Record<string, string> {
  return {
    aad_object_id: event.aadObjectId,
    conversation_id: event.conversationId,
    tenant_id: event.tenantId,
    from_name: event.fromName,
    ...(event.messageId ? { message_id: event.messageId } : {}),
    ...(event.ts ? { ts: event.ts } : {}),
    ...(event.messageEvent ? { message_event: event.messageEvent } : {}),
    ...(event.mentioned !== undefined ? { mentioned: event.mentioned ? 'true' : 'false' } : {}),
    ...(event.isGroup !== undefined ? { is_group: String(event.isGroup) } : {}),
    ...(event.teamId !== undefined ? { team_id: event.teamId } : {}),
    ...(event.teamName !== undefined ? { team_name: event.teamName } : {}),
    ...(event.teamsChannelId !== undefined ? { channel_id: event.teamsChannelId } : {}),
    ...(event.teamsChannelName !== undefined ? { channel_name: event.teamsChannelName } : {}),
    ...(event.threadId !== undefined ? { thread_id: event.threadId } : {}),
    ...(event.replyToId !== undefined ? { reply_to_id: event.replyToId } : {}),
    ...(event.mentions !== undefined ? { mentions: event.mentions.join(', ') } : {}),
  }
}

/** Hook the adapter calls when an unknown sender needs the pairing DM sent. */
export type PairingDmSender = (
  conversationReference: Partial<ConversationReference>,
  text: string,
) => Promise<void>

export interface AdapterDeps {
  config: Config
  adapter: CloudAdapter
  allowlist: Allowlist
  refs: ConversationRefStore
  onEvent: ChannelEventSink
  /** Phase 3 — pending pairings store. Optional so existing tests still wire. */
  pending?: PendingStore
  /** Phase 3 — permission relay. Optional so existing tests still wire. */
  permission?: PermissionRelay
  /** Phase 3 — send a pairing DM through CloudAdapter. Optional for tests. */
  sendPairingDm?: PairingDmSender
  /**
   * Attachment handler. When set, attachments on inbound messages are
   * downloaded and an annotation per file is appended to the text. Optional
   * so existing gate tests don't need to wire it; production server always
   * sets it.
   */
  attachments?: AttachmentHandler
  /** Optional typing-indicator pump. Started after the gate passes so the
   *  Teams user sees "is typing…" while Claude works. Optional so existing
   *  tests still wire. */
  typingPump?: TypingPump
  /** Outbox for outbound file downloads. When set, the listener serves
   *  `GET /files/<token>` from this store. Optional so existing tests still
   *  wire. */
  outbox?: Outbox
  /** Messages already delivered to Claude; edit/delete/restore events are only
   *  forwarded for these. Optional so existing tests still wire. */
  deliveredMessages?: DeliveredMessages
}

/** Set of AAD IDs we've already logged a "rejected" line for. */
const rejectedSeen = new Set<string>()
/** Senders already told (once) to pair in a 1:1 chat instead of a shared one. */
const groupPairingHinted = new Set<string>()

/**
 * Adapt a Bun-native fetch-style Request to the minimal `Request` interface
 * CloudAdapter expects. The SDK only reads `body`, `headers`, and `method`.
 * `body` must be the parsed JSON (the SDK throws if it gets a stream).
 */
class MalformedBodyError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'MalformedBodyError'
  }
}

async function toBotRequest(req: Request): Promise<{ body: unknown; headers: Record<string, string>; method: string }> {
  // Lowercase all headers so the SDK finds 'authorization' regardless of how
  // the upstream proxy capitalised it.
  const headers: Record<string, string> = {}
  req.headers.forEach((v, k) => {
    headers[k.toLowerCase()] = v
  })
  const raw = await req.text()
  if (raw.length === 0) {
    throw new MalformedBodyError('empty request body')
  }
  let body: unknown
  try {
    body = JSON.parse(raw)
  } catch {
    throw new MalformedBodyError('request body is not valid JSON')
  }
  if (!body || typeof body !== 'object') {
    throw new MalformedBodyError('request body must be a JSON object')
  }
  // CloudAdapter requires `type` on the activity; without it the framework
  // throws a stack-trace-style 500. Fail at the door with a 400 so random
  // scanners don't pollute the audit log.
  if (typeof (body as { type?: unknown }).type !== 'string') {
    throw new MalformedBodyError('request body is missing required activity.type')
  }
  return { body, headers, method: req.method }
}

/**
 * Build the minimal `Response` shape CloudAdapter requires, accumulating the
 * outcome so we can return a Bun-native `Response` once `end()` is called.
 *
 * The SDK calls `.status(n)`, `.header(name, val)`, `.send(body)`, `.end()`
 * in that order. We resolve a single promise on `.end()` so the caller can
 * await the response.
 */
function makeResponseSink(): {
  res: {
    socket: unknown
    status: (code: number) => unknown
    header: (name: string, value: unknown) => unknown
    send: (...args: unknown[]) => unknown
    end: (...args: unknown[]) => unknown
  }
  done: Promise<Response>
} {
  let statusCode = 200
  const headers = new Headers()
  let body: BodyInit | undefined
  let resolve!: (r: Response) => void
  const done = new Promise<Response>(r => (resolve = r))
  const res = {
    // CloudAdapter never uses `socket` for non-WS POSTs but the type
    // demands it exists.
    socket: undefined,
    status(code: number) {
      statusCode = code
      return res
    },
    header(name: string, value: unknown) {
      headers.set(name, String(value))
      return res
    },
    send(...args: unknown[]) {
      if (args.length === 0) return res
      const v = args[0]
      if (v === undefined || v === null) return res
      if (typeof v === 'string') {
        body = v
      } else if (v instanceof Uint8Array || v instanceof ArrayBuffer) {
        body = v as BodyInit
      } else {
        if (!headers.has('content-type')) headers.set('content-type', 'application/json')
        body = JSON.stringify(v)
      }
      return res
    },
    end(...args: unknown[]) {
      // Some callers pass the body to .end() directly.
      if (args.length > 0 && args[0] !== undefined && body === undefined) {
        res.send(...args)
      }
      resolve(new Response(body ?? null, { status: statusCode, headers }))
      return res
    },
  }
  return { res, done }
}

/** The channel / chat part of a conversation id (drops ';messageid=...'). */
export function baseConversationId(conversationId: string): string {
  const i = conversationId.indexOf(';')
  return i === -1 ? conversationId : conversationId.slice(0, i)
}

/**
 * Build the turn handler that bridges the SDK's TurnContext to our gate + sink.
 *
 * Exported so the unit tests can call it directly with a synthesised
 * TurnContext — bypassing the JWT validator (which requires Microsoft-signed
 * tokens, impractical to produce in tests).
 */
export function makeTurnHandler(deps: AdapterDeps): (ctx: TurnContext) => Promise<void> {
  return async function turn(ctx: TurnContext) {
    const a = ctx.activity
    // Diagnostic: log every inbound so we know what arrived.
    process.stderr.write(
      `teams channel: inbound type=${a.type} from=${a.from?.aadObjectId ?? '<none>'} convId=${a.conversation?.id?.slice(0, 20) ?? '<none>'} text=${JSON.stringify((a.text ?? '').slice(0, 60))}\n`,
    )
    // We only care about message activities in v1. Conversation update,
    // typing, etc. are ignored — they don't carry a payload Claude needs.
    const eventType = (a.channelData as { eventType?: unknown } | undefined)?.eventType
    const messageEvent: 'edited' | 'restored' | 'deleted' | undefined =
      a.type === 'messageUpdate' && eventType === 'editMessage' ? 'edited'
      : a.type === 'messageUpdate' && eventType === 'undeleteMessage' ? 'restored'
      : a.type === 'messageDelete' && eventType === 'softDeleteMessage' ? 'deleted'
      : undefined
    if (a.type !== 'message' && !messageEvent) return

    const aadObjectIdRaw = a.from?.aadObjectId
    if (!aadObjectIdRaw) {
      // Activities without an AAD Object ID can't be authenticated against
      // the tenant allowlist. Drop. Could be a Web Chat test message.
      process.stderr.write('teams channel: drop — no aadObjectId on activity\n')
      return
    }
    const aadObjectId = aadObjectIdRaw.toLowerCase()
    const tenantId = a.conversation?.tenantId

    // Defence-in-depth tenant check: even though the SDK already validates
    // the JWT against the configured tenant, we compare the conversation
    // tenant against the pinned tenant. Belt and braces — protects against
    // a hypothetical SDK bug or future config drift that lets a wider
    // audience through.
    if (deps.config.tenantId && tenantId && tenantId.toLowerCase() !== deps.config.tenantId.toLowerCase()) {
      process.stderr.write(
        `teams channel: drop — tenant mismatch (got ${tenantId}, expected pinned tenant)\n`,
      )
      return
    }

    const conversationId = a.conversation?.id
    if (!conversationId) {
      // Shouldn't happen for a message activity, but defend anyway.
      process.stderr.write('teams channel: drop — no conversation.id\n')
      return
    }
    const fromName = a.from?.name ?? ''

    // Channel / group-chat scope (manifest scopes 'team' + 'groupChat'). In a
    // SHARED conversation the bot must act ONLY when explicitly @mentioned --
    // otherwise it would react to every message (noise) and widen interaction
    // beyond intent. 1:1 ('personal') conversations are unaffected: no mention
    // is required and the text is used verbatim below, so the existing DM path
    // cannot regress. NOTE: this gates WHICH messages are considered; WHO may
    // interact is still the per-user aadObjectId allowlist below (unchanged) --
    // a channel member is not implicitly allowed just because the bot is in the
    // channel; each sender still pairs individually.
    // Robust shared-conversation detection. `conversationType` ('channel' /
    // 'groupChat') is the documented signal, but Teams does not always populate
    // it the way we expect (observed live 2026-06-29: a real channel message --
    // conversation id `19:...@thread.tacv2` -- arrived with conversationType not
    // equal to 'channel', so the old single-signal check fell through to the
    // personal path: surfaced verbatim, no @mention gate, no mention strip). So
    // we OR in two reliable signals: `conversation.isGroup` (true for any shared
    // conversation, false/undefined for a 1:1) and the `@thread.tacv2` team-channel
    // conversation-id shape. 1:1-SAFE by construction: a personal chat has
    // conversationType='personal' (not in the set), isGroup falsey, and a conv-id
    // that is NOT `@thread.tacv2` (that suffix is channel-specific). We deliberately
    // do NOT pattern-match `@thread.v2`/`@thread.skype` -- those are not needed
    // (group chats set isGroup) and matching them risks a false-positive on a 1:1.
    const conversationType = a.conversation?.conversationType
    const isGroupConv = (a.conversation as { isGroup?: boolean } | undefined)?.isGroup === true
    const threadConv = /@thread\.tacv2/.test(conversationId)
    const isGroupContext =
      conversationType === 'channel' ||
      conversationType === 'groupChat' ||
      isGroupConv ||
      threadConv
    if (messageEvent) {
      // Update of an earlier message. Runs before the @mention gate, pairing and
      // everything else: only a message we delivered, from the same still-allowed
      // sender, is forwarded. Never pairs, never sends anything to Teams.
      const targetId = a.id
      if (typeof targetId !== 'string' || targetId.length === 0) {
        process.stderr.write(`teams channel: drop — ${messageEvent} event without activity id\n`)
        return
      }
      const entry = deps.deliveredMessages?.get(baseConversationId(conversationId), targetId)
      if (!entry) {
        process.stderr.write(`teams channel: drop — ${messageEvent} event for a message that was not delivered\n`)
        return
      }
      if (entry.aadObjectId !== aadObjectId) {
        process.stderr.write(`teams channel: drop — ${messageEvent} event from a different sender than the original\n`)
        return
      }
      if (!deps.allowlist.isAllowed(aadObjectId)) {
        process.stderr.write(`teams channel: drop — ${messageEvent} event from a sender no longer in allowlist\n`)
        return
      }
      let eventText: string
      if (messageEvent === 'deleted') {
        eventText = '[Message deleted]'
      } else {
        const body = isGroupContext
          ? (TurnContext.removeRecipientMention(a) ?? '').trim()
          : (typeof a.text === 'string' ? a.text.trim() : '')
        const label = messageEvent === 'edited' ? '[Message edited]' : '[Message restored]'
        eventText = `${label} ${body.length > 0 ? body : '(the new text was not included)'}`
      }
      deps.onEvent({
        text: eventText,
        conversationId,
        aadObjectId,
        tenantId: tenantId ?? deps.config.tenantId,
        fromName,
        messageId: targetId,
        ts: a.timestamp ? new Date(a.timestamp).toISOString() : undefined,
        messageEvent,
      })
      return
    }
    // Diagnostic: the context signals on every inbound -- cheap, and the only way
    // to tell a mis-detected channel from a 1:1 after the fact (no other log has it).
    const mentionIds = (a.entities ?? [])
      .filter((e) => (e as { type?: string }).type === 'mention')
      .map((e) => (e as { mentioned?: { id?: string } }).mentioned?.id)
    process.stderr.write(
      `teams channel: ctx convType=${conversationType ?? '<none>'} isGroup=${isGroupConv} thread=${threadConv} -> group=${isGroupContext} recipient=${a.recipient?.id ?? '<none>'} mentions=${JSON.stringify(mentionIds)} convId=${conversationId.slice(0, 32)}\n`,
    )
    // One allowlist decision per turn: the unmentioned path and the gate below
    // must not disagree because the allowlist changed between two reads.
    const senderAllowed = deps.allowlist.isAllowed(aadObjectId)
    let mentioned: boolean | undefined
    if (isGroupContext) {
      const botId = a.recipient?.id
      mentioned = botId
        ? (a.entities ?? []).some(
            (e) =>
              (e as { type?: string }).type === 'mention' &&
              (e as { mentioned?: { id?: string } }).mentioned?.id === botId,
          )
        : false
      if (!mentioned) {
        const optOut = deps.config.unmentionedConversations
        const unmentionedOk = optOut !== undefined && (optOut.has('*') || optOut.has(baseConversationId(conversationId)))
        if (!unmentionedOk) {
          process.stderr.write(
            `teams channel: ignore: ${conversationType} message without bot @mention (convId=${conversationId.slice(0, 20)})\n`,
          )
          return
        }
        // Delivered without a mention: allowlisted senders only. A non-allowlisted
        // sender's unmentioned message never starts the pairing flow -- otherwise
        // every channel member who writes anything would get a pairing DM.
        if (!senderAllowed) {
          process.stderr.write(
            `teams channel: ignore: unmentioned ${conversationType ?? 'group'} message from a non-allowlisted sender (convId=${conversationId.slice(0, 20)})\n`,
          )
          return
        }
      }
    }

    // Allowlist gate — the single most important application-level check.
    // A non-allowlisted sender enters the pairing path (Phase 3) if a
    // pending store is wired; otherwise it's dropped silently.
    if (!senderAllowed) {
      if (deps.pending && deps.sendPairingDm) {
        const ref = TurnContext.getConversationReference(a)
        // Shared by the adapter's own detection, or by the relay's rule (isPersonalReference: an untyped
        // '@thread.' id is a group chat). An explicit 'personal' type stays 1:1. The @mention gate is not changed.
        if (isGroupContext || !isPersonalReference(ref)) {
          // Never put a pairing code in a conversation other people read.
          if (!groupPairingHinted.has(aadObjectId)) {
            groupPairingHinted.add(aadObjectId)
            process.stderr.write(
              `teams channel: pairing not offered in a shared conversation (aad=${aadObjectId})\n`,
            )
            deps.sendPairingDm(
              ref,
              'This bot is gated. To pair, send me a direct (1:1) message.',
            ).catch(err => {
              process.stderr.write(
                `teams channel: pairing DM send failed: ${err}\n`,
              )
            })
          }
          return
        }
        const decision = deps.pending.recordIncoming({
          aadObjectId,
          tenantId: tenantId ?? deps.config.tenantId,
          fromName,
          conversationId,
        })
        if (decision.action === 'suppress') {
          // Hit the per-sender reply cap or the global pending cap — drop
          // silently. Log once per AAD ID so the operator sees an attempted
          // amplification but the line doesn't repeat.
          if (!rejectedSeen.has(aadObjectId)) {
            rejectedSeen.add(aadObjectId)
            process.stderr.write(
              `teams channel: pairing suppressed for ${aadObjectId} (${decision.reason})\n`,
            )
          }
          return
        }
        const lead =
          decision.action === 'send_reminder' ? 'Still pending' : 'Hi'
        const text =
          `${lead} — this bot is gated. Ask the operator to run ` +
          `/teams:access pair ${decision.entry.code} in their terminal. ` +
          `Show this code: ${decision.entry.code}.`
        process.stderr.write(
          `teams channel: pairing ${decision.action} pair_id=${decision.entry.pair_id} aad=${aadObjectId}\n`,
        )
        deps.sendPairingDm(ref, text).catch(err => {
          process.stderr.write(
            `teams channel: pairing DM send failed: ${err}\n`,
          )
        })
        return
      }
      if (!rejectedSeen.has(aadObjectId)) {
        rejectedSeen.add(aadObjectId)
        process.stderr.write(
          `teams channel: drop — sender not in allowlist (aad_object_id=${aadObjectId})\n`,
        )
      }
      return
    }

    // Capture the conversation reference for the outbound side. We do this
    // before pushing to Claude so a fast reply can always find its way back.
    const ref = TurnContext.getConversationReference(a)
    deps.refs.put(conversationId, ref, aadObjectId)

    // Start the typing-indicator pump as early as possible — the user sees
    // "is typing…" while attachments are processed and Claude is notified.
    // The reply tool stops it when the response lands.
    deps.typingPump?.start(conversationId, ref)

    // In a channel/group, strip the bot's own <at>…</at> mention so Claude sees
    // the actual instruction ("@Bot summarize this" -> "summarize this"). For
    // personal chats there is no recipient mention, so use the text verbatim.
    const rawText = isGroupContext
      ? (TurnContext.removeRecipientMention(a) ?? '').trim()
      : (typeof a.text === 'string' ? a.text : '')
    const messageId = a.id
    const ts = a.timestamp ? new Date(a.timestamp).toISOString() : undefined

    // Attachment handling — runs only for allowlisted senders (we're below
    // the allowlist gate). Failures degrade to annotation lines so the
    // message still gets through. We deliberately await here: Hermes needs
    // file paths in the inbound MCP event, not a follow-up notification.
    let text = rawText
    const rawAttachments = Array.isArray(a.attachments) ? (a.attachments as Attachment[]) : []
    if (deps.attachments && rawAttachments.length > 0) {
      try {
        const { annotations } = await deps.attachments(rawAttachments, fromName, rawText)
        if (annotations.length > 0) {
          text = text.length > 0 ? `${text}\n\n${annotations.join('\n')}` : annotations.join('\n')
        }
      } catch (err) {
        // Defensive: a thrown error inside the handler must not drop the
        // message. Annotate so Hermes can apologise to the user.
        process.stderr.write(
          `teams channel: attachment handler threw: ${err}\n`,
        )
        const note = `[Attachment handling failed: ${err instanceof Error ? err.message : String(err)}]`
        text = text.length > 0 ? `${text}\n\n${note}` : note
      }
    }

    // Permission-reply intercept — the sender is already allowlisted, so we
    // trust them with verdicts. Strict regex (see permission/relay.ts) means
    // a bare "yes" or chatty prefix falls through to the regular channel
    // event. If the id doesn't match a pending request, treat the text as a
    // normal message — better than swallowing it silently.
    // A tool-approval verdict must be addressed to the bot: in a shared
    // conversation an unmentioned "yes <id>" is overheard text, not an answer.
    if (deps.permission && mentioned !== false) {
      // Match against the raw text — annotations appended for attachments
      // would otherwise break the anchored regex.
      const match = PERMISSION_REPLY_RE.exec(rawText)
      if (match) {
        const verdict = match[1]!.toLowerCase().startsWith('y') ? 'allow' : 'deny'
        const id = match[2]!.toLowerCase()
        const handled = deps.permission.resolve(id, verdict)
        if (handled) {
          process.stderr.write(
            `teams channel: permission verdict ${verdict} for ${id}\n`,
          )
          return
        }
        // Fall through — unknown id may be a coincidence (operator typing
        // about a different request). Forward as chat.
      }
    }

    if (typeof messageId === 'string' && messageId.length > 0) {
      deps.deliveredMessages?.remember(baseConversationId(conversationId), messageId, aadObjectId)
    }
    // Team, channel, thread, reply and mention context only in a team channel or group chat.
    const cd: any = isGroupContext ? a.channelData : undefined
    const obj = (v: unknown): any => (v !== null && typeof v === 'object' ? v : undefined)
    const team = obj(obj(cd)?.team)
    const chan = obj(obj(cd)?.channel)
    const mentionNames: string[] = []
    // At most 20 names, and the joined value (", ") stays within 200 characters: whole names only.
    let mentionsLen = 0
    for (const e of isGroupContext ? (a.entities ?? []) : []) {
      if ((e as { type?: string }).type !== 'mention') continue
      const m = obj((e as { mentioned?: unknown }).mentioned)
      if (!m || m.id === a.recipient?.id) continue
      const n = metaText(m.name) ?? metaText(m.id)
      if (n === undefined || mentionNames.includes(n)) continue
      const len = mentionNames.length === 0 ? n.length : mentionsLen + 2 + n.length
      if (len > 200) break
      mentionNames.push(n)
      mentionsLen = len
      if (mentionNames.length >= 20) break
    }

    deps.onEvent({
      text,
      conversationId,
      aadObjectId,
      tenantId: tenantId ?? deps.config.tenantId,
      fromName,
      messageId,
      ts,
      ...(mentioned !== undefined ? { mentioned } : {}),
      isGroup: isGroupContext,
      teamId: metaText(team?.id),
      teamName: metaText(team?.name),
      teamsChannelId: metaText(chan?.id),
      teamsChannelName: metaText(chan?.name),
      threadId: isGroupContext ? threadIdFromConversationId(conversationId) : undefined,
      replyToId: isGroupContext ? metaText(a.replyToId) : undefined,
      mentions: mentionNames.length > 0 ? mentionNames : undefined,
    })
  }
}

export interface RunningListener {
  port: number
  url: string
  stop(): Promise<void>
}

/**
 * Start the HTTP listener and return a handle the caller can stop later.
 * Mounts:
 *
 *   POST /api/messages — Bot Framework messaging endpoint. The reverse proxy
 *                         (Caddy / nginx / Cloudflare Tunnel) forwards here.
 *   GET  /health        — Unauthenticated liveness. Returns "ok".
 */
export function startListener(deps: AdapterDeps): RunningListener {
  const { config, adapter } = deps
  const turn = makeTurnHandler(deps)

  const server = Bun.serve({
    hostname: config.bindHost,
    port: config.port,
    development: false,
    // Fail loudly with EADDRINUSE if another instance is already listening.
    // Without this, Bun defaults to SO_REUSEPORT and lets multiple processes
    // share the port — Linux then distributes inbound connections round-robin,
    // which means a stray plugin invocation in another shell silently steals
    // half the Teams traffic. Confirmed during the 2026-05-18 cutover.
    reusePort: false,
    async fetch(req): Promise<Response> {
      const url = new URL(req.url)
      // GET /health — lets the reverse proxy + monitoring poke us without
      // a JWT. No state leak; just "process is alive".
      if (req.method === 'GET' && url.pathname === '/health') {
        return new Response('ok', { status: 200, headers: { 'content-type': 'text/plain' } })
      }
      // /files/<token> — outbound file download. No JWT here; the token IS
      // the secret. See docs/specs/2026-05-19-outbound-file-transfers.md.
      if (url.pathname.startsWith('/files/')) {
        if (!deps.outbox) {
          return new Response('not found', { status: 404 })
        }
        return handleFilesRequest(req, deps.outbox)
      }
      if (url.pathname === '/api/messages') {
        if (req.method !== 'POST') {
          return new Response('method not allowed', { status: 405 })
        }
        // Parse the body first so we can distinguish a malformed payload
        // from a real downstream error. Random scanners poking the public
        // endpoint shouldn't produce 500-stack-traces in the audit log.
        let botReq: { body: unknown; headers: Record<string, string>; method: string }
        try {
          botReq = await toBotRequest(req)
        } catch (err) {
          if (err instanceof MalformedBodyError) {
            // Info-level log only — these are noise, not incidents.
            process.stderr.write(
              `teams channel: 400 malformed /api/messages — ${err.message}\n`,
            )
            return new Response(err.message, {
              status: 400,
              headers: { 'content-type': 'text/plain' },
            })
          }
          process.stderr.write(`teams channel: body read failed: ${err}\n`)
          return new Response('bad request', { status: 400 })
        }
        const { res, done } = makeResponseSink()
        // Hand the request to CloudAdapter. The SDK validates the JWT, parses
        // the activity, calls our turn handler. If the turn handler throws,
        // it's surfaced as a 500. If JWT validation fails, 401.
        try {
          // The SDK's Request type is structurally compatible with botReq.
          // Cast at the boundary so the type-check stays honest elsewhere.
          await adapter.process(botReq as never, res as never, turn)
        } catch (err) {
          process.stderr.write(`teams channel: adapter.process threw: ${err}\n`)
          // adapter.process already wrote its own response on most errors;
          // if not, give the connector a 500 so it retries.
          // (The done promise hasn't resolved yet → we resolve manually.)
          res.status(500)
          res.end('internal error')
        }
        return done
      }
      return new Response('not found', { status: 404 })
    },
    error(err) {
      process.stderr.write(`teams channel: server error: ${err.message ?? err}\n`)
      return new Response('internal error', { status: 500 })
    },
  })

  // Note the listener URL — used by tests and logged on startup.
  const port = Number(server.port)
  const url = `http://${config.bindHost}:${port}`
  process.stderr.write(`teams channel: listening on ${url} (tenant pinned to ${config.tenantId.slice(0, 8)}…)\n`)

  return {
    port,
    url,
    async stop() {
      await server.stop(true)
    },
  }
}
