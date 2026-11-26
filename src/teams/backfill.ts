/**
 * Graph backfill of channel messages Teams never delivered to the bot (opt-in, TEAMS_BACKFILL).
 *
 * Teams occasionally drops a channel message on the floor: the bot has the RSC permission
 * ChannelMessage.Read.Group anyway (thread history), so a slow poll over Microsoft Graph can find what was lost.
 * Every message found that way is fed through the SAME turn handler as a live message — same tenant check, same
 * allowlist, same @mention gate, same dedup store — so exactly one of the two paths delivers it, in either order
 * and across a restart. The poll only ever awaits its own calls; it never blocks or delays the live path.
 *
 * What it does per channel (only channels with a stored thread reference, i.e. ones a live message already came
 * from): read the newest 50 roots with their inline replies, pick the messages from humans (not bots, not system
 * events, not deleted) created in (watermark, now - grace], and submit them oldest first. `grace` gives the live
 * path the first chance; the watermark is only advanced by a cycle that fully succeeded, and a failing submit sets
 * it back so the message is retried. A 403 (no RSC grant in that team) shelves the channel for six hours; a
 * network error, a 5xx or a 429 stops the whole cycle and grows the delay before the next one.
 *
 * Limits: 5 pages per channel per cycle, 200 inline replies per root (Graph's $expand=replies cap — a thread with
 * more replies than that is only partially backfilled), edits are not re-delivered, attachments are not
 * reconstructed (a note is appended instead).
 */

import type { Activity } from 'botbuilder'
import {
  closeSync,
  constants,
  existsSync,
  fstatSync,
  mkdirSync,
  openSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'fs'
import { dirname } from 'path'
import type { ConversationRefStore } from './conversationRefs.js'
import type { DecisionLog } from './decisionLog.js'
import { htmlToText, isGraphV1Link, type FetchFn } from './threadHistory.js'

const GRAPH_PREFIX = 'https://graph.microsoft.com/v1.0/'
/** Graph pages one channel per cycle: $top is capped at 50 and $expand=replies at 200 inline replies per root. */
const MAX_PAGES = 5
const MAX_CHANNELS = 1000
/** How long a channel that answered 403 (no ChannelMessage.Read.Group grant in that team) is left alone. */
const FORBIDDEN_MS = 6 * 60 * 60 * 1000
const MAX_BACKOFF_MS = 30 * 60 * 1000
/** A message whose submit keeps failing is given up on after this many tries. */
const MAX_SUBMIT_FAILURES = 3

/** The subset of a Graph chatMessage this code reads. */
export interface GraphChatMessage {
  id?: string
  replyToId?: string | null
  createdDateTime?: string
  lastModifiedDateTime?: string | null
  messageType?: string
  deletedDateTime?: string | null
  from?: { user?: { id?: string; displayName?: string }; application?: { id?: string; displayName?: string } } | null
  body?: { contentType?: string; content?: string }
  mentions?: Array<{ mentionText?: string; mentioned?: { user?: { id?: string; displayName?: string }; application?: { id?: string } } }>
  attachments?: unknown[]
  hostedContents?: unknown[]
  /** Only the first 200 replies of a root come inline ($expand=replies); replies@odata.nextLink is ignored. */
  replies?: GraphChatMessage[]
}

export interface BackfillDeps {
  getToken(): Promise<string>
  fetch?: FetchFn
  now?: () => number
  refs: Pick<ConversationRefStore, 'channelThreadRefs'>
  /** The live channel -> team map the adapter fills (teamGroups.get). */
  teamGroupFor(channelId: string): string | undefined
  botAppId: string
  tenantId: string
  stateFile: string
  intervalMs: number
  graceMs: number
  maxAgeMs: number
  /** Runs the live message handler on the rebuilt activity (listener.submitActivity). */
  submit(activity: Activity): Promise<void>
  decisions?: DecisionLog
  setTimer?: (fn: () => void, ms: number) => unknown
  clearTimer?: (t: unknown) => void
}

interface BackfillState {
  /** Per channel: everything created at or before this time has been handed over. */
  watermarks: Map<string, number>
  /** Per channel: the team AAD group id, learned from the live map so it survives a restart. */
  groups: Map<string, string>
  /** Per channel: do not poll again before this time (403). */
  forbiddenUntil: Map<string, number>
}

/** Why a channel (or the whole cycle) gave up. `forbidden` is per channel only; the rest stop the cycle. */
type Failure = 'network' | 'http_5xx' | 'http_error' | 'bad_response' | 'throttled' | 'forbidden'

interface CycleOutcome {
  ok: boolean
  /** Set for a 429 that carried Retry-After. */
  retryAfterMs?: number
}

function errMsg(err: unknown): string {
  return err instanceof Error ? err.message : String(err)
}

function timeMs(iso: string | null | undefined): number | undefined {
  if (typeof iso !== 'string' || iso.length === 0) return undefined
  const t = Date.parse(iso)
  return Number.isFinite(t) ? t : undefined
}

/** The newest lastModifiedDateTime of a root's whole reply chain (createdDateTime when a message was never touched). */
function chainLastModifiedMs(root: GraphChatMessage): number {
  let newest = timeMs(root.lastModifiedDateTime) ?? timeMs(root.createdDateTime) ?? Number.NEGATIVE_INFINITY
  for (const reply of Array.isArray(root.replies) ? root.replies : []) {
    const at = timeMs(reply.lastModifiedDateTime) ?? timeMs(reply.createdDateTime)
    if (at !== undefined && at > newest) newest = at
  }
  return newest
}

/** The body of a Graph message as the text of a Bot Framework activity: html is flattened the way thread history
 *  flattens it, but the <at>Name</at> tags stay (a live activity carries them, and the mention gate needs them). */
export function graphBodyToText(m: GraphChatMessage): string {
  const content = typeof m.body?.content === 'string' ? m.body.content : ''
  // Graph writes mentions as <at id="0">Name</at>; htmlToText with keepMentionTags turns them into the <at>Name</at>
  // a Bot Framework activity carries (the mention gate needs them).
  const text = m.body?.contentType === 'html' ? htmlToText(content, { keepMentionTags: true }) : content
  const count =
    (Array.isArray(m.attachments) ? m.attachments.length : 0) + (Array.isArray(m.hostedContents) ? m.hostedContents.length : 0)
  return count > 0 ? `${text}\n[backfilled: ${count} attachment(s) not included]` : text
}

export interface ActivityFromGraphArgs {
  /** The channel id (the conversation id without ';messageid=...'). */
  channelId: string
  /** The team's AAD group id (Graph's team id). */
  groupId: string
  tenantId: string
  serviceUrl: string
  /** The bot's MRI, from the stored conversation reference. */
  botId: string
  /** The bot's display name from the stored reference, when known: what the live path has as recipient.name. */
  botName?: string
  /** The bot's app (client) id, to recognise mentions of this bot. */
  botAppId: string
  /** The root message id of the thread (m.id for a root, the root's id for a reply). */
  rootId: string
  message: GraphChatMessage
}

/** A Bot Framework mention entity. */
type MentionEntity = { type: 'mention'; text: string; mentioned: { id: string; name: string } }

/**
 * A Bot Framework activity that looks like the one Teams would have sent for this Graph message, minus what the
 * poll cannot know: from has NO id (only aadObjectId), so nothing about the sender is learned from it, and
 * channelData.backfilled marks it so the event and the decision log can say which messages were recovered.
 */
export function activityFromGraph(a: ActivityFromGraphArgs): Activity {
  const m = a.message
  const user = m.from?.user
  const displayName = user?.displayName
  const isReply = typeof m.replyToId === 'string' && m.replyToId.length > 0
  const entities: MentionEntity[] = []
  for (const mention of Array.isArray(m.mentions) ? m.mentions : []) {
    const name = typeof mention.mentionText === 'string' ? mention.mentionText : ''
    const mentioned = mention.mentioned
    const appId = mentioned?.application?.id
    const userId = mentioned?.user?.id
    const userName = mentioned?.user?.displayName
    if (typeof appId === 'string' && appId === a.botAppId) {
      // The bot itself: the mention gate compares this id with the recipient's.
      entities.push({ type: 'mention', text: `<at>${name}</at>`, mentioned: { id: a.botId, name } })
    } else if (typeof userId === 'string' && userId.length > 0) {
      entities.push({
        type: 'mention',
        text: `<at>${name}</at>`,
        mentioned: { id: userId, name: typeof userName === 'string' ? userName : name },
      })
    }
  }
  return {
    type: 'message',
    id: m.id,
    // Graph's ISO string; the turn handler only ever does new Date(a.timestamp).
    timestamp: m.createdDateTime as unknown as Activity['timestamp'],
    channelId: 'msteams',
    serviceUrl: a.serviceUrl,
    from: { aadObjectId: user?.id, name: typeof displayName === 'string' ? displayName : '' },
    recipient: typeof a.botName === 'string' && a.botName.length > 0 ? { id: a.botId, name: a.botName } : { id: a.botId },
    conversation: {
      id: `${a.channelId};messageid=${a.rootId}`,
      conversationType: 'channel',
      isGroup: true,
      tenantId: a.tenantId,
    },
    ...(isReply ? { replyToId: a.rootId } : {}),
    text: graphBodyToText(m),
    entities: entities as Activity['entities'],
    channelData: {
      tenant: { id: a.tenantId },
      team: { aadGroupId: a.groupId },
      channel: { id: a.channelId },
      backfilled: true,
    },
    // Deliberately not a full Activity: from carries no id, conversation no name, and recipient a name only when the
    // stored reference had one (see above). The
    // live path likewise checks only `type` on the posted JSON before the handler sees it.
  } as unknown as Activity
}

export interface Backfill {
  /** One poll cycle. Returns immediately (does nothing) when one is already running. */
  runOnce(): Promise<void>
  /** Schedules the first cycle `intervalMs` from now, then one after every cycle. */
  start(): void
  /** Clears the timer; a stopped backfill never schedules again. */
  stop(): void
}

export function createBackfill(deps: BackfillDeps): Backfill {
  const now = deps.now ?? Date.now
  const doFetch: FetchFn = deps.fetch ?? ((u, i) => fetch(u, i))
  const state: BackfillState = { watermarks: new Map(), groups: new Map(), forbiddenUntil: new Map() }
  const submitFailures = new Map<string, number>()
  const notedNoTeam = new Set<string>()
  const notedForbidden = new Set<string>()
  const notedPageLimit = new Set<string>()
  let warned = false
  let running = false
  let stopped = false
  let timer: unknown
  let failures = 0

  const warn = (line: string): void => {
    if (warned) return
    warned = true
    process.stderr.write(`${line}\n`)
  }

  // ── state file ────────────────────────────────────────────────────────────────────────────────────────────

  /** Re-inserts the key, so a Map's insertion order is the order channels were last updated, then bounds the map. */
  function touch<V>(map: Map<string, V>, key: string, value: V): void {
    map.delete(key)
    map.set(key, value)
    while (map.size > MAX_CHANNELS) map.delete(map.keys().next().value as string)
  }

  function readMap(parsed: unknown, key: string, valueIs: (v: unknown) => boolean): Map<string, unknown> {
    const raw = (parsed as Record<string, unknown> | null)?.[key]
    if (raw === undefined || raw === null) return new Map()
    if (typeof raw !== 'object' || Array.isArray(raw)) throw new Error('unexpected format')
    const out = new Map<string, unknown>()
    for (const [k, v] of Object.entries(raw as Record<string, unknown>)) {
      if (typeof k !== 'string' || k.length === 0 || k.length > 512 || !valueIs(v)) throw new Error('unexpected format')
      out.set(k, v)
    }
    return out
  }

  function loadState(path: string): void {
    let fd: number
    try {
      // O_NONBLOCK: opening a FIFO here must not hang startup before the regular-file check below.
      fd = openSync(path, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK)
    } catch (err: any) {
      if (err?.code !== 'ENOENT') {
        warn(`teams channel: backfill state file cannot be opened (${err?.code === 'ELOOP' ? 'a symlink' : errMsg(err)}), starting empty`)
      }
      return
    }
    try {
      const st = fstatSync(fd)
      if (!st.isFile()) warn('teams channel: backfill state file is not a regular file, starting empty')
      else {
        const me = typeof process.getuid === 'function' ? process.getuid() : undefined
        if (me !== undefined && st.uid !== me) warn('teams channel: backfill state file belongs to another user, starting empty')
        else if (st.mode & 0o022) warn('teams channel: backfill state file is writable by group or others, starting empty')
        else {
          const parsed: unknown = JSON.parse(readFileSync(fd, 'utf8'))
          if (!parsed || typeof parsed !== 'object' || (parsed as { version?: unknown }).version !== 1) {
            throw new Error('unexpected format')
          }
          const watermarks = readMap(parsed, 'watermarks', v => typeof v === 'number' && Number.isFinite(v))
          const groups = readMap(parsed, 'groups', v => typeof v === 'string' && (v as string).length > 0 && (v as string).length <= 512)
          const forbidden = readMap(parsed, 'forbiddenUntil', v => typeof v === 'number' && Number.isFinite(v))
          for (const [k, v] of watermarks) state.watermarks.set(k, v as number)
          for (const [k, v] of groups) state.groups.set(k, v as string)
          for (const [k, v] of forbidden) state.forbiddenUntil.set(k, v as number)
        }
      }
    } catch (err) {
      warn(`teams channel: backfill state file unreadable, starting empty (${errMsg(err)})`)
    } finally {
      try { closeSync(fd) } catch {}
    }
  }

  function saveState(): void {
    const tmp = `${deps.stateFile}.tmp`
    try {
      const dir = dirname(deps.stateFile)
      if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 })
      writeFileSync(
        tmp,
        JSON.stringify({
          version: 1,
          watermarks: Object.fromEntries(state.watermarks),
          groups: Object.fromEntries(state.groups),
          forbiddenUntil: Object.fromEntries(state.forbiddenUntil),
        }),
        { mode: 0o600 },
      )
      renameSync(tmp, deps.stateFile)
    } catch (err) {
      try { rmSync(tmp, { force: true }) } catch {}
      warn(`teams channel: could not save backfill state (${errMsg(err)})`)
    }
  }

  loadState(deps.stateFile)

  // ── one channel ───────────────────────────────────────────────────────────────────────────────────────────

  function retryAfterMs(res: Response): number | undefined {
    let raw: string | null
    try {
      raw = res.headers.get('retry-after')
    } catch {
      return undefined
    }
    if (raw === null) return undefined
    const seconds = Number.parseInt(raw, 10)
    return Number.isFinite(seconds) && seconds > 0 ? seconds * 1000 : undefined
  }

  /** One channel's pages: newest first, following a v1 next link while the page's last root is still in range. */
  async function fetchChannel(
    groupId: string,
    channelId: string,
    lower: number,
  ): Promise<{ roots: GraphChatMessage[]; truncated: boolean } | { failure: Failure; retryAfterMs?: number }> {
    let url: string | undefined =
      `${GRAPH_PREFIX}teams/${encodeURIComponent(groupId)}` +
      `/channels/${encodeURIComponent(channelId)}` +
      `/messages?$top=50&$expand=replies`
    const roots: GraphChatMessage[] = []
    for (let page = 0; page < MAX_PAGES; page++) {
      const current = url
      if (current === undefined) break
      const token = await deps.getToken()
      let res: Response
      try {
        res = await doFetch(current, {
          method: 'GET',
          headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
          redirect: 'error',
        })
      } catch {
        return { failure: 'network' }
      }
      if (res.status === 403) return { failure: 'forbidden' }
      if (res.status === 429) return { failure: 'throttled', retryAfterMs: retryAfterMs(res) }
      if (res.status >= 500) return { failure: 'http_5xx' }
      if (res.status < 200 || res.status >= 300) return { failure: 'http_error' }
      let body: { value?: unknown; '@odata.nextLink'?: unknown }
      try {
        body = (await res.json()) as { value?: unknown; '@odata.nextLink'?: unknown }
      } catch {
        return { failure: 'bad_response' }
      }
      const value = Array.isArray(body?.value) ? (body.value as GraphChatMessage[]) : []
      roots.push(...value)
      const last = value[value.length - 1]
      let next: string | undefined = typeof body?.['@odata.nextLink'] === 'string' ? body['@odata.nextLink'] : undefined
      // A next link may only point at Graph v1.0, and only when the page's last root is still in range: the
      // response is sorted by the last modified date of the whole reply chain, newest first.
      if (next !== undefined && !isGraphV1Link(next)) next = undefined
      else if (next !== undefined && (last === undefined || chainLastModifiedMs(last) < lower)) next = undefined
      url = next
    }
    // A next link still set here means the page budget ran out while chains in range were left unread.
    return { roots, truncated: url !== undefined }
  }

  async function runChannel(
    entry: { channelId: string; conversationId: string; serviceUrl: string; botId: string; botName?: string },
    nowAtStart: number,
    upper: number,
  ): Promise<CycleOutcome> {
    const channelId = entry.channelId
    const live = deps.teamGroupFor(channelId)
    if (live !== undefined && state.groups.get(channelId) !== live) {
      touch(state.groups, channelId, live)
      saveState()
    }
    const groupId = live ?? state.groups.get(channelId)
    if (groupId === undefined) {
      // The live map learns a channel's team from a channel message; until one arrived there is nothing to poll.
      if (!notedNoTeam.has(channelId)) {
        notedNoTeam.add(channelId)
        deps.decisions?.record({
          decision: 'backfill_skip',
          reason: 'no_team_id',
          source: 'backfill',
          conversationId: entry.conversationId,
        })
      }
      return { ok: true }
    }
    if ((state.forbiddenUntil.get(channelId) ?? 0) > nowAtStart) return { ok: true }

    const watermark = state.watermarks.get(channelId)
    if (watermark === undefined) {
      // First sight: nothing older than the grace window is ever picked up, so the poll does not replay history.
      touch(state.watermarks, channelId, upper)
      saveState()
      return { ok: true }
    }
    const lower = Math.max(watermark, nowAtStart - deps.maxAgeMs)

    const pages = await fetchChannel(groupId, channelId, lower)
    if ('failure' in pages) {
      if (pages.failure === 'forbidden') {
        touch(state.forbiddenUntil, channelId, nowAtStart + FORBIDDEN_MS)
        saveState()
        if (!notedForbidden.has(channelId)) {
          notedForbidden.add(channelId)
          deps.decisions?.record({
            decision: 'backfill_error',
            reason: 'forbidden',
            source: 'backfill',
            conversationId: entry.conversationId,
          })
          process.stderr.write(
            `teams channel: backfill got 403 for a channel in team ${groupId.slice(0, 8)}…: the bot needs the ` +
              'ChannelMessage.Read.Group grant there (scripts/make-teams-manifest.sh --rsc-all-messages)\n',
          )
        }
        return { ok: true }
      }
      deps.decisions?.record({ decision: 'backfill_error', reason: pages.failure, source: 'backfill' })
      return { ok: false, ...(pages.retryAfterMs !== undefined ? { retryAfterMs: pages.retryAfterMs } : {}) }
    }

    if (pages.truncated) {
      // More reply chains are in range than one cycle reads (MAX_PAGES pages, newest first). The older ones are not
      // read, and the watermark below still moves past them: reading again from the newest would only return the same
      // chains. Recorded every time, so the loss is visible; one stderr line per channel.
      deps.decisions?.record({
        decision: 'backfill_skip',
        reason: 'page_limit',
        source: 'backfill',
        conversationId: entry.conversationId,
      })
      if (!notedPageLimit.has(channelId)) {
        notedPageLimit.add(channelId)
        process.stderr.write(
          `teams channel: backfill read ${MAX_PAGES} pages for one channel and stopped; older messages in that window were skipped\n`,
        )
      }
    }

    // Every human message in the window, oldest first (roots and their inline replies together).
    const candidates: Array<{ m: GraphChatMessage; rootId: string; createdMs: number; id: string }> = []
    const consider = (m: GraphChatMessage, rootId: string): void => {
      if (m.messageType !== 'message') return
      if (m.deletedDateTime !== null && m.deletedDateTime !== undefined) return
      const userId = m.from?.user?.id
      // from.application is a bot (this one included): its messages are not people talking to the bot.
      if (typeof userId !== 'string' || userId.length === 0) return
      const created = timeMs(m.createdDateTime)
      if (created === undefined || !(created > lower && created <= upper)) return
      candidates.push({ m, rootId, createdMs: created, id: typeof m.id === 'string' ? m.id : '' })
    }
    for (const root of pages.roots) {
      const rootId = typeof root.id === 'string' ? root.id : ''
      if (rootId.length === 0) continue
      consider(root, rootId)
      for (const reply of Array.isArray(root.replies) ? root.replies : []) consider(reply, rootId)
    }
    candidates.sort((a, b) => (a.createdMs !== b.createdMs ? a.createdMs - b.createdMs : a.id < b.id ? -1 : a.id > b.id ? 1 : 0))

    for (const c of candidates) {
      // stop() while a cycle runs: nothing more is submitted, and the watermark stays where it was, so the next start
      // tries the rest again (the seen ids drop what was already handed over).
      if (stopped) return { ok: true }
      const activity = activityFromGraph({
        channelId,
        groupId,
        tenantId: deps.tenantId,
        serviceUrl: entry.serviceUrl,
        botId: entry.botId,
        botName: entry.botName,
        botAppId: deps.botAppId,
        rootId: c.rootId,
        message: c.m,
      })
      const conversationId = activity.conversation?.id
      deps.decisions?.record({ decision: 'backfill_submit', reason: 'submit', source: 'backfill', messageId: c.id, conversationId })
      try {
        await deps.submit(activity)
      } catch (err) {
        // Keyed by channel and id, like the seen-id store: a message id is only unique within its conversation.
        const failureKey = `${channelId}\u0000${c.id}`
        const count = (submitFailures.get(failureKey) ?? 0) + 1
        submitFailures.set(failureKey, count)
        process.stderr.write(`teams channel: backfill submit failed (${errMsg(err)})\n`)
        if (count >= MAX_SUBMIT_FAILURES) {
          // Given up on: let the watermark pass it instead of retrying forever.
          deps.decisions?.record({
            decision: 'backfill_skip',
            reason: 'submit_failed_3x',
            source: 'backfill',
            messageId: c.id,
            conversationId,
          })
          continue
        }
        // Back to just before this message, so the next cycle tries it (and the ones after it) again.
        touch(state.watermarks, channelId, c.createdMs - 1)
        saveState()
        return { ok: true }
      }
    }
    touch(state.watermarks, channelId, upper)
    saveState()
    return { ok: true }
  }

  // ── one cycle ─────────────────────────────────────────────────────────────────────────────────────────────

  async function cycle(): Promise<CycleOutcome> {
    const nowAtStart = now()
    const upper = nowAtStart - deps.graceMs
    for (const entry of deps.refs.channelThreadRefs()) {
      // stop() while a cycle runs: no further channel is polled.
      if (stopped) return { ok: true }
      const r = await runChannel(entry, nowAtStart, upper)
      if (!r.ok) {
        // A cycle that did not finish leaves the watermarks of the channels it never reached where they were,
        // so nothing is skipped by a failed poll.
        return r
      }
    }
    return { ok: true }
  }

  // ── scheduling ────────────────────────────────────────────────────────────────────────────────────────────

  const clearTimer = deps.clearTimer ?? ((t: unknown) => clearTimeout(t as ReturnType<typeof setTimeout>))
  const setTimer =
    deps.setTimer ??
    ((fn: () => void, ms: number) => {
      const t = setTimeout(fn, ms)
      // The poll must never keep the process alive on its own.
      if (typeof (t as { unref?: () => void }).unref === 'function') (t as { unref: () => void }).unref()
      return t
    })

  function schedule(ms: number): void {
    if (stopped) return
    // The timer callback returns the cycle's promise, so a test seam can await it.
    timer = setTimer(() => tick(), ms)
  }

  async function tick(): Promise<void> {
    const r = await poll()
    if (stopped) return
    if (r.ok) {
      failures = 0
      schedule(deps.intervalMs)
      return
    }
    failures += 1
    // Never below two intervals: with a long interval a failure must not make the poll run more often than a success.
    const cap = Math.max(MAX_BACKOFF_MS, 2 * deps.intervalMs)
    const backoff = Math.min(deps.intervalMs * 2 ** failures, cap)
    // A 429 that said how long to wait is waited out, even when the plain backoff would be shorter.
    schedule(Math.max(backoff, r.retryAfterMs ?? 0))
  }

  /** One cycle, guarded so a second caller while one is running does nothing and fetches nothing. Never throws:
   *  the poll runs from a timer, so a surprise would otherwise become an unhandled rejection. */
  async function poll(): Promise<CycleOutcome> {
    if (running) return { ok: true }
    running = true
    try {
      return await cycle()
    } catch (err) {
      process.stderr.write(`teams channel: backfill cycle failed (${errMsg(err)})\n`)
      return { ok: false }
    } finally {
      running = false
    }
  }

  return {
    async runOnce() {
      await poll()
    },
    start() {
      schedule(deps.intervalMs)
    },
    stop() {
      stopped = true
      clearTimer(timer)
    },
  }
}
