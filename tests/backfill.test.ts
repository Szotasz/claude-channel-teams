/**
 * Graph backfill tests: the activity built from a Graph message, the candidate window, paging, the per-channel
 * (403) and whole-cycle (network / 5xx / 429) failure handling, scheduling, and — through makeTurnHandler, with
 * the fixtures of tests/inbound-dedupe.test.ts and tests/gate.test.ts — that a live and a backfilled delivery of
 * the same message run through the same gates and the same dedup store, in either order and across a restart.
 */

import { afterAll, describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { type Activity } from 'botbuilder'

import { activityFromGraph, createBackfill, type GraphChatMessage } from '../src/teams/backfill'
import { channelEventMeta, makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createDeliveredMessages } from '../src/teams/deliveredMessages'
import { createInboundIds } from '../src/teams/inboundIds'
import { createDecisionLog } from '../src/teams/decisionLog'
import type { FetchFn } from '../src/teams/threadHistory'

const tmpDirs: string[] = []
function tempDir(prefix: string): string {
  const dir = mkdtempSync(join(tmpdir(), prefix))
  tmpDirs.push(dir)
  return dir
}
afterAll(() => {
  for (const d of tmpDirs) rmSync(d, { recursive: true, force: true })
})

const J = (b: unknown, status = 200, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(b), { status, headers: { 'content-type': 'application/json', ...headers } })

type Call = { url: string; init?: RequestInit }
const G = 'https://graph.microsoft.com/v1.0/'
const CHANNEL = '19:chan@thread.tacv2'
const TEAM = 'group-1'
const T0 = Date.parse('2026-06-01T12:00:00Z')
const iso = (ms: number) => new Date(ms).toISOString()
const USER = '00000000-0000-4000-8000-000000000001'
const TENANT = '00000000-0000-4000-8000-000000000002'

function gmsg(id: string, at: number, extra: Record<string, unknown> = {}): GraphChatMessage {
  return {
    id,
    replyToId: null,
    createdDateTime: iso(at),
    lastModifiedDateTime: iso(at),
    messageType: 'message',
    deletedDateTime: null,
    from: { user: { id: USER, displayName: 'Ann' } },
    body: { contentType: 'text', content: `text ${id}` },
    ...extra,
  } as GraphChatMessage
}

interface ThreadRef { channelId: string; conversationId: string; serviceUrl: string; botId: string; botName?: string; lastSeen: number }
const ref = (channelId: string, lastSeen = T0): ThreadRef => ({
  channelId,
  conversationId: `${channelId};messageid=1700000000000`,
  serviceUrl: 'https://smba.trafficmanager.net/emea/',
  botId: 'bot-mri',
  lastSeen,
})
const fakeRefs = (list: ThreadRef[]) => ({ channelThreadRefs: () => [...list] })

const lines = (file: string) =>
  readFileSync(file, 'utf8').split('\n').filter(l => l.length > 0).map(l => JSON.parse(l) as Record<string, unknown>)

interface SetupOpts {
  handler?: (url: string, init?: RequestInit) => Response | Promise<Response>
  refs?: ThreadRef[]
  teams?: Record<string, string>
  submit?: (activity: Partial<Activity>) => Promise<void>
  graceMs?: number
  intervalMs?: number
  maxAgeMs?: number
  now?: number
}
function setup(o: SetupOpts = {}) {
  const dir = tempDir('cct-backfill-')
  const calls: Call[] = []
  const fetchFn: FetchFn = async (url, init) => {
    calls.push({ url, init })
    return o.handler ? await o.handler(url, init) : J({ value: [] })
  }
  const submitted: Partial<Activity>[] = []
  const teams = new Map(Object.entries(o.teams ?? { [CHANNEL]: TEAM }))
  let nowMs = o.now ?? T0
  const delays: number[] = []
  const cleared: unknown[] = []
  let timerFn: () => unknown = () => {}
  const decisionFile = join(dir, 'decisions.log')
  const stateFile = join(dir, 'backfill-state.json')
  const b = createBackfill({
    getToken: async () => 'TOKEN',
    fetch: fetchFn,
    now: () => nowMs,
    refs: fakeRefs(o.refs ?? [ref(CHANNEL)]),
    teamGroupFor: id => teams.get(id),
    botAppId: 'app-1',
    tenantId: 'tenant-1',
    stateFile,
    intervalMs: o.intervalMs ?? 180_000,
    graceMs: o.graceMs ?? 90_000,
    maxAgeMs: o.maxAgeMs ?? 7_200_000,
    submit: o.submit ?? (async a => { submitted.push(a) }),
    decisions: createDecisionLog({ file: decisionFile, maxBytes: 1 << 20 }),
    setTimer: (fn, ms) => {
      delays.push(ms)
      timerFn = fn
      return delays.length
    },
    clearTimer: t => { cleared.push(t) },
  })
  return {
    b,
    calls,
    submitted,
    stateFile,
    decisionFile,
    delays,
    cleared,
    fire: () => Promise.resolve(timerFn()),
    advance: (ms: number) => { nowMs += ms },
    watermarks: () => (JSON.parse(readFileSync(stateFile, 'utf8')) as { watermarks: Record<string, number> }).watermarks,
  }
}

describe('activityFromGraph', () => {
  const base = {
    channelId: CHANNEL,
    groupId: TEAM,
    tenantId: TENANT,
    serviceUrl: 'https://smba.trafficmanager.net/emea/',
    botId: 'bot-mri',
    botAppId: 'app-1',
    rootId: '1700000000000',
  }

  test('a root with a bot mention: the exact fields', () => {
    const a = activityFromGraph({
      ...base,
      message: {
        id: '1700000000001',
        replyToId: null,
        createdDateTime: iso(T0),
        messageType: 'message',
        deletedDateTime: null,
        from: { user: { id: USER, displayName: 'Ann' } },
        body: { contentType: 'text', content: 'hello' },
        mentions: [{ mentionText: 'Bot', mentioned: { application: { id: 'app-1' } } }],
      },
    })
    expect(a.type).toBe('message')
    expect(a.id).toBe('1700000000001')
    expect(a.timestamp).toBe(iso(T0))
    expect(a.channelId).toBe('msteams')
    expect(a.serviceUrl).toBe('https://smba.trafficmanager.net/emea/')
    expect(a.from).toEqual({ aadObjectId: USER, name: 'Ann' })
    expect(a.recipient).toEqual({ id: 'bot-mri' })
    expect(a.conversation).toEqual({
      id: `${CHANNEL};messageid=1700000000000`,
      conversationType: 'channel',
      isGroup: true,
      tenantId: TENANT,
    })
    expect(a.replyToId).toBeUndefined()
    expect(a.text).toBe('hello')
    expect(a.entities).toEqual([{ type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-mri', name: 'Bot' } }])
    expect(a.channelData).toEqual({
      tenant: { id: TENANT },
      team: { aadGroupId: TEAM },
      channel: { id: CHANNEL },
      backfilled: true,
    })
  })

  test('a reply: the root id in the conversation id and in replyToId', () => {
    const a = activityFromGraph({
      ...base,
      message: gmsg('1700000000002', T0, { replyToId: '1700000000000' }),
    })
    expect(a.conversation?.id).toBe(`${CHANNEL};messageid=1700000000000`)
    expect(a.replyToId).toBe('1700000000000')
  })

  test('html body: <at id="0">Bot</at> becomes <at>Bot</at> and the tag survives', () => {
    const a = activityFromGraph({
      ...base,
      message: gmsg('m', T0, {
        body: { contentType: 'html', content: '<div><at id="0">Bot</at> please summarise<br></div>' },
      }),
    })
    expect(a.text).toBe('<at>Bot</at> please summarise')
  })

  test('a text body is used as is, plus the attachment note when there are any', () => {
    const plain = activityFromGraph({ ...base, message: gmsg('m', T0) })
    expect(plain.text).toBe('text m')
    const withAttachments = activityFromGraph({
      ...base,
      message: gmsg('m', T0, { attachments: [{ id: 'a1' }], hostedContents: [{ id: 'h1' }] }),
    })
    expect(withAttachments.text).toBe('text m\n[backfilled: 2 attachment(s) not included]')
  })

  test('a mention of another person carries their id; unknown mentions are skipped', () => {
    const a = activityFromGraph({
      ...base,
      message: gmsg('m', T0, {
        mentions: [
          { mentionText: 'Bob', mentioned: { user: { id: '00000000-0000-4000-8000-000000000009', displayName: 'Bob' } } },
          { mentionText: 'tag', mentioned: { conversation: { id: 'x' } } },
        ],
      }),
    })
    expect(a.entities).toEqual([
      { type: 'mention', text: '<at>Bob</at>', mentioned: { id: '00000000-0000-4000-8000-000000000009', name: 'Bob' } },
    ])
  })

  test('the bot name from the stored reference becomes recipient.name, an empty one is left out', () => {
    const message = {
      id: '1700000000001',
      replyToId: null,
      createdDateTime: iso(T0),
      messageType: 'message',
      deletedDateTime: null,
      from: { user: { id: USER, displayName: 'Ann' } },
      body: { contentType: 'text', content: 'hello' },
    } as GraphChatMessage
    expect(activityFromGraph({ ...base, botName: 'Bot', message }).recipient).toEqual({ id: 'bot-mri', name: 'Bot' })
    expect(activityFromGraph({ ...base, botName: '', message }).recipient).toEqual({ id: 'bot-mri' })
  })
})

describe('the candidate window', () => {
  test('first sight sets the watermark to now - grace, fetches nothing, delivers nothing', async () => {
    const s = setup({ handler: () => J({ value: [gmsg('m', T0)] }) })
    await s.b.runOnce()
    expect(s.calls.length).toBe(0)
    expect(s.submitted.length).toBe(0)
    expect(s.watermarks()[CHANNEL]).toBe(T0 - 90_000)
    // The next cycle delivers a message created after that watermark.
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.submitted.map(a => a.id)).toEqual(['m'])
  })

  test('bot, system, deleted, stale and too-fresh messages are not candidates', async () => {
    const s = setup({
      handler: () =>
        J({
          value: [
            gmsg('ok', T0),
            gmsg('fresh', T0 + 100_000),
            gmsg('stale', T0 - 95_000),
            gmsg('bot', T0, { from: { application: { id: 'app-1', displayName: 'Bot' } } }),
            gmsg('sys', T0, { messageType: 'systemEventMessage' }),
            gmsg('del', T0, { deletedDateTime: iso(T0) }),
          ],
        }),
    })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.submitted.map(a => a.id)).toEqual(['ok'])
    // Everything the poll sends is marked as backfilled.
    expect(s.submitted[0]!.channelData).toMatchObject({ backfilled: true })
    expect(s.calls[0]!.url).toContain(`/teams/${TEAM}/channels/${encodeURIComponent(CHANNEL)}/messages?$top=50&$expand=replies`)
    expect((s.calls[0]!.init!.headers as Record<string, string>).Authorization).toBe('Bearer TOKEN')
  })

  test('a message older than now - maxAge is never delivered, even above the watermark', async () => {
    const s = setup({
      maxAgeMs: 300_000,
      handler: () => J({ value: [gmsg('recent', T0 + 400_000), gmsg('too-old', T0 + 200_000)] }),
    })
    await s.b.runOnce()
    s.advance(600_000)
    await s.b.runOnce()
    expect(s.submitted.map(a => a.id)).toEqual(['recent'])
  })

  test('replies of a root are delivered too, oldest first', async () => {
    const root = gmsg('r', T0, { replies: [gmsg('a2', T0 + 20_000, { replyToId: 'r' }), gmsg('a1', T0 + 10_000, { replyToId: 'r' })] })
    const s = setup({ handler: () => J({ value: [root] }) })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.submitted.map(a => a.id)).toEqual(['r', 'a1', 'a2'])
    expect(s.submitted[1]!.replyToId).toBe('r')
  })

  test('the poll hands the stored bot name to the activity it submits', async () => {
    const s = setup({ refs: [{ ...ref(CHANNEL), botName: 'Bot' }], handler: () => J({ value: [gmsg('m1', T0)] }) })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.submitted.map(a => a.recipient)).toEqual([{ id: 'bot-mri', name: 'Bot' }])
  })
})

describe('paging', () => {
  test('follows a v1 next link', async () => {
    const s = setup({
      handler: url => (url === `${G}next` ? J({ value: [gmsg('p2', T0)] }) : J({ value: [gmsg('p1', T0)], '@odata.nextLink': `${G}next` })),
    })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.calls.length).toBe(2)
    expect(s.submitted.map(a => a.id)).toEqual(['p1', 'p2'])
  })

  test('refuses a non-v1 next link (beta, a foreign host, dot segments)', async () => {
    for (const link of ['https://graph.microsoft.com/beta/teams/x', 'https://evil.example/x', `${G}%2e%2e/beta/teams/x`]) {
      const s = setup({ handler: () => J({ value: [gmsg('p1', T0)], '@odata.nextLink': link }) })
      await s.b.runOnce()
      s.advance(180_000)
      await s.b.runOnce()
      expect(s.calls.length).toBe(1)
      expect(s.submitted.map(a => a.id)).toEqual(['p1'])
      expect(s.calls.some(c => c.url.includes('evil') || c.url.includes('beta'))).toBe(false)
    }
  })

  test('stops when the last root of the page is older than the window', async () => {
    const s = setup({
      handler: () => J({ value: [gmsg('new', T0), gmsg('old', T0 - 500_000)], '@odata.nextLink': `${G}next` }),
    })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.calls.length).toBe(1)
    expect(s.submitted.map(a => a.id)).toEqual(['new'])
  })

  test('keeps going while the last root has a reply newer than the window floor', async () => {
    const page = gmsg('new', T0 - 500_000, { replies: [gmsg('fresh-reply', T0 + 10_000, { replyToId: 'new' })] })
    const s = setup({
      handler: url => (url === `${G}next` ? J({ value: [] }) : J({ value: [page], '@odata.nextLink': `${G}next` })),
    })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.calls.length).toBe(2)
    expect(s.submitted.map(a => a.id)).toEqual(['fresh-reply'])
  })

  test('stops at five pages', async () => {
    let n = 0
    const s = setup({ handler: () => J({ value: [gmsg(`p${++n}`, T0)], '@odata.nextLink': `${G}page` }) })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.calls.length).toBe(5)
  })

  test('hitting the page limit is recorded as a skip, and the watermark still moves on', async () => {
    let n = 0
    const s = setup({ handler: () => J({ value: [gmsg(`p${++n}`, T0)], '@odata.nextLink': `${G}page` }) })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    const skips = lines(s.decisionFile).filter(l => l.decision === 'backfill_skip' && l.reason === 'page_limit')
    expect(skips).toHaveLength(1)
    expect(skips[0]).toMatchObject({ source: 'backfill', channel: CHANNEL })
    expect(s.watermarks()[CHANNEL]).toBe(T0 + 180_000 - 90_000)
  })

  test('five pages that end on their own are not a page limit', async () => {
    let n = 0
    const s = setup({
      handler: () => {
        n += 1
        return J({ value: [gmsg(`p${n}`, T0)], ...(n < 5 ? { '@odata.nextLink': `${G}page` } : {}) })
      },
    })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.calls.length).toBe(5)
    expect(s.submitted.length).toBe(5)
    expect(lines(s.decisionFile).filter(l => l.reason === 'page_limit')).toEqual([])
  })

  test('an empty page advances the watermark and submits nothing', async () => {
    const s = setup({ handler: () => J({ value: [] }) })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.submitted.length).toBe(0)
    expect(s.watermarks()[CHANNEL]).toBe(T0 + 180_000 - 90_000)
  })
})

describe('a 403 channel', () => {
  const A = '19:a@thread.tacv2'
  const B = '19:b@thread.tacv2'
  const setupTwo = () =>
    setup({
      refs: [ref(A), ref(B)],
      teams: { [A]: 'team-a', [B]: 'team-b' },
      handler: url => (url.includes(encodeURIComponent(A)) ? J({}, 403) : J({ value: [gmsg('in-b', T0)] })),
    })
  const fetchesOf = (s: ReturnType<typeof setupTwo>, channel: string) =>
    s.calls.filter(c => c.url.includes(encodeURIComponent(channel))).length

  test('is recorded once, left alone for six hours, other channels still polled', async () => {
    const s = setupTwo()
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.submitted.map(a => a.id)).toEqual(['in-b'])
    expect(fetchesOf(s, A)).toBe(1)
    expect(fetchesOf(s, B)).toBe(1)
    expect(lines(s.decisionFile).filter(l => l.decision === 'backfill_error' && l.reason === 'forbidden')).toHaveLength(1)
    // Still inside the six hours: A is not fetched again, B is.
    s.advance(3_600_000)
    await s.b.runOnce()
    expect(fetchesOf(s, A)).toBe(1)
    expect(fetchesOf(s, B)).toBe(2)
    expect(lines(s.decisionFile).filter(l => l.decision === 'backfill_error' && l.reason === 'forbidden')).toHaveLength(1)
    // After the six hours A is tried again; it fails again, and that second 403 is not recorded a second time.
    s.advance(6 * 3_600_000)
    await s.b.runOnce()
    expect(fetchesOf(s, A)).toBe(2)
    expect(lines(s.decisionFile).filter(l => l.decision === 'backfill_error' && l.reason === 'forbidden')).toHaveLength(1)
  })

  test('the six-hour back-off is written to the state file at once', async () => {
    const s = setup({ refs: [ref(A)], teams: { [A]: 'team-a' }, handler: () => J({}, 403) })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    const state = JSON.parse(readFileSync(s.stateFile, 'utf8')) as { forbiddenUntil: Record<string, number> }
    expect(state.forbiddenUntil[A]).toBe(T0 + 180_000 + 6 * 3_600_000)
  })
})

describe('a failing cycle', () => {
  test('a network error stops the cycle and leaves the watermarks alone', async () => {
    const s = setup({ handler: () => J({ value: [] }) })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    const before = s.watermarks()[CHANNEL]
    expect(before).toBe(T0 + 180_000 - 90_000)
    const failing = setup({ handler: () => { throw new Error('ECONNRESET') } })
    await failing.b.runOnce()
    failing.advance(180_000)
    await failing.b.runOnce()
    expect(failing.calls.length).toBe(1)
    expect(failing.submitted.length).toBe(0)
    expect(failing.watermarks()[CHANNEL]).toBe(T0 - 90_000)
  })

  test('a 5xx stops the cycle and is recorded', async () => {
    const s = setup({ handler: () => J({}, 500) })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce()
    expect(s.calls.length).toBe(1)
    expect(s.submitted.length).toBe(0)
    expect(lines(s.decisionFile).filter(l => l.decision === 'backfill_error' && l.reason === 'http_5xx')).toHaveLength(1)
  })

  test('a 429 with Retry-After is waited out, the backoff doubles and is capped, a success resets it', async () => {
    let mode: 'throttled' | 'ok' | 'error' = 'throttled'
    const s = setup({
      intervalMs: 10_000,
      handler: () => (mode === 'throttled' ? J({}, 429, { 'Retry-After': '120' }) : mode === 'error' ? J({}, 500) : J({ value: [] })),
    })
    s.b.start()
    expect(s.delays).toEqual([10_000])
    await s.fire() // first sight: no fetch, a successful cycle
    expect(s.delays).toEqual([10_000, 10_000])
    s.advance(10_000)
    await s.fire() // throttled: at least the Retry-After, and the doubled interval
    expect(s.delays[2]).toBeGreaterThanOrEqual(120_000)
    mode = 'ok'
    s.advance(s.delays[2]!)
    await s.fire() // success: back to the plain interval
    expect(s.delays[3]).toBe(10_000)
    expect(lines(s.decisionFile).filter(l => l.decision === 'backfill_error' && l.reason === 'throttled')).toHaveLength(1)
    // The success reset the failure count: the next plain failure backs off by exactly twice the interval.
    mode = 'error'
    s.advance(s.delays[3]!)
    await s.fire()
    expect(s.delays[4]).toBe(20_000)
  })

  test('the backoff never exceeds thirty minutes', async () => {
    const s = setup({ intervalMs: 60_000, handler: () => J({}, 500) })
    s.b.start()
    for (let i = 0; i < 7; i++) await s.fire()
    // The first cycle is the first sight (a success), the other six all fail.
    expect(s.delays.slice(1)).toEqual([60_000, 120_000, 240_000, 480_000, 960_000, 1_800_000, 1_800_000])
  })

  test('with a long interval a failure never makes the poll run more often than a success', async () => {
    const s = setup({ intervalMs: 3_600_000, handler: () => J({}, 500) })
    s.b.start()
    for (let i = 0; i < 4; i++) await s.fire()
    // The first cycle is the first sight (a success), the other three all fail: two intervals, and capped there.
    expect(s.delays.slice(1)).toEqual([3_600_000, 7_200_000, 7_200_000, 7_200_000])
  })

  test('stop() clears the timer and nothing is scheduled afterwards', async () => {
    const s = setup({ handler: () => J({ value: [] }) })
    s.b.start()
    expect(s.delays).toEqual([180_000])
    s.b.stop()
    expect(s.cleared).toEqual([1])
    await s.fire()
    expect(s.delays).toEqual([180_000])
  })

  test('start() after stop() schedules nothing', () => {
    const s = setup({ handler: () => J({ value: [] }) })
    s.b.stop()
    s.b.start()
    expect(s.delays).toEqual([])
  })

  test('stop() during a cycle: nothing more is submitted, no other channel is polled, the watermark stays', async () => {
    const A = '19:a@thread.tacv2'
    const B = '19:b@thread.tacv2'
    let stopNow: () => void = () => {}
    const got: string[] = []
    const s = setup({
      refs: [ref(A), ref(B)],
      teams: { [A]: 'team-a', [B]: 'team-b' },
      handler: url =>
        J({ value: url.includes(encodeURIComponent(A)) ? [gmsg('a1', T0), gmsg('a2', T0 + 1000)] : [gmsg('b1', T0)] }),
      submit: async a => {
        got.push(String(a.id))
        stopNow()
      },
    })
    stopNow = () => s.b.stop()
    await s.b.runOnce()
    const before = s.watermarks()
    s.advance(180_000)
    await s.b.runOnce()
    expect(got).toEqual(['a1'])
    expect(s.calls.filter(c => c.url.includes(encodeURIComponent(B))).length).toBe(0)
    expect(s.watermarks()[A]).toBe(before[A]!)
  })
})

describe('a failing submit', () => {
  test('is retried the next cycle and given up on after three failures', async () => {
    const attempts: string[] = []
    const s = setup({
      handler: () => J({ value: [gmsg('m1', T0)] }),
      submit: async a => {
        attempts.push(a.id!)
        throw new Error('the handler exploded')
      },
    })
    await s.b.runOnce()
    s.advance(180_000)
    await s.b.runOnce() // failure 1: the watermark moves to just before the message
    expect(attempts).toEqual(['m1'])
    expect(s.watermarks()[CHANNEL]).toBe(T0 - 1)
    s.advance(180_000)
    await s.b.runOnce() // failure 2
    expect(attempts).toEqual(['m1', 'm1'])
    expect(s.watermarks()[CHANNEL]).toBe(T0 - 1)
    s.advance(180_000)
    await s.b.runOnce() // failure 3: given up on, the watermark passes it
    expect(attempts).toEqual(['m1', 'm1', 'm1'])
    expect(s.watermarks()[CHANNEL]).toBe(T0 + 3 * 180_000 - 90_000)
    expect(lines(s.decisionFile).filter(l => l.decision === 'backfill_skip' && l.reason === 'submit_failed_3x')).toHaveLength(1)
    s.advance(180_000)
    await s.b.runOnce()
    expect(attempts.length).toBe(3)
  })

  test('the failure count is per channel: the same message id in two channels gets three tries in each', async () => {
    const A = '19:a@thread.tacv2'
    const B = '19:b@thread.tacv2'
    const s = setup({
      refs: [ref(A), ref(B)],
      teams: { [A]: 'team-a', [B]: 'team-b' },
      handler: () => J({ value: [gmsg('m1', T0)] }),
      submit: async () => {
        throw new Error('the handler exploded')
      },
    })
    const givenUp = () =>
      lines(s.decisionFile).filter(l => l.decision === 'backfill_skip' && l.reason === 'submit_failed_3x')
    await s.b.runOnce() // first sight of both channels
    s.advance(180_000)
    await s.b.runOnce() // failure 1 in A and in B
    s.advance(180_000)
    await s.b.runOnce() // failure 2 in A and in B: neither is given up on yet
    expect(givenUp()).toHaveLength(0)
    expect(s.watermarks()[A]).toBe(T0 - 1)
    expect(s.watermarks()[B]).toBe(T0 - 1)
    s.advance(180_000)
    await s.b.runOnce() // failure 3 in each: both are given up on
    expect(givenUp()).toHaveLength(2)
  })
})

describe('no overlap', () => {
  test('runOnce() while a cycle is running returns without fetching again', async () => {
    let release!: () => void
    const gate = new Promise<void>(r => { release = r })
    const s = setup({ handler: async () => { await gate; return J({ value: [] }) } })
    await s.b.runOnce() // the first sight: the watermark is set, nothing is fetched
    s.advance(180_000)
    const first = s.b.runOnce()
    await new Promise(r => setTimeout(r, 0))
    expect(s.calls.length).toBe(1)
    await s.b.runOnce()
    expect(s.calls.length).toBe(1)
    release()
    await first
  })
})

// ── the same gates and the same dedup as a live message ────────────────────────────────────────────────────

const MENTION = { type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-mri', name: 'Bot' } }
const THREAD = `${CHANNEL};messageid=1700000000000`

function liveActivity(o: Record<string, unknown> = {}): Activity {
  return {
    type: 'message',
    id: '1700000000001',
    timestamp: new Date(T0),
    text: '<at>Bot</at> hi',
    from: { id: 'user-1', name: 'Test User', aadObjectId: USER },
    conversation: { id: THREAD, tenantId: TENANT, conversationType: 'channel', isGroup: true },
    recipient: { id: 'bot-mri', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/emea/',
    entities: [MENTION],
    ...o,
  } as unknown as Activity
}

function backfilledActivity(o: Record<string, unknown> = {}): Partial<Activity> {
  return activityFromGraph({
    channelId: CHANNEL,
    groupId: TEAM,
    tenantId: TENANT,
    serviceUrl: 'https://smba.trafficmanager.net/emea/',
    botId: 'bot-mri',
    botAppId: 'app-1',
    rootId: '1700000000000',
    message: {
      id: '1700000000001',
      replyToId: null,
      createdDateTime: iso(T0),
      messageType: 'message',
      deletedDateTime: null,
      from: { user: { id: USER, displayName: 'Test User' } },
      body: { contentType: 'text', content: '<at>Bot</at> hi' },
      mentions: [{ mentionText: 'Bot', mentioned: { application: { id: 'app-1' } } }],
      ...o,
    },
  })
}

interface TurnSetupOpts { idsFile?: string; unmentioned?: string[]; allow?: boolean }
function turnSetup(opts: TurnSetupOpts = {}) {
  const dir = tempDir('cct-backfill-turn-')
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  if (opts.allow !== false) allowlist.addEntry(USER)
  const events: any[] = []
  const known: string[] = []
  const refs = createConversationRefStore()
  const handler = makeTurnHandler({
    config: {
      stateDir: dir,
      allowlistFile: join(dir, 'allowlist.json'),
      appId: 'app-1',
      appPassword: 'p',
      appType: 'SingleTenant',
      tenantId: TENANT,
      bindHost: '127.0.0.1',
      port: 0,
      ...(opts.unmentioned ? { unmentionedConversations: new Set(opts.unmentioned) } : {}),
    } as any,
    adapter: {} as any,
    allowlist,
    refs,
    onEvent: e => events.push(e),
    deliveredMessages: createDeliveredMessages(),
    inboundIds: createInboundIds(opts.idsFile ? { file: opts.idsFile } : {}),
    knownUsers: { set: (_aad, mri) => known.push(mri) },
  } as any)
  return { run: (a: Partial<Activity>) => handler({ activity: a } as any), events, refs, known, dir }
}

describe('live and backfill share one dedup and one gate', () => {
  test('live then backfill: one event', async () => {
    const s = turnSetup()
    await s.run(liveActivity())
    await s.run(backfilledActivity())
    expect(s.events.length).toBe(1)
    expect(s.events[0]!.backfilled).toBeUndefined()
  })

  test('backfill then live: one event, and the event says it was recovered', async () => {
    const s = turnSetup()
    await s.run(backfilledActivity())
    await s.run(liveActivity())
    expect(s.events.length).toBe(1)
    expect(s.events[0]!.backfilled).toBe(true)
    expect(channelEventMeta(s.events[0]!).backfilled).toBe('true')
    expect(s.events[0]!.text).toBe('hi')
  })

  test('live, a restart (the ids reloaded from the file), then backfill: still one delivery', async () => {
    const dir = tempDir('cct-backfill-restart-')
    const idsFile = join(dir, 'inbound-ids.json')
    const live = turnSetup({ idsFile })
    await live.run(liveActivity())
    expect(live.events.length).toBe(1)
    const restarted = turnSetup({ idsFile })
    await restarted.run(backfilledActivity())
    expect(restarted.events.length).toBe(0)
  })

  test('the same outcome for every gate combination, live and backfilled', async () => {
    for (const allow of [true, false]) {
      for (const mentioned of [true, false]) {
        for (const unmentioned of [undefined, ['*']]) {
          const live = turnSetup({ allow, unmentioned })
          const back = turnSetup({ allow, unmentioned })
          await live.run(liveActivity(mentioned ? {} : { entities: [], text: 'hi' }))
          await back.run(
            backfilledActivity(mentioned ? {} : { body: { contentType: 'text', content: 'hi' }, mentions: [] }),
          )
          // allowlisted / mentioned / opted-in unmentioned -> the live count and the backfilled count.
          expect([allow, mentioned, unmentioned?.join(',') ?? 'unset', live.events.length]).toEqual([
            allow, mentioned, unmentioned?.join(',') ?? 'unset', back.events.length,
          ])
        }
      }
    }
  })

  test('a backfilled turn keeps the stored reference routable and learns nothing about the sender', async () => {
    const s = turnSetup()
    await s.run(backfilledActivity())
    const stored = s.refs.get(THREAD)
    expect(stored?.ref.serviceUrl).toBe('https://smba.trafficmanager.net/emea/')
    expect(stored?.ref.bot?.id).toBe('bot-mri')
    expect(stored?.ref.conversation?.id).toBe(THREAD)
    expect(s.known).toEqual([])
  })

  test('the listener hands a backfilled activity to the same handler as the HTTP path', () => {
    const src = readFileSync(join(import.meta.dir, '../src/teams/adapter.ts'), 'utf8')
    expect(src).toContain('const handler = makeTurnHandler(deps)\n  const turn = makeEarlyAckTurn(handler)')
    expect(src).toContain('    submitActivity(activity) {\n      return handler({ activity })\n    },')
  })
})

describe('ConversationRefStore.channelThreadRefs', () => {
  test('one entry per channel, the most recently seen ref, only complete refs', () => {
    let t = 0
    const store = createConversationRefStore({ now: () => t })
    const mk = (id: string, extra: Record<string, unknown> = {}) =>
      ({ conversation: { id }, serviceUrl: 'https://smba.trafficmanager.net/emea/', bot: { id: 'bot-mri' }, ...extra })
    t = 10
    store.put('19:a@thread.tacv2;messageid=1', mk('19:a@thread.tacv2;messageid=1') as any, 'aad-1')
    t = 20
    store.put('19:a@thread.tacv2;messageid=2', mk('19:a@thread.tacv2;messageid=2') as any, 'aad-1')
    t = 15
    store.put('19:b@thread.tacv2;messageid=1', mk('19:b@thread.tacv2;messageid=1', { serviceUrl: undefined }) as any, 'aad-1')
    t = 25
    store.put('19:c@thread.tacv2;messageid=1', mk('19:c@thread.tacv2;messageid=1', { bot: {} }) as any, 'aad-1')
    t = 30
    store.put('19:g@thread.v2;messageid=1', mk('19:g@thread.v2;messageid=1') as any, 'aad-1')
    t = 35
    store.put('a:personal1', mk('a:personal1') as any, 'aad-1')
    t = 40
    store.put('19:plain@thread.tacv2', mk('19:plain@thread.tacv2') as any, 'aad-1')
    const refs = store.channelThreadRefs()
    expect(refs.map(r => r.channelId)).toEqual(['19:a@thread.tacv2', '19:plain@thread.tacv2'])
    expect(refs[0]).toEqual({
      channelId: '19:a@thread.tacv2',
      conversationId: '19:a@thread.tacv2;messageid=2',
      serviceUrl: 'https://smba.trafficmanager.net/emea/',
      botId: 'bot-mri',
      lastSeen: 20,
    })
  })

  test('the bot name is carried when the reference has one', () => {
    let t = 0
    const store = createConversationRefStore({ now: () => t })
    const mk = (id: string, bot: Record<string, unknown>) =>
      ({ conversation: { id }, serviceUrl: 'https://smba.trafficmanager.net/emea/', bot })
    t = 10
    store.put('19:a@thread.tacv2;messageid=1', mk('19:a@thread.tacv2;messageid=1', { id: 'bot-mri', name: 'Bot' }) as any, 'aad-1')
    t = 20
    store.put('19:b@thread.tacv2;messageid=1', mk('19:b@thread.tacv2;messageid=1', { id: 'bot-mri' }) as any, 'aad-1')
    const refs = store.channelThreadRefs()
    expect(refs.map(r => r.channelId)).toEqual(['19:a@thread.tacv2', '19:b@thread.tacv2'])
    expect(refs[0]!.botName).toBe('Bot')
    expect('botName' in refs[1]!).toBe(false)
  })
})
