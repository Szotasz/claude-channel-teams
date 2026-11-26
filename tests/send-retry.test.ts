import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
// Imported only to measure what the SDK itself retries (TEAMSSDK1126: nothing).
import { ConnectorClient, HttpClient } from '@microsoft/agents-hosting'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createReplySender } from '../src/teams/reply'
import { errorStatus, isPluginRetryable, retryAfterMs, withSendRetry } from '../src/teams/sendRetry'

function hdrs(rec: Record<string, string>) {
  return { get: (name: string) => rec[name.toLowerCase()] }
}
const err = (statusCode: unknown, headers?: Record<string, string>) =>
  Object.assign(new Error('x'), { statusCode, ...(headers ? { response: { headers: hdrs(headers) } } : {}) })

describe('retryAfterMs', () => {
  test('cases', () => {
    const now = () => 1_000_000
    expect(retryAfterMs(hdrs({}), now)).toBeUndefined()
    expect(retryAfterMs(hdrs({ 'retry-after': '0' }), now)).toBe(0)
    expect(retryAfterMs(hdrs({ 'retry-after': '1' }), now)).toBe(1000)
    expect(retryAfterMs(hdrs({ 'retry-after-ms': '250' }), now)).toBe(250)
    expect(retryAfterMs(hdrs({ 'x-ms-retry-after-ms': '7' }), now)).toBe(7)
    expect(retryAfterMs(hdrs({ 'retry-after-ms': '250', 'retry-after': '9' }), now)).toBe(250)
    expect(retryAfterMs(hdrs({ 'retry-after': 'soon' }), now)).toBeUndefined()
    expect(retryAfterMs(hdrs({ 'retry-after': new Date(1_003_000).toUTCString() }), now)).toBe(3000)
    expect(retryAfterMs(hdrs({ 'retry-after': new Date(0).toUTCString() }), now)).toBe(0)
    expect(retryAfterMs(hdrs({ 'retry-after': '' }), now)).toBeUndefined()
    expect(retryAfterMs(hdrs({ 'retry-after': 'Infinity' }), now)).toBeUndefined()
    expect(retryAfterMs(hdrs({ 'retry-after': '-1' }), now)).toBe(0)
    expect(retryAfterMs(undefined, now)).toBeUndefined()
    expect(retryAfterMs({ get: 5 }, now)).toBeUndefined()
    expect(retryAfterMs({ get: () => { throw new Error('boom') } }, now)).toBeUndefined()
    // The Agents SDK's HttpError carries a fetch Headers object.
    expect(retryAfterMs(new Headers({ 'Retry-After': '2' }), now)).toBe(2000)
  })
})

describe('isPluginRetryable', () => {
  test('cases', () => {
    for (const s of [408, 412, 429, 500, 502, 503, 504]) expect(isPluginRetryable({ statusCode: s })).toBe(true)
    for (const s of [400, 401, 403, 404, 409, 501, 505]) expect(isPluginRetryable({ statusCode: s })).toBe(false)
    // botframework-connector's RestError had statusCode, the Agents SDK's HttpError has status
    expect(isPluginRetryable({ status: 429 })).toBe(true)
    expect(isPluginRetryable({ status: 400 })).toBe(false)
    expect(errorStatus({ status: 503 })).toBe(503)
    expect(isPluginRetryable(err(429, { 'retry-after': '1' }))).toBe(true)
    expect(isPluginRetryable({ statusCode: '429' })).toBe(false)
    expect(isPluginRetryable(null)).toBe(false)
    expect(isPluginRetryable('x')).toBe(false)
    expect(isPluginRetryable(new Error('x'))).toBe(false)
    const thrower = { statusCode: 429, get response(): never { throw new Error('boom') } }
    expect(isPluginRetryable(thrower)).toBe(true)
  })
})

describe('withSendRetry', () => {
  function harness(results: Array<unknown>, opts: { random?: () => number; now?: () => number } = {}) {
    const sleeps: number[] = []
    const lines: string[] = []
    let calls = 0
    const fn = async () => {
      const r = results[Math.min(calls++, results.length - 1)]
      if (r instanceof Error) throw r
      return r
    }
    const deps = {
      sleep: async (ms: number) => { sleeps.push(ms) },
      random: opts.random ?? (() => 0),
      ...(opts.now ? { now: opts.now } : {}),
      log: (l: string) => { lines.push(l) },
    }
    return { fn, deps, sleeps, lines, calls: () => calls }
  }

  test('success first', async () => {
    const h = harness(['ok'])
    expect(await withSendRetry('reply', h.fn, h.deps)).toBe('ok')
    expect(h.calls()).toBe(1)
    expect(h.sleeps).toEqual([])
    expect(h.lines).toEqual([])
  })

  test('429 twice then success (random 0)', async () => {
    const h = harness([err(429), err(429), 'ok'])
    expect(await withSendRetry('reply', h.fn, h.deps)).toBe('ok')
    expect(h.calls()).toBe(3)
    expect(h.sleeps).toEqual([500, 1000])
    expect(h.lines.length).toBe(2)
    expect(h.lines[0]).toBe('teams channel: reply got HTTP 429, attempt 2/3 after 500 ms\n')
    expect(h.lines[1]).toBe('teams channel: reply got HTTP 429, attempt 3/3 after 1000 ms\n')
  })

  test('429 twice then success (random 1)', async () => {
    const h = harness([err(429), err(429), 'ok'], { random: () => 1 })
    await withSendRetry('reply', h.fn, h.deps)
    expect(h.sleeps).toEqual([1000, 2000])
  })

  test('412 always rejects with the same error', async () => {
    const e = err(412)
    const h = harness([e])
    await expect(withSendRetry('reply', h.fn, h.deps)).rejects.toBe(e)
    expect(h.calls()).toBe(3)
    expect(h.sleeps.length).toBe(2)
  })

  test('400 is not retried', async () => {
    const e = err(400)
    const h = harness([e])
    await expect(withSendRetry('reply', h.fn, h.deps)).rejects.toBe(e)
    expect(h.calls()).toBe(1)
    expect(h.sleeps).toEqual([])
  })

  // TEAMSSDK1126: the Agents SDK does not retry a throttling answer itself; its wait is honoured here.
  test('429 with a usable Retry-After waits at least that long', async () => {
    const h = harness([err(429, { 'retry-after': '2' }), 'ok'])
    expect(await withSendRetry('reply', h.fn, h.deps)).toBe('ok')
    expect(h.calls()).toBe(2)
    expect(h.sleeps).toEqual([2000])
  })

  test('a Retry-After past the budget is not waited for', async () => {
    const e = err(429, { 'retry-after': '30' })
    const h = harness([e])
    await expect(withSendRetry('reply', h.fn, h.deps)).rejects.toBe(e)
    expect(h.calls()).toBe(1)
    expect(h.sleeps).toEqual([])
  })

  test('5xx (not 501/505) and 408 are retried, an Agents SDK HttpError (status) too', async () => {
    for (const status of [500, 502, 503, 504, 408]) {
      const h = harness([Object.assign(new Error('x'), { status }), 'ok'])
      expect(await withSendRetry('reply', h.fn, h.deps)).toBe('ok')
      expect(h.calls()).toBe(2)
    }
    for (const status of [501, 505]) {
      const h = harness([Object.assign(new Error('x'), { status })])
      await expect(withSendRetry('reply', h.fn, h.deps)).rejects.toBeDefined()
      expect(h.calls()).toBe(1)
    }
  })

  test('budget exceeded at once', async () => {
    let n = 0
    const e = err(429)
    const h = harness([e], { now: () => (n++ === 0 ? 0 : 9_600) })
    await expect(withSendRetry('reply', h.fn, h.deps)).rejects.toBe(e)
    expect(h.calls()).toBe(1)
    expect(h.sleeps).toEqual([])
  })

  test('a late timer does not start an attempt after the budget', async () => {
    const e = err(429)
    let t = 0
    let calls = 0
    const sleeps: number[] = []
    // The first attempt ends at 9.4 s, so a 500 ms wait is allowed; the timer then fires at 10.2 s.
    const fn = async () => { calls++; t = 9_400; throw e }
    const lines: string[] = []
    const deps = { now: () => t, random: () => 0, log: (l: string) => { lines.push(l) }, sleep: async (ms: number) => { sleeps.push(ms); t = 10_200 } }
    await expect(withSendRetry('reply', fn, deps)).rejects.toBe(e)
    expect(calls).toBe(1)
    expect(sleeps).toEqual([500])
    // No line for an attempt that does not happen.
    expect(lines).toEqual([])
  })

  test('budget exceeded the second time', async () => {
    let n = 0
    const e = err(429)
    const h = harness([e], { now: () => (n++ === 0 ? 0 : 9_500) })
    await expect(withSendRetry('reply', h.fn, h.deps)).rejects.toBe(e)
    expect(h.calls()).toBe(2)
    expect(h.sleeps).toEqual([500])
  })
})

// The real Agents SDK connector client over a stubbed fetch: it retries nothing, so every attempt is the plugin's.
const realFetch = globalThis.fetch
let answers: Array<{ status: number; headers?: Record<string, string> }> = []
let fetchCalls = 0
beforeEach(() => { fetchCalls = 0 })
afterEach(() => { globalThis.fetch = realFetch })
function client(statuses: number[], headers: Record<string, string> = {}) {
  answers = statuses.map(status => ({ status, headers }))
  globalThis.fetch = (async () => {
    const a = answers[Math.min(fetchCalls, answers.length - 1)]!
    fetchCalls++
    const body = a.status === 200 ? '{"id":"m1"}' : '{"error":{"code":"x","message":"y"}}'
    return new Response(body, { status: a.status, headers: { 'content-type': 'application/json', ...(a.status === 200 ? {} : a.headers) } })
  }) as typeof fetch
  const c = new ConnectorClient(new HttpClient({ baseURL: 'https://smba.example.test', headers: { Authorization: 'Bearer t' } }))
  return { c, calls: () => fetchCalls }
}

describe('real SDK', () => {
  const deps = { sleep: async () => {}, log: () => {} }
  const msg = { type: 'message', text: 'hi' } as any

  test('429 without headers: the plugin retries, the SDK does not', async () => {
    const { c, calls } = client([429])
    await expect(withSendRetry('t', () => c.sendToConversation('a:1', msg), deps)).rejects.toMatchObject({ status: 429 })
    expect(calls()).toBe(3)
  })

  test('429 with retry-after 0: still three attempts in all (no SDK retries under them)', async () => {
    const { c, calls } = client([429], { 'retry-after': '0' })
    await expect(withSendRetry('t', () => c.sendToConversation('a:1', msg), deps)).rejects.toMatchObject({ status: 429 })
    expect(calls()).toBe(3)
  })

  test('503 then 200: retried by the plugin', async () => {
    const { c, calls } = client([503, 200])
    const r = await withSendRetry('t', () => c.sendToConversation('a:1', msg), deps)
    expect((r as any).id).toBe('m1')
    expect(calls()).toBe(2)
  })

  test('412 then 200', async () => {
    const { c, calls } = client([412, 200])
    const r = await withSendRetry('t', () => c.sendToConversation('a:1', msg), deps)
    expect((r as any).id).toBe('m1')
    expect(calls()).toBe(2)
  })

  test('400 is not retried', async () => {
    const { c, calls } = client([400])
    await expect(withSendRetry('t', () => c.sendToConversation('a:1', msg), deps)).rejects.toBeDefined()
    expect(calls()).toBe(1)
  })
})

describe('reply.ts behaviour', () => {
  const ID = '00000000-0000-4000-8000-000000000001'
  function setup(ctx: any) {
    const dir = mkdtempSync(join(tmpdir(), 'cct-sr-'))
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(ID)
    const refs = createConversationRefStore()
    refs.put('conv-abc', { conversation: { id: 'conv-abc' } } as any, ID)
    let continues = 0
    const adapter = {
      async continueConversation(_a: string, _r: any, logic: any) {
        continues++
        await logic(ctx)
      },
    } as any
    const config = { stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'app', appPassword: 'pw', appType: 'SingleTenant' as const, tenantId: 't', bindHost: '127.0.0.1', port: 0 }
    return { sender: createReplySender({ config, adapter, allowlist, refs }), continues: () => continues, dir }
  }

  test('sendReply retries a 429 inside one continueConversation', async () => {
    let sends = 0
    const s = setup({
      async sendActivity() {
        if (sends++ === 0) throw { statusCode: 429 }
        return { id: 'm9' }
      },
    })
    try {
      const t0 = Date.now()
      expect(await s.sender.sendReply('conv-abc', 'hi')).toBe('m9')
      // The real wait: at least 500 ms before the second attempt (a little slack for the timer).
      expect(Date.now() - t0).toBeGreaterThanOrEqual(450)
      expect(sends).toBe(2)
      expect(s.continues()).toBe(1)
    } finally {
      rmSync(s.dir, { recursive: true, force: true })
    }
  })

  test('editReply does not retry a 400', async () => {
    let updates = 0
    const s = setup({
      async updateActivity() {
        updates++
        throw { statusCode: 400 }
      },
    })
    try {
      await expect(s.sender.editReply('conv-abc', 'm1', 'x')).rejects.toBeDefined()
      expect(updates).toBe(1)
    } finally {
      rmSync(s.dir, { recursive: true, force: true })
    }
  })
})

describe('wiring', () => {
  const src = (f: string) => readFileSync(join(import.meta.dir, '..', 'src', f), 'utf8')
  const count = (s: string, needle: string) => s.split(needle).length - 1
  test('the wrapped calls', () => {
    const reply = src('teams/reply.ts')
    const sendFile = src('teams/sendFile.ts')
    const server = src('server.ts')
    expect(reply).toContain("withSendRetry('reply', () => turnContext.sendActivity(")
    expect(reply).toContain("withSendRetry('edit', () => turnContext.updateActivity(")
    expect(sendFile).toContain("withSendRetry('send_file', () => turnContext.sendActivity(activity as any))")
    const raw = "withSendRetry('send', () => turnContext.sendActivity(asActivity({ type: 'message', text })))"
    expect(count(server, raw)).toBe(1)
    expect(server).toContain("await turnContext.sendActivity(asActivity({ type: 'typing' }))")
    // The acknowledgement (see ack-placeholder) adds its own send and the reply's overwrite of it.
    expect(reply).toContain("withSendRetry('reply', () => turnContext.updateActivity(asActivity({ ...activity, id: ackId })))")
    expect(reply).toContain("withSendRetry('ack', () => turnContext.sendActivity(asActivity({ type: 'message', text })))")
    // TEAMSSDK1126: the Agents SDK retries nothing, so every user-visible outbound call is wrapped here
    // (Samu's review on #60): the status update and the reaction in reply.ts, the SharePoint card in send_file.
    expect(reply).toContain("withSendRetry('status', () => turnContext.sendActivity(")
    expect(reply).toContain("await withSendRetry('react', async () => {")
    expect(sendFile).toContain("withSendRetry('send_file', () => turnContext.sendActivity(spActivity as any))")
    expect(count(reply, 'withSendRetry(')).toBe(6)
    expect(count(sendFile, 'withSendRetry(')).toBe(2)
    expect(count(server, 'withSendRetry(')).toBe(1)
    expect(src('teams/targeted.ts')).toContain("await withSendRetry('targeted', async () => {")
    expect(src('teams/channelPost.ts')).toContain("withSendRetry('channel_post', () => send(")
    expect(src('teams/directMessage.ts')).toContain("await withSendRetry('dm_open', async () => {")
    expect(src('teams/directMessage.ts')).toContain("withSendRetry('dm', () =>")
    expect(src('teams/mentions.ts')).toContain("withSendRetry('member_lookup', () => TeamsInfo.getMember(ctx, id))")
    expect(src('teams/adapter.ts')).toContain("withSendRetry('team_details', () => TeamsInfo.getTeamDetails(c, teamId))")
  })
})
