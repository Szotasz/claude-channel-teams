import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
// Imported only to measure the SDK's own retries.
import { ConnectorClient } from 'botframework-connector'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createReplySender } from '../src/teams/reply'
import { isPluginRetryable, sdkRetriesThrottle, withSendRetry } from '../src/teams/sendRetry'

function hdrs(rec: Record<string, string>) {
  return { get: (name: string) => rec[name.toLowerCase()] }
}
const err = (statusCode: unknown, headers?: Record<string, string>) =>
  Object.assign(new Error('x'), { statusCode, ...(headers ? { response: { headers: hdrs(headers) } } : {}) })

describe('sdkRetriesThrottle', () => {
  test('cases', () => {
    expect(sdkRetriesThrottle(hdrs({}))).toBe(false)
    expect(sdkRetriesThrottle(hdrs({ 'retry-after': '0' }))).toBe(true)
    expect(sdkRetriesThrottle(hdrs({ 'retry-after': '1' }))).toBe(true)
    expect(sdkRetriesThrottle(hdrs({ 'retry-after-ms': '0' }))).toBe(true)
    expect(sdkRetriesThrottle(hdrs({ 'x-ms-retry-after-ms': '1' }))).toBe(true)
    expect(sdkRetriesThrottle(hdrs({ 'retry-after': 'soon' }))).toBe(false)
    expect(sdkRetriesThrottle(hdrs({ 'retry-after': new Date(Date.now() + 1000).toUTCString() }))).toBe(true)
    expect(sdkRetriesThrottle(hdrs({ 'retry-after': '' }))).toBe(false)
    // Measured: the SDK does not retry a non-finite wait, and the first header that parses decides.
    expect(sdkRetriesThrottle(hdrs({ 'retry-after': 'Infinity' }))).toBe(false)
    expect(sdkRetriesThrottle(hdrs({ 'retry-after-ms': 'Infinity', 'retry-after': '0' }))).toBe(false)
    expect(sdkRetriesThrottle(hdrs({ 'retry-after': '-1' }))).toBe(true)
    expect(sdkRetriesThrottle(undefined)).toBe(false)
    expect(sdkRetriesThrottle({ get: 5 })).toBe(false)
    expect(sdkRetriesThrottle({ get: () => { throw new Error('boom') } })).toBe(false)
  })
})

describe('isPluginRetryable', () => {
  test('cases', () => {
    expect(isPluginRetryable({ statusCode: 412 })).toBe(true)
    expect(isPluginRetryable(err(412, { 'retry-after': '0' }))).toBe(true)
    expect(isPluginRetryable({ statusCode: 429 })).toBe(true)
    expect(isPluginRetryable(err(429, { 'retry-after': '1' }))).toBe(false)
    expect(isPluginRetryable(err(429, { 'retry-after': 'soon' }))).toBe(true)
    for (const s of [500, 503, 400, 404]) expect(isPluginRetryable({ statusCode: s })).toBe(false)
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

  test('429 with usable Retry-After is left to the SDK', async () => {
    const e = err(429, { 'retry-after': '1' })
    const h = harness([e])
    await expect(withSendRetry('reply', h.fn, h.deps)).rejects.toBe(e)
    expect(h.calls()).toBe(1)
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

class H {
  o: Record<string, string>
  constructor(o: Record<string, string>) {
    this.o = Object.fromEntries(Object.entries(o).map(([k, v]) => [k.toLowerCase(), v]))
  }
  get(n: string) { return this.o[n.toLowerCase()] }
  contains(n: string) { return n.toLowerCase() in this.o }
  set(n: string, v: unknown) { this.o[n.toLowerCase()] = String(v) }
  remove(n: string) { const had = this.contains(n); delete this.o[n.toLowerCase()]; return had }
  toJson() { return { ...this.o } }
  rawHeaders() { return { ...this.o } }
  headersArray() { return Object.entries(this.o).map(([name, value]) => ({ name, value })) }
  headerNames() { return Object.keys(this.o) }
  headerValues() { return Object.values(this.o) }
  clone() { return new H({ ...this.o }) }
}

function client(statuses: number[], headers: Record<string, string> = {}) {
  let n = 0
  const c = new ConnectorClient({ signRequest: async (r: any) => r } as any, {
    baseUri: 'https://smba.example.test',
    httpClient: {
      async sendRequest(request: any) {
        const status = statuses[Math.min(n, statuses.length - 1)]!
        n++
        const h = new H({ 'content-type': 'application/json', ...(status === 200 ? {} : headers) })
        const bodyAsText = status === 200 ? '{"id":"m1"}' : '{"error":{"code":"x","message":"y"}}'
        return { request, status, headers: h, bodyAsText }
      },
    },
  } as any)
  return { c, calls: () => n }
}

describe('real SDK', () => {
  const deps = { sleep: async () => {}, log: () => {} }
  const msg = { type: 'message', text: 'hi' } as any

  test('429 without headers: plugin retries', async () => {
    const { c, calls } = client([429])
    await expect(withSendRetry('t', () => c.conversations.sendToConversation('a:1', msg), deps)).rejects.toMatchObject({ statusCode: 429 })
    expect(calls()).toBe(3)
  })

  test('429 with retry-after 0: only the SDK retries', async () => {
    const { c, calls } = client([429], { 'retry-after': '0' })
    await expect(withSendRetry('t', () => c.conversations.sendToConversation('a:1', msg), deps)).rejects.toMatchObject({ statusCode: 429 })
    expect(calls()).toBe(4)
  })

  test('429 with a non-finite Retry-After: the SDK does not retry, the plugin does', async () => {
    const { c, calls } = client([429], { 'retry-after': 'Infinity' })
    await expect(withSendRetry('t', () => c.conversations.sendToConversation('a:1', msg), deps)).rejects.toMatchObject({ statusCode: 429 })
    expect(calls()).toBe(3)
  })

  test('412 then 200', async () => {
    const { c, calls } = client([412, 200])
    const r = await withSendRetry('t', () => c.conversations.sendToConversation('a:1', msg), deps)
    expect((r as any).id).toBe('m1')
    expect(calls()).toBe(2)
  })

  test('400 is not retried', async () => {
    const { c, calls } = client([400])
    await expect(withSendRetry('t', () => c.conversations.sendToConversation('a:1', msg), deps)).rejects.toBeDefined()
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
      async continueConversationAsync(_a: string, _r: any, logic: any) {
        continues++
        await logic(ctx)
      },
    } as any
    const config = { stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'app', appPassword: 'pw', appType: 'SingleTenant' as const, tenantId: 't', bindHost: '127.0.0.1', port: 0 }
    return { sender: createReplySender({ config, adapter, allowlist, refs }), continues: () => continues, dir }
  }

  test('sendReply retries a 429 inside one continueConversationAsync', async () => {
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
  test('four wrapped calls', () => {
    const reply = src('teams/reply.ts')
    const sendFile = src('teams/sendFile.ts')
    const server = src('server.ts')
    expect(reply).toContain("withSendRetry('reply', () => turnContext.sendActivity(")
    expect(reply).toContain("withSendRetry('edit', () => turnContext.updateActivity(")
    expect(sendFile).toContain("withSendRetry('send_file', () => turnContext.sendActivity(activity as any))")
    const raw = "withSendRetry('send', () => turnContext.sendActivity({ type: 'message', text }))"
    expect(count(server, raw)).toBe(1)
    expect(server).toContain("await turnContext.sendActivity({ type: 'typing' })")
    expect(count(reply, 'withSendRetry(')).toBe(2)
    expect(count(sendFile, 'withSendRetry(')).toBe(1)
    expect(count(server, 'withSendRetry(')).toBe(1)
  })
})
