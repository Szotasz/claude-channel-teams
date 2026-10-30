import { afterAll, describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync, readFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { loadConfig, parseAckDelayMs, parseAckText } from '../src/config'
import { ACK_MAX_AGE_MS, createAckPlaceholders } from '../src/teams/ackPlaceholder'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createReplySender } from '../src/teams/reply'
import { createStatusStreams } from '../src/teams/statusStream'
// Every temp dir of this file, removed after the last test (they used to stay in the system temp dir).
const tmpDirs: string[] = []
function tempDir(prefix: string): string {
  const dir = mkdtempSync(join(tmpdir(), prefix))
  tmpDirs.push(dir)
  return dir
}
afterAll(() => {
  for (const d of tmpDirs) rmSync(d, { recursive: true, force: true })
})

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'

describe('ack config', () => {
  test('parseAckText', () => {
    expect(parseAckText(undefined)).toBeUndefined()
    expect(parseAckText('')).toBeUndefined()
    expect(parseAckText('  ')).toBeUndefined()
    expect(parseAckText(' Got it ')).toBe('Got it')
    expect(parseAckText('a'.repeat(200))).toBe('a'.repeat(200))
    expect(parseAckText('a'.repeat(201))).toBeNull()
    expect(parseAckText('a\nb')).toBeNull()
    expect(parseAckText('a\u0085b')).toBeNull()
    expect(parseAckText('a\u007fb')).toBeNull()
  })

  test('parseAckDelayMs', () => {
    expect(parseAckDelayMs(undefined)).toBe(3000)
    expect(parseAckDelayMs('')).toBe(3000)
    expect(parseAckDelayMs('0')).toBe(0)
    expect(parseAckDelayMs('10000')).toBe(10000)
    for (const bad of ['10001', '-1', '1.5', 'x']) expect(parseAckDelayMs(bad)).toBeNull()
  })

  function load(extra: Record<string, string>) {
    const lines: string[] = []
    const orig = process.stderr.write
    const saved = { ...process.env }
    for (const k of ['TEAMS_ACK_TEXT', 'TEAMS_ACK_DELAY_MS']) delete process.env[k]
    Object.assign(process.env, {
      TEAMS_STATE_DIR: tempDir('cct-ack-'),
      TEAMS_BOT_APP_ID: 'app',
      TEAMS_BOT_APP_PASSWORD: 'pw',
      TEAMS_BOT_TENANT_ID: TENANT_ID,
      ...extra,
    })
    process.stderr.write = ((s: string) => (lines.push(String(s)), true)) as any
    try {
      return { config: loadConfig(), lines }
    } finally {
      process.stderr.write = orig
      for (const k of Object.keys(process.env)) if (!(k in saved)) delete process.env[k]
      Object.assign(process.env, saved)
    }
  }

  test('loadConfig', () => {
    const a = load({ TEAMS_ACK_TEXT: 'Got it' })
    expect(a.config.ackText).toBe('Got it')
    expect(a.config.ackDelayMs).toBe(3000)
    const b = load({ TEAMS_ACK_DELAY_MS: 'x' })
    expect(b.config.ackDelayMs).toBe(3000)
    expect(b.lines).toContain('teams channel: TEAMS_ACK_DELAY_MS must be an integer from 0 to 10000; using 3000\n')
    const c = load({ TEAMS_ACK_TEXT: 'a\u0001b' })
    expect(c.config.ackText).toBeUndefined()
    expect(c.lines).toContain(
      'teams channel: TEAMS_ACK_TEXT must be 1-200 characters without control characters; the acknowledgement stays off\n',
    )
  })
})

function setup(send: (c: string) => Promise<string | undefined>, delayMs = 3000) {
  let t = 1000
  const timers: { fn: () => void; ms: number }[] = []
  const cleared: unknown[] = []
  const logs: string[] = []
  const acks = createAckPlaceholders({
    delayMs,
    send,
    now: () => t,
    setTimer: (fn, ms) => (timers.push({ fn, ms }), timers.length - 1),
    clearTimer: h => cleared.push(h),
    log: l => logs.push(l),
  })
  return { acks, timers, cleared, logs, advance: (ms: number) => (t += ms) }
}

describe('createAckPlaceholders', () => {
  test('take before the timer fires', async () => {
    let calls = 0
    const s = setup(async () => (calls++, 'x'))
    s.acks.arm('c')
    expect(await s.acks.take('c')).toBeUndefined()
    expect(s.cleared).toEqual([0])
    expect(calls).toBe(0)
  })

  test('fired ack is taken once; timer uses delayMs', async () => {
    const s = setup(async () => 'ack1', 1234)
    s.acks.arm('c')
    expect(s.timers[0]!.ms).toBe(1234)
    s.timers[0]!.fn()
    expect(await s.acks.take('c')).toBe('ack1')
    expect(await s.acks.take('c')).toBeUndefined()
  })

  test('arm twice gives one timer', () => {
    const s = setup(async () => 'a')
    s.acks.arm('c')
    s.acks.arm('c')
    expect(s.timers.length).toBe(1)
  })

  test('take waits for a send in flight', async () => {
    let resolve!: (v: string) => void
    const s = setup(() => new Promise<string>(r => (resolve = r)))
    s.acks.arm('c')
    s.timers[0]!.fn()
    let got: string | undefined
    const p = s.acks.take('c').then(v => (got = v))
    await Promise.resolve()
    expect(got).toBeUndefined()
    resolve('late')
    await p
    expect(got).toBe('late')
  })

  test('send rejects', async () => {
    const s = setup(async () => {
      throw new Error('boom')
    })
    s.acks.arm('c')
    s.timers[0]!.fn()
    expect(await s.acks.take('c')).toBeUndefined()
    expect(s.logs).toEqual(['teams channel: acknowledgement failed conv=c: boom\n'])
  })

  test('send resolves undefined', async () => {
    const s = setup(async () => undefined)
    s.acks.arm('c')
    s.timers[0]!.fn()
    expect(await s.acks.take('c')).toBeUndefined()
    // Without an id nothing is pending, so the next delivered message can arm a new acknowledgement.
    s.acks.arm('d')
    s.timers[1]!.fn()
    await Promise.resolve()
    await Promise.resolve()
    s.acks.arm('d')
    expect(s.timers.length).toBe(3)
  })

  test('arm while the acknowledgement is being sent adds no timer', async () => {
    let resolve!: (v: string) => void
    const s = setup(() => new Promise<string>(r => (resolve = r)))
    s.acks.arm('c')
    s.timers[0]!.fn()
    s.acks.arm('c')
    expect(s.timers.length).toBe(1)
    resolve('a')
    expect(await s.acks.take('c')).toBe('a')
  })

  test('old ack is not returned and arm starts a new timer', async () => {
    const s = setup(async () => 'a')
    s.acks.arm('c')
    s.timers[0]!.fn()
    await Promise.resolve()
    await Promise.resolve()
    s.advance(ACK_MAX_AGE_MS)
    s.acks.arm('c')
    expect(s.timers.length).toBe(2)
    s.acks.cancel('c')
    // old sent ack replaced by the timer, then cancelled
    expect(await s.acks.take('c')).toBeUndefined()

    const s2 = setup(async () => 'a')
    s2.acks.arm('c')
    s2.timers[0]!.fn()
    await Promise.resolve()
    await Promise.resolve()
    s2.advance(ACK_MAX_AGE_MS)
    expect(await s2.acks.take('c')).toBeUndefined()
  })

  test('arm while a young ack is pending adds no timer', async () => {
    const s = setup(async () => 'a')
    s.acks.arm('c')
    s.timers[0]!.fn()
    await Promise.resolve()
    await Promise.resolve()
    s.advance(1000)
    s.acks.arm('c')
    expect(s.timers.length).toBe(1)
  })

  test('cancel clears a timer but leaves a sent ack', async () => {
    const s = setup(async () => 'a')
    s.acks.arm('c')
    s.acks.cancel('c')
    expect(s.cleared).toEqual([0])
    s.acks.arm('d')
    s.timers[1]!.fn()
    await Promise.resolve()
    await Promise.resolve()
    s.acks.cancel('d')
    expect(await s.acks.take('d')).toBe('a')
  })

  test('stopAll clears pending timers', () => {
    const s = setup(async () => 'a')
    s.acks.arm('c')
    s.acks.arm('d')
    s.acks.stopAll()
    expect(s.cleared).toEqual([0, 1])
    // A message delivered while the plugin shuts down arms nothing.
    s.acks.arm('e')
    expect(s.timers.length).toBe(2)
    return s.acks.take('c').then(id => expect(id).toBeUndefined())
  })
})

describe('reply with acknowledgement', () => {
  function make(takeResult: string | undefined, updateThrows = false, allowed = true, realAcks?: ReturnType<typeof createAckPlaceholders>) {
    const dir = tempDir('cct-ackreply-')
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    if (allowed) allowlist.addEntry(FIXTURE_ID)
    const refs = createConversationRefStore()
    refs.put('conv-abc', { conversation: { id: 'conv-abc' } } as any, FIXTURE_ID)
    const sent: any[] = []
    const updated: any[] = []
    const adapter = {
      async continueConversationAsync(_a: string, _r: any, logic: any) {
        await logic({
          async sendActivity(act: any) {
            sent.push(act)
            return { id: 'new1' }
          },
          async updateActivity(act: any) {
            updated.push(act)
            if (updateThrows) throw { statusCode: 400 }
          },
        })
      },
    } as any
    const acks = realAcks ?? { arm() {}, take: async () => takeResult, cancel() {}, stopAll() {} }
    const replier = createReplySender({
      config: {
        stateDir: '/tmp/unused',
        allowlistFile: '/tmp/unused',
        appId: 'app',
        appPassword: 'pw',
        appType: 'SingleTenant' as const,
        tenantId: TENANT_ID,
        bindHost: '127.0.0.1',
        port: 0,
      } as any,
      adapter,
      allowlist,
      refs,
      acks,
    })
    return { replier, sent, updated }
  }

  async function captureStderr<T>(fn: () => Promise<T>) {
    const lines: string[] = []
    const orig = process.stderr.write
    process.stderr.write = ((s: string) => (lines.push(String(s)), true)) as any
    try {
      return { result: await fn(), lines }
    } finally {
      process.stderr.write = orig
    }
  }

  test('overwrites the acknowledgement', async () => {
    const { replier, sent, updated } = make('ack1')
    const id = await replier.sendReply('conv-abc', 'final')
    expect(id).toBe('ack1')
    expect(sent).toEqual([])
    expect(updated.length).toBe(1)
    expect(updated[0].id).toBe('ack1')
    expect(updated[0].text).toBe('final')
    expect(updated[0].textFormat).toBe('markdown')
  })

  test('the AI label entity goes into the update', async () => {
    const { replier, updated } = make('ack1')
    await replier.sendReply('conv-abc', 'final', undefined, { aiGenerated: true })
    expect(updated[0].entities.length).toBe(1)
    expect(updated[0].entities[0].type).toBe('https://schema.org/Message')
  })

  test('with the real store: a reply during the acknowledgement send waits for it and overwrites it', async () => {
    let resolve!: (id: string) => void
    let fire!: () => void
    const acks = createAckPlaceholders({
      delayMs: 3000,
      send: () => new Promise<string>(r => { resolve = r }),
      setTimer: fn => { fire = fn; return 1 },
      clearTimer: () => {},
      log: () => {},
    })
    const { replier, sent, updated } = make(undefined, false, true, acks)
    acks.arm('conv-abc')
    fire()
    const p = replier.sendReply('conv-abc', 'final')
    await Promise.resolve()
    expect(updated.length).toBe(0)
    resolve('ack9')
    expect(await p).toBe('ack9')
    expect(sent).toEqual([])
    expect(updated[0].id).toBe('ack9')
  })

  test('card goes into the update', async () => {
    const { replier, updated } = make('ack1')
    const card = { type: 'AdaptiveCard', version: '1.5', body: [{ type: 'TextBlock', text: 'hi' }] }
    await replier.sendReply('conv-abc', 'final', card)
    expect(updated[0].attachments.length).toBe(1)
    expect(updated[0].attachments[0].content).toEqual(card)
  })

  test('falls back to a new message when the update fails', async () => {
    const { replier, sent, updated } = make('ack1', true)
    const { result, lines } = await captureStderr(() => replier.sendReply('conv-abc', 'final'))
    expect(updated.length).toBe(1)
    expect(sent.length).toBe(1)
    expect(result).toBe('new1')
    expect(lines.some(l => l.startsWith('teams channel: overwriting the acknowledgement failed'))).toBe(true)
  })

  test('no ack sends normally', async () => {
    const { replier, sent, updated } = make(undefined)
    expect(await replier.sendReply('conv-abc', 'final')).toBe('new1')
    expect(sent.length).toBe(1)
    expect(updated).toEqual([])
  })

  test('sendAck', async () => {
    const { replier, sent } = make(undefined)
    expect(await replier.sendAck('conv-abc', 'Got it')).toBe('new1')
    expect(sent).toEqual([{ type: 'message', text: 'Got it' }])
    const other = make(undefined, false, false)
    await expect(other.replier.sendAck('conv-abc', 'Got it')).rejects.toThrow()
  })
})

describe('server wiring', () => {
  const src = readFileSync(join(import.meta.dir, '../src/server.ts'), 'utf8')
  test('arm, cancel, stopAll, deps', () => {
    expect(src).toMatch(/if \(acks && !event\.messageEvent && event\.mentioned !== false\) acks\.arm\(event\.conversationId\)/)
    const sf = src.slice(src.indexOf("case 'send_file'"))
    expect(sf.slice(0, sf.indexOf("case '", 10))).toContain('acks?.cancel(conversationId)')
    const sd = src.slice(src.indexOf('typingPump.stopAll()'))
    expect(sd.slice(0, 80)).toContain('acks?.stopAll()')
    expect(src).toMatch(/createReplySender\(\{[^}]*\backs\b[^}]*\}\)/)
  })
  test('the acknowledgement is armed only after the notification to Claude succeeded', () => {
    const sink = src.slice(src.indexOf('const onEvent: ChannelEventSink'))
    const ok = sink.indexOf("process.stderr.write('teams channel: notify ok\\n')")
    const arm = sink.indexOf('acks.arm(event.conversationId)')
    const failed = sink.indexOf('.catch(err => {')
    expect(ok).toBeGreaterThan(-1)
    expect(arm).toBeGreaterThan(ok)
    expect(failed).toBeGreaterThan(arm)
  })
})

describe('reply with acknowledgement and a status stream', () => {
  function make() {
    const dir = tempDir('cct-ackstream-')
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(FIXTURE_ID)
    const refs = createConversationRefStore()
    refs.put('dm', { conversation: { id: 'a:1', conversationType: 'personal' } } as any, FIXTURE_ID)
    const sent: any[] = []
    const updated: any[] = []
    const adapter = {
      async continueConversationAsync(_a: string, _r: any, logic: any) {
        await logic({
          async sendActivity(act: any) {
            sent.push(act)
            return { id: 's-1' }
          },
          async updateActivity(act: any) {
            updated.push(act)
          },
        })
      },
    } as any
    const replier = createReplySender({
      config: {
        stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'app', appPassword: 'pw',
        appType: 'SingleTenant' as const, tenantId: TENANT_ID, bindHost: '127.0.0.1', port: 0,
      } as any,
      adapter,
      allowlist,
      refs,
      acks: { arm() {}, take: async () => 'ack1', cancel() {}, stopAll() {} },
      statusStreams: createStatusStreams({ sleep: async () => {} }),
    })
    return { replier, sent, updated }
  }

  test('a reply that closes the status stream leaves the acknowledgement as sent', async () => {
    const s = make()
    await s.replier.sendStatus('dm', 'Working')
    const id = await s.replier.sendReply('dm', 'Done')
    expect(s.updated).toEqual([])
    expect(s.sent[1].text).toBe('Done')
    expect(s.sent[1].entities).toEqual([{ type: 'streaminfo', streamId: 's-1', streamType: 'final' }])
    expect(id).toBe('s-1')
  })

  test('with no open stream the reply overwrites the acknowledgement', async () => {
    const s = make()
    const id = await s.replier.sendReply('dm', 'Done')
    expect(id).toBe('ack1')
    expect(s.sent).toEqual([])
    expect(s.updated.length).toBe(1)
    expect(s.updated[0].text).toBe('Done')
  })
})
