import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createDeliveredMessages } from '../src/teams/deliveredMessages'
import { createInboundIds, MAX_ID_CONVERSATIONS, MAX_IDS_PER_CONVERSATION } from '../src/teams/inboundIds'

const USER = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

const personal = { id: 'a:personal1', tenantId: TENANT_ID, conversationType: 'personal' }

function act(o: Record<string, unknown> = {}): Activity {
  return {
    type: 'message',
    text: 'hello',
    id: 'm1',
    timestamp: new Date('2026-05-18T19:43:00Z'),
    from: { id: 'user-1', name: 'Test User', aadObjectId: USER },
    conversation: personal,
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    ...o,
  } as unknown as Activity
}

const edit = (o: Record<string, unknown> = {}) =>
  act({ type: 'messageUpdate', channelData: { eventType: 'editMessage' }, text: 'new text', ...o })

function setup(opts: { inboundIds?: boolean; onEventThrowsOnce?: boolean; attachmentsGate?: Promise<void>; onAttachments?: () => void } = {}) {
  let pairings = 0
  const dir = mkdtempSync(join(tmpdir(), 'cct-dedupe-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(USER)
  const events: any[] = []
  let attCalls = 0
  let starts = 0
  let threw = false
  const config = {
    stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'a', appPassword: 'p',
    appType: 'SingleTenant', tenantId: TENANT_ID, bindHost: '127.0.0.1', port: 0,
    receivedFilesDir: '/tmp/unused',
  }
  const turn = makeTurnHandler({
    config: config as any,
    adapter: {} as any,
    allowlist,
    refs: createConversationRefStore(),
    onEvent: (e: any) => {
      if (opts.onEventThrowsOnce && !threw) { threw = true; throw new Error('boom') }
      events.push(e)
    },
    pending: { recordIncoming: () => { pairings++; return { action: 'suppress', reason: 'x' } } } as any,
    sendPairingDm: async () => {},
    typingPump: { start: () => { starts++ }, stop: () => {} },
    attachments: async () => { attCalls++; opts.onAttachments?.(); await opts.attachmentsGate; return { annotations: [] } },
    deliveredMessages: createDeliveredMessages(),
    ...(opts.inboundIds === false ? {} : { inboundIds: createInboundIds() }),
  } as any)
  const run = (a: Activity) => turn(new TurnContext(new NoopAdapter() as any, a))
  return { run, events, counts: () => ({ attCalls, starts }), pairings: () => pairings }
}

const att = [{ contentType: 'application/pdf', name: 'a.pdf', contentUrl: 'https://x/a.pdf' }]
const mention = { type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-1', name: 'Bot' } }
const chan = (id: string) => ({ id, tenantId: TENANT_ID, conversationType: 'channel' })

describe('createInboundIds', () => {
  test('the bounds', () => {
    expect(MAX_IDS_PER_CONVERSATION).toBe(500)
    expect(MAX_ID_CONVERSATIONS).toBe(1000)
  })
  test('first seen / repeat / other conversation', () => {
    const d = createInboundIds()
    expect(d.firstSeen('c', '1')).toBe(true)
    expect(d.firstSeen('c', '1')).toBe(false)
    expect(d.firstSeen('d', '1')).toBe(true)
  })

  test('forget', () => {
    const d = createInboundIds()
    d.firstSeen('c', '1')
    d.forget('c', '1')
    expect(d.firstSeen('c', '1')).toBe(true)
    expect(() => d.forget('nope', '1')).not.toThrow()
    expect(() => d.forget('c', 'nope')).not.toThrow()
  })

  test('per-conversation bound', () => {
    const full = createInboundIds()
    for (let i = 0; i < MAX_IDS_PER_CONVERSATION; i++) full.firstSeen('c', String(i))
    // Exactly at the bound nothing is dropped yet.
    expect(full.firstSeen('c', '0')).toBe(false)
    const d = createInboundIds()
    for (let i = 0; i <= MAX_IDS_PER_CONVERSATION; i++) d.firstSeen('c', String(i))
    expect(d.firstSeen('c', '500')).toBe(false)
    expect(d.firstSeen('c', '0')).toBe(true)
  })

  test('conversation bound', () => {
    const full = createInboundIds()
    for (let i = 0; i < MAX_ID_CONVERSATIONS; i++) full.firstSeen(`k${i}`, 'x')
    expect(full.firstSeen('k0', 'x')).toBe(false)
    const d = createInboundIds()
    for (let i = 0; i <= MAX_ID_CONVERSATIONS; i++) d.firstSeen(`k${i}`, 'x')
    expect(d.firstSeen('k1000', 'x')).toBe(false)
    expect(d.firstSeen('k0', 'x')).toBe(true)
  })

  test('a conversation emptied by forget no longer takes a slot', () => {
    const d = createInboundIds()
    for (let i = 0; i < MAX_ID_CONVERSATIONS - 1; i++) d.firstSeen(`k${i}`, 'x')
    d.firstSeen('c', '1')
    d.forget('c', '1')
    // 999 conversations are left, so one more fits without dropping the oldest.
    d.firstSeen(`k${MAX_ID_CONVERSATIONS - 1}`, 'x')
    expect(d.firstSeen('k0', 'x')).toBe(false)
  })
  test('touching moves to newest', () => {
    const d = createInboundIds()
    for (let i = 0; i < MAX_ID_CONVERSATIONS; i++) d.firstSeen(`k${i}`, 'x')
    d.firstSeen('k0', 'y')
    d.firstSeen('k1000', 'x')
    expect(d.firstSeen('k1', 'x')).toBe(true)
    expect(d.firstSeen('k0', 'x')).toBe(false)
  })
})

describe('repeated delivery in the turn handler', () => {
  test('same activity twice', async () => {
    const s = setup()
    await s.run(act({ attachments: att }))
    await s.run(act({ attachments: att }))
    expect(s.events.length).toBe(1)
    expect(s.counts()).toEqual({ attCalls: 1, starts: 1 })
  })

  test('a repeat that arrives while the first is still downloading is dropped', async () => {
    let release!: () => void
    let downloading!: () => void
    const inDownload = new Promise<void>(r => { downloading = r })
    const s = setup({ attachmentsGate: new Promise<void>(r => { release = r }), onAttachments: () => downloading() })
    const first = s.run(act({ attachments: att }))
    // Wait until the first turn is inside the attachment download.
    await inDownload
    await s.run(act({ attachments: att }))
    release()
    await first
    expect(s.events.length).toBe(1)
    expect(s.counts()).toEqual({ attCalls: 1, starts: 1 })
  })
  test('same id in different conversations', async () => {
    const s = setup()
    await s.run(act())
    await s.run(act({ conversation: { ...personal, id: 'a:personal2' } }))
    expect(s.events.length).toBe(2)
  })

  test('same id in two threads of one channel', async () => {
    const s = setup()
    const c = (id: string) => act({ conversation: chan(id), text: '<at>Bot</at> hi', entities: [mention] })
    await s.run(c('19:c@thread.tacv2;messageid=1'))
    await s.run(c('19:c@thread.tacv2;messageid=2'))
    expect(s.events.length).toBe(1)
  })

  test('message without id', async () => {
    const s = setup()
    await s.run(act({ id: undefined }))
    await s.run(act({ id: undefined }))
    await s.run(act({ id: '' }))
    await s.run(act({ id: '' }))
    expect(s.events.length).toBe(4)
  })
  test('a repeat from an unknown sender does not reach pairing again, and is logged once', async () => {
    const s = setup()
    const stranger = act({ id: 's1', from: { id: 'user-9', name: 'Test Stranger', aadObjectId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' } })
    const lines: string[] = []
    const write = process.stderr.write
    process.stderr.write = ((chunk: unknown) => { lines.push(String(chunk)); return true }) as typeof process.stderr.write
    let firstRunLines = 0
    try {
      await s.run(stranger)
      firstRunLines = lines.length
      await s.run(stranger)
    } finally {
      process.stderr.write = write
    }
    expect(s.pairings()).toBe(1)
    // The repeat writes exactly one line, and it is the drop line.
    const repeatLines = lines.slice(firstRunLines)
    expect(repeatLines.length).toBe(1)
    expect(repeatLines[0]).toContain('repeated delivery of message id=s1')
  })

  test('failed turn is forgotten', async () => {
    const s = setup({ onEventThrowsOnce: true })
    await expect(s.run(act())).rejects.toThrow()
    await s.run(act())
    expect(s.events.length).toBe(1)
  })

  test('repeated edit, delete and restore events are forwarded', async () => {
    const s = setup()
    await s.run(act())
    await s.run(edit())
    await s.run(edit())
    const del = act({ type: 'messageDelete', channelData: { eventType: 'softDeleteMessage' }, text: undefined })
    await s.run(del)
    await s.run(del)
    const restore = edit({ channelData: { eventType: 'undeleteMessage' } })
    await s.run(restore)
    await s.run(restore)
    expect(s.events.length).toBe(7)
  })

  test('not wired', async () => {
    const s = setup({ inboundIds: false })
    await s.run(act())
    await s.run(act())
    expect(s.events.length).toBe(2)
  })
})

describe('server wiring', () => {
  const src = readFileSync(join(import.meta.dir, '../src/server.ts'), 'utf8')
  test('creates the store', () => {
    expect(src).toContain('const inboundIds = createInboundIds()')
  })
  test('passes it to startListener', () => {
    const start = src.indexOf('startListener({')
    expect(start).toBeGreaterThan(-1)
    const block = src.slice(start, src.indexOf('\n})', start))
    expect(block).toContain('inboundIds,')
  })
})
