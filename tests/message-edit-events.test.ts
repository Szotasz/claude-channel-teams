import { describe, expect, test, afterAll } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import {
  createDeliveredMessages,
  MAX_CONVERSATIONS,
  MAX_MESSAGES_PER_CONVERSATION,
} from '../src/teams/deliveredMessages'

const USER = '00000000-0000-4000-8000-000000000001'
const OTHER = '00000000-0000-4000-8000-000000000003'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const STRANGER = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'

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
    text: 'original',
    id: 'msg-1',
    timestamp: new Date('2026-05-18T19:43:00Z'),
    from: { id: 'user-1', name: 'Adnan', aadObjectId: USER },
    conversation: personal,
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    ...o,
  } as unknown as Activity
}

const edit = (o: Record<string, unknown> = {}) =>
  act({ type: 'messageUpdate', channelData: { eventType: 'editMessage' }, text: 'new text', ...o })

const tmpDirs: string[] = []
afterAll(() => {
  for (const d of tmpDirs) rmSync(d, { recursive: true, force: true })
})

function setup(wire = true) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-edit-'))
  tmpDirs.push(dir)
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(USER)
  allowlist.addEntry(OTHER)
  const events: any[] = []
  const pairing: string[] = []
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
    onEvent: e => events.push(e),
    pending: { recordIncoming: () => { pairing.push('pending'); return { action: 'suppress', reason: 'x' } } } as any,
    sendPairingDm: async () => { pairing.push('dm') },
    ...(wire ? { deliveredMessages: createDeliveredMessages() } : {}),
  } as any)
  const run = (a: Activity) => turn(new TurnContext(new NoopAdapter() as any, a))
  return { run, events, allowlist, pairing }
}

describe('message edit / delete / restore events', () => {
  test('G1 edit of a delivered message', async () => {
    const s = setup()
    await s.run(act())
    await s.run(edit())
    expect(s.events.length).toBe(2)
    expect(s.events[1].text).toBe('[Message edited] new text')
    expect(s.events[1].messageEvent).toBe('edited')
    expect(s.events[1].messageId).toBe('msg-1')
  })

  test('G2 never delivered', async () => {
    const s = setup()
    await s.run(edit({ id: 'other' }))
    expect(s.events.length).toBe(0)
  })

  test('G3 different sender', async () => {
    const s = setup()
    await s.run(act())
    await s.run(edit({ from: { id: 'u2', name: 'B', aadObjectId: OTHER } }))
    expect(s.events.length).toBe(1)
  })

  test('G4 sender removed from allowlist', async () => {
    const s = setup()
    await s.run(act())
    s.allowlist.removeEntry(USER)
    await s.run(edit())
    expect(s.events.length).toBe(1)
  })

  test('G5 soft delete', async () => {
    const s = setup()
    await s.run(act())
    await s.run(act({ type: 'messageDelete', channelData: { eventType: 'softDeleteMessage' }, text: undefined }))
    expect(s.events[1].text).toBe('[Message deleted]')
    expect(s.events[1].messageEvent).toBe('deleted')
  })

  test('G6 undelete', async () => {
    const s = setup()
    await s.run(act())
    await s.run(edit({ channelData: { eventType: 'undeleteMessage' }, text: 'restored text' }))
    expect(s.events[1].text).toBe('[Message restored] restored text')
    expect(s.events[1].messageEvent).toBe('restored')
  })

  test('G7 channel with thread suffix and mention', async () => {
    const s = setup()
    const mention = {
      type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-1', name: 'Bot' },
    }
    const chan = (id: string) => ({ id, tenantId: TENANT_ID, conversationType: 'channel' })
    await s.run(act({
      conversation: chan('19:c@thread.tacv2;messageid=111'),
      text: '<at>Bot</at> hello',
      entities: [mention],
    }))
    await s.run(edit({
      conversation: chan('19:c@thread.tacv2'),
      text: '<at>Bot</at> fixed link',
      entities: [mention],
    }))
    expect(s.events.length).toBe(2)
    expect(s.events[1].text).toBe('[Message edited] fixed link')
  })

  test('G12 channel, the other way round: delivered without the thread suffix, edited with it', async () => {
    const s = setup()
    const mention = {
      type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-1', name: 'Bot' },
    }
    const chan = (id: string) => ({ id, tenantId: TENANT_ID, conversationType: 'channel' })
    await s.run(act({ conversation: chan('19:c@thread.tacv2'), text: '<at>Bot</at> hello', entities: [mention] }))
    await s.run(edit({ conversation: chan('19:c@thread.tacv2;messageid=111'), text: 'fixed', entities: [] }))
    expect(s.events.length).toBe(2)
    expect(s.events[1].text).toBe('[Message edited] fixed')
  })

  test('G8 edit without text', async () => {
    const s = setup()
    await s.run(act())
    await s.run(edit({ text: undefined }))
    expect(s.events[1].text).toBe('[Message edited] (the new text was not included)')
  })

  test('G9 other event types are ignored', async () => {
    const s = setup()
    await s.run(act())
    await s.run(edit({ channelData: { eventType: 'somethingElse' } }))
    await s.run(act({ type: 'messageDelete', channelData: { eventType: 'hardDeleteMessage' } }))
    expect(s.events.length).toBe(1)
  })

  test('G10 not wired', async () => {
    const s = setup(false)
    await s.run(act())
    await s.run(edit())
    expect(s.events.length).toBe(1)
  })

  test('G11 non-allowlisted stranger in a channel', async () => {
    const s = setup()
    await s.run(edit({
      from: { id: 'u3', name: 'S', aadObjectId: STRANGER },
      conversation: { id: '19:c@thread.tacv2', tenantId: TENANT_ID, conversationType: 'channel' },
    }))
    expect(s.events.length).toBe(0)
    expect(s.pairing.length).toBe(0)
  })
})

describe('deliveredMessages bounds', () => {
  test('501st id drops the oldest', () => {
    const d = createDeliveredMessages()
    for (let i = 0; i <= MAX_MESSAGES_PER_CONVERSATION; i++) d.remember('c', `m${i}`, 'u')
    expect(d.get('c', 'm0')).toBeUndefined()
    expect(d.get('c', 'm1')).toEqual({ aadObjectId: 'u' })
    expect(d.get('c', `m${MAX_MESSAGES_PER_CONVERSATION}`)).toBeDefined()
  })

  test('1001st conversation drops the oldest', () => {
    const d = createDeliveredMessages()
    for (let i = 0; i <= MAX_CONVERSATIONS; i++) d.remember(`c${i}`, 'm', 'u')
    expect(d.get('c0', 'm')).toBeUndefined()
    expect(d.get('c1', 'm')).toBeDefined()
  })

  test('re-remembering moves to newest', () => {
    const d = createDeliveredMessages()
    for (let i = 0; i < MAX_MESSAGES_PER_CONVERSATION; i++) d.remember('c', `m${i}`, 'u')
    d.remember('c', 'm0', 'u')
    d.remember('c', 'new', 'u')
    expect(d.get('c', 'm0')).toBeDefined()
    expect(d.get('c', 'm1')).toBeUndefined()

    const e = createDeliveredMessages()
    for (let i = 0; i < MAX_CONVERSATIONS; i++) e.remember(`c${i}`, 'm', 'u')
    e.remember('c0', 'm2', 'u')
    e.remember('cNew', 'm', 'u')
    expect(e.get('c0', 'm')).toBeDefined()
    expect(e.get('c1', 'm')).toBeUndefined()
  })
})
