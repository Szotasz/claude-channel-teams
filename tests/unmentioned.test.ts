import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler, baseConversationId } from '../src/teams/adapter'
import { parseIdList } from '../src/config'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const UNKNOWN_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const CHAN1 = '19:chan1@thread.tacv2;messageid=111'

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function makeActivity(overrides: Partial<Activity> = {}): Activity {
  return {
    type: 'message',
    text: 'hello there',
    id: 'msg-1',
    timestamp: new Date('2026-05-18T19:43:00Z'),
    from: { id: 'user-1', name: 'Adnan Khan', aadObjectId: FIXTURE_ID },
    conversation: { id: CHAN1, tenantId: TENANT_ID, conversationType: 'channel' },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    ...overrides,
  } as unknown as Activity
}

function setup(unmentioned?: Set<string>, extra: Record<string, unknown> = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-unmentioned-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(FIXTURE_ID)
  const refs = createConversationRefStore()
  const events: any[] = []
  const config = {
    stateDir: '/tmp/unused',
    allowlistFile: '/tmp/unused',
    appId: 'app-id',
    appPassword: 'app-pw',
    appType: 'SingleTenant' as const,
    tenantId: TENANT_ID,
    bindHost: '127.0.0.1',
    port: 0,
    receivedFilesDir: '/tmp/unused',
    ...(unmentioned ? { unmentionedConversations: unmentioned } : {}),
  }
  const turn = makeTurnHandler({
    config: config as any,
    adapter: {} as any,
    allowlist,
    refs,
    onEvent: e => events.push(e),
    ...extra,
  } as any)
  const run = (a: Activity) => turn(new TurnContext(new NoopAdapter() as any, a))
  return { run, events, refs }
}

const chan2 = (o: Partial<Activity> = {}) =>
  makeActivity({
    conversation: { id: '19:chan2@thread.tacv2;messageid=5', tenantId: TENANT_ID, conversationType: 'channel' } as any,
    ...o,
  })

describe('unmentioned delivery', () => {
  test('default: unmentioned channel message is ignored', async () => {
    const s = setup()
    await s.run(makeActivity())
    expect(s.events.length).toBe(0)
  })

  test('opted-in channel delivers unmentioned message', async () => {
    const s = setup(new Set(['19:chan1@thread.tacv2']))
    await s.run(makeActivity())
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('hello there')
    expect(s.events[0].mentioned).toBe(false)
    expect(s.refs.get(CHAN1)).toBeDefined()
  })

  test('other channel is still ignored', async () => {
    const s = setup(new Set(['19:chan1@thread.tacv2']))
    await s.run(chan2())
    expect(s.events.length).toBe(0)
  })

  test("'*' opts in every shared conversation", async () => {
    const s = setup(new Set(['*']))
    await s.run(chan2())
    expect(s.events.length).toBe(1)
  })

  test('non-allowlisted unmentioned sender is not paired', async () => {
    const dmCalls: any[] = []
    const incoming: any[] = []
    const pending = {
      recordIncoming: (x: any) => {
        incoming.push(x)
        return { action: 'suppress', reason: 'test' }
      },
    }
    const s = setup(new Set(['19:chan1@thread.tacv2']), {
      pending,
      sendPairingDm: async (...a: any[]) => { dmCalls.push(a) },
    })
    await s.run(makeActivity({ from: { id: 'u2', name: 'X', aadObjectId: UNKNOWN_ID } as any }))
    expect(s.events.length).toBe(0)
    expect(dmCalls.length).toBe(0)
    expect(incoming.length).toBe(0)
  })

  test('the allowlist is read once per turn (no pairing if it flips between reads)', async () => {
    let calls = 0
    const incoming: any[] = []
    const flipping = { isAllowed: () => calls++ === 0 }
    const s = setup(new Set(['19:chan1@thread.tacv2']), {
      allowlist: flipping,
      pending: { recordIncoming: (x: any) => { incoming.push(x); return { action: 'suppress', reason: 'test' } } },
      sendPairingDm: async () => {},
    })
    await s.run(makeActivity())
    expect(calls).toBe(1)
    expect(s.events.length).toBe(1)
    expect(incoming.length).toBe(0)
  })

  test('a MENTIONED unknown sender in an opted-in channel still gets the pairing path', async () => {
    const incoming: any[] = []
    const s = setup(new Set(['19:chan1@thread.tacv2']), {
      pending: { recordIncoming: (x: any) => { incoming.push(x); return { action: 'suppress', reason: 'test' } } },
      sendPairingDm: async () => {},
    })
    await s.run(makeActivity({
      from: { id: 'u2', name: 'X', aadObjectId: UNKNOWN_ID } as any,
      text: '<at>Bot</at> hello',
      entities: [{ type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-1', name: 'Bot' } }] as any,
    }))
    expect(s.events.length).toBe(0)
    expect(incoming.length).toBe(1)
  })

  test('an unmentioned "yes <id>" is not a tool-approval verdict, it is a normal message', async () => {
    const resolved: any[] = []
    const s = setup(new Set(['19:chan1@thread.tacv2']), {
      permission: { resolve: (id: string, v: string) => { resolved.push([id, v]); return true } },
    })
    await s.run(makeActivity({ text: 'yes abcde' }))
    expect(resolved.length).toBe(0)
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('yes abcde')
  })

  test('a mentioned "yes <id>" in a channel is still a verdict', async () => {
    const resolved: any[] = []
    const s = setup(new Set(['19:chan1@thread.tacv2']), {
      permission: { resolve: (id: string, v: string) => { resolved.push([id, v]); return true } },
    })
    await s.run(makeActivity({
      text: '<at>Bot</at> yes abcde',
      entities: [{ type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-1', name: 'Bot' } }] as any,
    }))
    expect(resolved).toEqual([['abcde', 'allow']])
  })

  test('mentioned message reports mentioned=true', async () => {
    const s = setup(new Set(['19:chan1@thread.tacv2']))
    await s.run(makeActivity({
      text: '<at>Bot</at> hello',
      entities: [{ type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-1', name: 'Bot' } }] as any,
    }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].mentioned).toBe(true)
  })

  test('1:1 event has no mentioned key', async () => {
    const s = setup(new Set())
    await s.run(makeActivity({
      conversation: { id: 'a:personal-1', tenantId: TENANT_ID, conversationType: 'personal' } as any,
    }))
    expect(s.events.length).toBe(1)
    expect('mentioned' in s.events[0]).toBe(false)
  })

  test('baseConversationId', () => {
    expect(baseConversationId('19:x@thread.tacv2;messageid=9')).toBe('19:x@thread.tacv2')
    expect(baseConversationId('a:1')).toBe('a:1')
  })

  test('parseIdList', () => {
    expect(parseIdList(' a , ,b,* ')).toEqual(new Set(['a', 'b', '*']))
    expect(parseIdList(undefined).size).toBe(0)
  })
})
