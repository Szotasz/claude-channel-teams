import { describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { createPendingStore } from '../src/pairing/pair'
import { createAllowlist } from '../src/pairing/allowlist'
import {
  createConversationRefStore,
  isPersonalReference,
} from '../src/teams/conversationRefs'
import { makeTurnHandler } from '../src/teams/adapter'

const SENDER_ID = '00000000-0000-4000-8000-0000000000a1'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function makeConfig() {
  return {
    stateDir: '/tmp/unused',
    allowlistFile: '/tmp/unused',
    pendingFile: '/tmp/unused',
    appId: 'app-id',
    appPassword: 'app-pw',
    appType: 'SingleTenant' as const,
    tenantId: TENANT_ID,
    bindHost: '127.0.0.1',
    port: 0,
  }
}

function activity(overrides: Partial<Activity> = {}): Activity {
  return {
    type: 'message',
    text: 'hi',
    id: 'msg-1',
    timestamp: new Date('2026-05-17T19:30:00Z'),
    from: { id: 'user-1', name: 'Stranger', aadObjectId: SENDER_ID },
    conversation: { id: 'a:1xyz', tenantId: TENANT_ID },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://example.invalid/',
    ...overrides,
  } as unknown as Activity
}

const channelActivity = (id: string) =>
  activity({
    id,
    text: '<at>Bot</at> hi',
    conversation: { id: '19:chan@thread.tacv2', conversationType: 'channel', tenantId: TENANT_ID } as any,
    entities: [{ type: 'mention', mentioned: { id: 'bot-1', name: 'Bot' }, text: '<at>Bot</at>' }] as any,
  })

function setup() {
  const dir = mkdtempSync(join(tmpdir(), 'cct-chansafe-'))
  const pending = createPendingStore(join(dir, 'pending.json'))
  const events: any[] = []
  const sent: string[] = []
  const handler = makeTurnHandler({
    config: makeConfig() as any,
    adapter: {} as any,
    allowlist: createAllowlist(join(dir, 'allow.json')),
    refs: createConversationRefStore(),
    onEvent: e => events.push(e),
    pending,
    sendPairingDm: async (_ref, text) => {
      sent.push(text)
    },
  })
  return { dir, pending, events, sent, handler }
}

describe('isPersonalReference', () => {
  test('classifies conversations', () => {
    expect(isPersonalReference({ conversation: { id: 'a:1xyz', conversationType: 'personal' } as any })).toBe(true)
    expect(isPersonalReference({ conversation: { id: 'a:1', conversationType: 'meeting', isGroup: false } as any })).toBe(false)
    expect(isPersonalReference({ conversation: { id: 'a:1', conversationType: '', isGroup: false } as any })).toBe(false)
    // An explicit 'personal' type wins over the id shape.
    expect(isPersonalReference({ conversation: { id: '19:x@thread.v2', conversationType: 'personal' } as any })).toBe(true)
    expect(isPersonalReference({ conversation: { id: 'a:1xyz' } as any })).toBe(true)
    expect(isPersonalReference({ conversation: { id: '19:x@thread.tacv2' } as any })).toBe(false)
    expect(isPersonalReference({ conversation: { id: '19:x@thread.v2', conversationType: 'groupChat' } as any })).toBe(false)
    expect(isPersonalReference({ conversation: { id: 'a:1', isGroup: true } as any })).toBe(false)
    expect(isPersonalReference({ conversation: { id: 'a:1', conversationType: 'channel' } as any })).toBe(false)
    expect(isPersonalReference(undefined)).toBe(false)
    expect(isPersonalReference({})).toBe(false)
  })
})

describe('mostRecentPersonalConversationId', () => {
  test('skips shared conversations', async () => {
    const refs = createConversationRefStore()
    refs.put('a:1xyz', { conversation: { id: 'a:1xyz', conversationType: 'personal' } as any }, SENDER_ID)
    await Bun.sleep(5)
    refs.put('19:c@thread.tacv2', { conversation: { id: '19:c@thread.tacv2', conversationType: 'channel' } as any }, SENDER_ID)
    expect(refs.mostRecentConversationId()).toBe('19:c@thread.tacv2')
    expect(refs.mostRecentPersonalConversationId()).toBe('a:1xyz')

    const only = createConversationRefStore()
    only.put('19:c@thread.tacv2', { conversation: { id: '19:c@thread.tacv2', conversationType: 'channel' } as any }, SENDER_ID)
    expect(only.mostRecentPersonalConversationId()).toBeUndefined()
  })
})

describe('pairing in shared conversations', () => {
  test('channel: no pending entry, one hint, no code, no event', async () => {
    const { dir, pending, events, sent, handler } = setup()
    await handler(new TurnContext(new NoopAdapter() as any, channelActivity('m1')))
    expect(pending.list().length).toBe(0)
    expect(sent).toEqual(['This bot is gated. To pair, send me a direct (1:1) message.'])
    expect(sent[0]).not.toMatch(/\b[0-9a-f]{6}\b/i)
    await handler(new TurnContext(new NoopAdapter() as any, channelActivity('m2')))
    expect(sent.length).toBe(1)
    expect(pending.list().length).toBe(0)
    expect(events).toEqual([])
    rmSync(dir, { recursive: true })
  })

  test('1:1: existing pairing path with code', async () => {
    const { dir, pending, sent, handler } = setup()
    await handler(new TurnContext(new NoopAdapter() as any, activity()))
    expect(pending.list().length).toBe(1)
    expect(sent.length).toBe(1)
    expect(sent[0]).toContain(`/teams:access pair ${pending.list()[0]!.code}`)
    rmSync(dir, { recursive: true })
  })

  test('an untyped @thread.v2 group chat is shared: no pending entry, the hint, no code', async () => {
    const { dir, pending, sent, handler } = setup()
    const other = '00000000-0000-4000-8000-0000000000aa'
    await handler(new TurnContext(new NoopAdapter() as any, activity({
      from: { id: 'user-2', name: 'Other', aadObjectId: other },
      conversation: { id: '19:abc@thread.v2', tenantId: TENANT_ID } as any,
    })))
    expect(pending.list().length).toBe(0)
    expect(sent).toEqual(['This bot is gated. To pair, send me a direct (1:1) message.'])
    rmSync(dir, { recursive: true })
  })

  test('an explicit personal type stays 1:1 for pairing, whatever the id', async () => {
    const { dir, pending, sent, handler } = setup()
    await handler(new TurnContext(new NoopAdapter() as any, activity({
      from: { id: 'user-3', name: 'Third', aadObjectId: '00000000-0000-4000-8000-0000000000bb' },
      conversation: { id: '19:x@thread.v2', conversationType: 'personal', tenantId: TENANT_ID } as any,
    })))
    expect(pending.list().length).toBe(1)
    expect(sent[0]).toContain(`/teams:access pair ${pending.list()[0]!.code}`)
    rmSync(dir, { recursive: true })
  })
})
