import { describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { createPendingStore } from '../src/pairing/pair'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { makeTurnHandler } from '../src/teams/adapter'
import { targetedMessagesEnabled } from '../src/config'
import {
  TARGETED_PROMPT_HINT,
  buildTargetedActivity,
  createTargetedSender,
  deliverPrompt,
  selectPromptConversation,
  targetedActivityUrl,
  type PromptDelivery,
} from '../src/teams/targeted'

const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const C = (n: number) => `00000000-0000-4000-8000-0000000000c${n}`

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function activity(overrides: Partial<Activity> = {}): Activity {
  return {
    type: 'message',
    text: 'hi',
    id: 'msg-1',
    timestamp: new Date('2026-05-17T19:30:00Z'),
    from: { id: 'user-1', name: 'Stranger', aadObjectId: C(5) },
    conversation: { id: 'a:1xyz', tenantId: TENANT_ID },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://example.invalid/',
    ...overrides,
  } as unknown as Activity
}

const channelActivity = (id: string, aad: string) =>
  activity({
    id,
    from: { id: 'user-x', name: 'Stranger', aadObjectId: aad },
    text: '<at>Bot</at> hi',
    conversation: { id: '19:chan@thread.tacv2', conversationType: 'channel', tenantId: TENANT_ID } as any,
    entities: [{ type: 'mention', mentioned: { id: 'bot-1', name: 'Bot' }, text: '<at>Bot</at>' }] as any,
  })

function setup(rejectTargeted = false, sendTargetedOverride?: (ref: any, text: string) => Promise<void>) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-targeted-'))
  const pending = createPendingStore(join(dir, 'pending.json'))
  const events: any[] = []
  const dms: string[] = []
  const targeted: { ref: any; text: string }[] = []
  const handler = makeTurnHandler({
    config: {
      stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', pendingFile: '/tmp/unused',
      appId: 'app-id', appPassword: 'pw', appType: 'SingleTenant', tenantId: TENANT_ID,
      bindHost: '127.0.0.1', port: 0, targetedMessages: true,
    } as any,
    adapter: {} as any,
    allowlist: createAllowlist(join(dir, 'allow.json')),
    refs: createConversationRefStore(),
    onEvent: e => events.push(e),
    pending,
    sendPairingDm: async (_r, text) => { dms.push(text) },
    sendPairingTargeted: sendTargetedOverride ?? (async (ref, text) => {
      targeted.push({ ref, text })
      if (rejectTargeted) throw new Error('boom')
    }),
  })
  return { dir, pending, events, dms, targeted, handler }
}

const channelRef = {
  channelId: 'msteams',
  serviceUrl: 'https://smba.example.invalid/teams/',
  bot: { id: '28:bot', name: 'Bot' },
  user: { id: '29:user-1', name: 'Op', aadObjectId: C(1) },
  conversation: { id: '19:abc@thread.tacv2', conversationType: 'channel' },
} as any

test('T1 targetedMessagesEnabled', () => {
  for (const v of ['1', 'true', ' TRUE ']) expect(targetedMessagesEnabled({ TEAMS_TARGETED_MESSAGES: v })).toBe(true)
  for (const v of [undefined, '', '0', 'yes', 'false']) expect(targetedMessagesEnabled({ TEAMS_TARGETED_MESSAGES: v })).toBe(false)
})

test('T2 targetedActivityUrl', () => {
  expect(targetedActivityUrl('https://smba.example.invalid/teams/', '19:abc@thread.tacv2')).toBe(
    'https://smba.example.invalid/teams/v3/conversations/19%3Aabc%40thread.tacv2/activities?isTargetedActivity=true',
  )
})

test('T3 buildTargetedActivity', () => {
  const a = buildTargetedActivity(channelRef, 'yo') as any
  expect(a.recipient).toEqual({ id: '29:user-1', name: 'Op', aadObjectId: C(1), isTargeted: true })
  expect(a.from).toEqual(channelRef.bot)
  expect(a.conversation).toEqual(channelRef.conversation)
  expect(a.type).toBe('message')
  expect(a.text).toBe('yo')
  expect(() => buildTargetedActivity({ ...channelRef, user: undefined }, 'x')).toThrow('targeted message needs a user id')
  expect(() => buildTargetedActivity({ ...channelRef, conversation: undefined }, 'x')).toThrow('targeted message needs a conversation id')
})

describe('T4 createTargetedSender', () => {
  const KEY = Symbol('ConnectorClientKey')
  function fake(status: number, withClient = true) {
    const reqs: any[] = []
    const calls: string[] = []
    const client = { sendRequest: async (r: any) => { reqs.push(r); return { status } } }
    const adapter = {
      continueConversationAsync: async (appId: string, _ref: any, logic: any) => {
        calls.push(appId)
        await logic({ adapter: { ConnectorClientKey: KEY }, turnState: new Map(withClient ? [[KEY, client]] : []) } as any)
      },
    }
    return { reqs, calls, adapter }
  }
  test('success', async () => {
    const f = fake(201)
    await createTargetedSender(f.adapter, 'app-1')(channelRef, 'hello')
    expect(f.reqs.length).toBe(1)
    const req = f.reqs[0]
    expect(req.method).toBe('POST')
    expect(req.url).toBe('https://smba.example.invalid/teams/v3/conversations/19%3Aabc%40thread.tacv2/activities?isTargetedActivity=true')
    const body = JSON.parse(req.body)
    expect(body.recipient.isTargeted).toBe(true)
    expect(body.recipient.id).toBe('29:user-1')
    expect(req.headers.get('content-type')).toBe('application/json')
    expect(f.calls).toEqual(['app-1'])
  })
  test('failures', async () => {
    await expect(createTargetedSender(fake(403).adapter, 'a')(channelRef, 'x')).rejects.toThrow('targeted send failed (HTTP 403)')
    await expect(createTargetedSender(fake(201, false).adapter, 'a')(channelRef, 'x')).rejects.toThrow('no connector client for this turn')
    const f = fake(201)
    await expect(createTargetedSender(f.adapter, 'a')({ ...channelRef, user: undefined }, 'x')).rejects.toThrow()
    expect(f.calls.length).toBe(0)
  })
})

const dmRef = { conversation: { id: 'a:dm', conversationType: 'personal' }, user: { id: 'u' } } as any
const chanRef = { ...channelRef, conversation: { id: '19:chan@thread.tacv2', conversationType: 'channel' } }

async function seeded() {
  const refs = createConversationRefStore()
  refs.put('a:dm', dmRef, C(1))
  await Bun.sleep(5)
  refs.put('19:chan@thread.tacv2', chanRef, C(2))
  return refs
}

test('T5 selectPromptConversation', async () => {
  const refs = await seeded()
  const allowed = new Set([C(1), C(2)])
  const base = { refs, isAllowed: (id: string) => allowed.has(id), sendPlain: async () => {} }
  const st = async () => {}
  expect(selectPromptConversation(base)).toBe('a:dm')
  expect(selectPromptConversation({ ...base, sendTargeted: st })).toBe('19:chan@thread.tacv2')
  expect(selectPromptConversation({ ...base, sendTargeted: st, isAllowed: id => id === C(1) })).toBeUndefined()
  expect(selectPromptConversation({ ...base, refs: createConversationRefStore(), sendTargeted: st })).toBeUndefined()
})

describe('T6 deliverPrompt', () => {
  async function mk(opts: { targeted?: 'ok' | 'fail' | 'none'; dm?: boolean; dmAllowed?: boolean }) {
    const refs = createConversationRefStore()
    if (opts.dm) { refs.put('a:dm', dmRef, C(1)); await Bun.sleep(5) }
    refs.put('19:chan@thread.tacv2', chanRef, C(2))
    const plain: any[] = []
    const targeted: any[] = []
    const d: PromptDelivery = {
      refs,
      isAllowed: id => (id === C(1) ? opts.dmAllowed !== false : true),
      sendPlain: async (ref, text) => { plain.push({ ref, text }) },
      sendTargeted: opts.targeted === 'none' ? undefined : async (ref, text) => {
        targeted.push({ ref, text })
        if (opts.targeted === 'fail') throw new Error('nope')
      },
    }
    return { d, plain, targeted }
  }
  test('a personal', async () => {
    const m = await mk({ dm: true, targeted: 'ok' })
    await deliverPrompt(m.d, 'a:dm', 'P')
    expect(m.plain).toEqual([{ ref: dmRef, text: 'P' }])
    expect(m.targeted.length).toBe(0)
  })
  test('b shared targeted ok', async () => {
    const m = await mk({ targeted: 'ok' })
    await deliverPrompt(m.d, '19:chan@thread.tacv2', 'P')
    expect(m.targeted).toEqual([{ ref: chanRef, text: 'P' + TARGETED_PROMPT_HINT }])
    expect(m.plain.length).toBe(0)
  })
  test('c fallback to 1:1', async () => {
    const m = await mk({ targeted: 'fail', dm: true })
    await deliverPrompt(m.d, '19:chan@thread.tacv2', 'P')
    expect(m.plain).toEqual([{ ref: dmRef, text: 'P' }])
  })
  test('d no 1:1', async () => {
    const m = await mk({ targeted: 'fail' })
    await expect(deliverPrompt(m.d, '19:chan@thread.tacv2', 'P')).rejects.toThrow(/no 1:1 chat/)
    expect(m.plain.length).toBe(0)
  })
  test('e 1:1 not allowed', async () => {
    const m = await mk({ targeted: 'fail', dm: true, dmAllowed: false })
    await expect(deliverPrompt(m.d, '19:chan@thread.tacv2', 'P')).rejects.toThrow()
    expect(m.plain.length).toBe(0)
  })
  test('f no sendTargeted', async () => {
    const m = await mk({ targeted: 'none', dm: true })
    await expect(deliverPrompt(m.d, '19:chan@thread.tacv2', 'P')).rejects.toThrow(/refusing to post/)
    expect(m.plain.length).toBe(0)
  })
  test('g unknown id', async () => {
    const m = await mk({ targeted: 'ok' })
    await expect(deliverPrompt(m.d, 'zzz', 'P')).rejects.toThrow(/no conversation reference/)
  })
})

describe('adapter with targeted messages', () => {
  test('T7 channel: targeted code, no plain DM', async () => {
    const { dir, pending, events, dms, targeted, handler } = setup()
    await handler(new TurnContext(new NoopAdapter() as any, channelActivity('m1', C(6))))
    const list = pending.list()
    expect(list.length).toBe(1)
    expect(list[0]!.conversation_id).toBe('19:chan@thread.tacv2')
    expect(targeted.length).toBe(1)
    expect(targeted[0]!.text).toContain(`/teams:access pair ${list[0]!.code}`)
    expect(targeted[0]!.ref.user.aadObjectId).toBe(C(6))
    expect(dms.length).toBe(0)
    expect(events).toEqual([])
    await handler(new TurnContext(new NoopAdapter() as any, channelActivity('m2', C(6))))
    expect(targeted.length).toBe(1)
    expect(pending.list().length).toBe(1)
    rmSync(dir, { recursive: true })
  })

  test('T8 targeted failure falls back to the hint', async () => {
    const { dir, pending, dms, targeted, handler } = setup(true)
    await handler(new TurnContext(new NoopAdapter() as any, channelActivity('m1', C(7))))
    await Bun.sleep(5)
    expect(targeted.length).toBe(1)
    expect(pending.list().length).toBe(0)
    expect(dms).toEqual(['This bot is gated. To pair, send me a direct (1:1) message.'])
    expect(dms[0]).not.toMatch(/\b[0-9a-f]{6}\b/i)
    rmSync(dir, { recursive: true })
  })

  test('T10 a late failure keeps an entry that changed since', async () => {
    let fail!: (err: Error) => void
    const { dir, pending, dms, handler } = setup(false, () => new Promise<void>((_, reject) => { fail = reject }))
    await handler(new TurnContext(new NoopAdapter() as any, channelActivity('m1', C(9))))
    const [entry] = pending.list()
    expect(entry).toBeDefined()
    // The targeted send still hangs when the sender comes back after the reminder interval.
    const later = pending.recordIncoming({
      aadObjectId: C(9), tenantId: TENANT_ID, fromName: 'Stranger', conversationId: 'a:dm-9',
      now: Date.now() + 11 * 60 * 1000,
    })
    expect(later.action).toBe('send_reminder')
    fail(new Error('late'))
    await Bun.sleep(5)
    expect(pending.findByPairId(entry!.pair_id)?.reply_count).toBe(2)
    expect(dms).toEqual(['This bot is gated. To pair, send me a direct (1:1) message.'])
    rmSync(dir, { recursive: true })
  })

  test('T11 a failed targeted reminder keeps the entry', async () => {
    const { dir, pending, targeted, handler } = setup(true)
    const aad = '00000000-0000-4000-8000-0000000000d1'
    // An entry from 11 minutes ago, so the next message is a reminder.
    const first = pending.recordIncoming({
      aadObjectId: aad, tenantId: TENANT_ID, fromName: 'Stranger', conversationId: '19:chan@thread.tacv2',
      now: Date.now() - 11 * 60 * 1000,
    })
    expect(first.action).toBe('send_initial')
    await handler(new TurnContext(new NoopAdapter() as any, channelActivity('m1', aad)))
    await Bun.sleep(5)
    expect(targeted.length).toBe(1)
    expect(targeted[0]!.text).toStartWith('Still pending')
    expect(pending.list().map(e => e.aad_object_id)).toEqual([aad])
    rmSync(dir, { recursive: true })
  })

  test('T9 1:1 unchanged', async () => {
    const { dir, pending, dms, targeted, handler } = setup()
    await handler(new TurnContext(new NoopAdapter() as any, activity({
      from: { id: 'user-9', name: 'Other', aadObjectId: C(8) } as any,
    })))
    expect(dms.length).toBe(1)
    expect(dms[0]).toContain(`/teams:access pair ${pending.list()[0]!.code}`)
    expect(targeted.length).toBe(0)
    rmSync(dir, { recursive: true })
  })
})
