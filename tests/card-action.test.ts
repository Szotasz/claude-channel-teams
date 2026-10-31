import { describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { makeEarlyAckTurn } from '../src/teams/earlyAck'
import { cardActionFrom, cardActionText, cleanLabel } from '../src/teams/cardAction'
import { createAllowlist } from '../src/pairing/allowlist'
import { createPendingStore } from '../src/pairing/pair'
import { createConversationRefStore } from '../src/teams/conversationRefs'

const SENDER = '00000000-0000-4000-8000-000000000001'
const UNKNOWN = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
const TENANT = '00000000-0000-4000-8000-000000000002'

class RecordingAdapter extends BotAdapter {
  sent: any[] = []
  async sendActivities(_ctx: any, activities: any[]): Promise<any> {
    this.sent.push(...activities)
    return activities.map(() => ({ id: 'x' }))
  }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function act(overrides: Record<string, unknown> = {}): Activity {
  return {
    type: 'invoke',
    name: 'adaptiveCard/action',
    id: 'inv-1',
    replyToId: 'card-msg-1',
    timestamp: new Date('2026-05-18T19:43:00Z'),
    from: { id: 'user-1', name: 'Adnan Khan', aadObjectId: SENDER },
    conversation: { id: 'conv-abc', tenantId: TENANT },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    value: { action: { type: 'Action.Execute', title: 'Go', verb: 'go', data: { tender: 'T-1' } } },
    ...overrides,
  } as unknown as Activity
}

function setup(sender = SENDER, extra: Record<string, unknown> = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-card-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(SENDER)
  const pending = createPendingStore(join(dir, 'pending.json'))
  const events: any[] = []
  const bot = new RecordingAdapter()
  const refs = createConversationRefStore()
  const turn = makeTurnHandler({
    config: {
      stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'a', appPassword: 'p', appType: 'SingleTenant',
      tenantId: TENANT, bindHost: '127.0.0.1', port: 0, receivedFilesDir: '/tmp/unused',
    } as any,
    adapter: {} as any,
    allowlist,
    refs,
    onEvent: e => events.push(e),
    pending,
    ...extra,
  })
  const run = (a: Activity) => turn(new TurnContext(bot as any, a))
  return { run, turn, events, bot, pending, refs, cleanup: () => rmSync(dir, { recursive: true }) }
}

describe('adapter: card actions', () => {
  test('A1 execute invoke from an allowlisted sender', async () => {
    const t = setup()
    await t.run(act())
    expect(t.events.length).toBe(1)
    expect(t.events[0].text).toBe('[Card action] title="Go" verb="go"\ndata: {"tender":"T-1"}')
    expect(t.events[0].cardAction).toEqual({ kind: 'execute', verb: 'go' })
    expect(t.events[0].messageId).toBe('card-msg-1')
    expect(t.bot.sent.length).toBe(1)
    expect(t.bot.sent[0].type).toBe('invokeResponse')
    expect(t.bot.sent[0].value.status).toBe(200)
    expect(t.bot.sent[0].value.body.value).toBe('Received.')
    t.cleanup()
  })

  test('A2 no @mention needed in a team channel', async () => {
    const t = setup()
    await t.run(act({ conversation: { id: '19:x@thread.tacv2', conversationType: 'channel', tenantId: TENANT } }))
    expect(t.events.length).toBe(1)
    t.cleanup()
  })

  test('A3 sender not on the allowlist', async () => {
    const t = setup()
    await t.run(act({ from: { id: 'u2', name: 'Eve', aadObjectId: UNKNOWN } }))
    expect(t.events).toEqual([])
    expect(t.bot.sent.length).toBe(1)
    expect(t.bot.sent[0].value.body.value).toBe('Not allowed.')
    expect(t.pending.list?.() ?? []).toEqual([])
    t.cleanup()
  })

  test('A4 submit never reaches the permission-reply match', async () => {
    const calls: any[] = []
    const t = setup(SENDER, { permission: { resolve: (...args: any[]) => (calls.push(args), true) } })
    await t.run(act({ type: 'message', name: undefined, text: '', value: { text: 'y abcde' } }))
    expect(calls).toEqual([])
    expect(t.events.length).toBe(1)
    expect(t.events[0].cardAction.kind).toBe('submit')
    expect(t.events[0].text.startsWith('[Card action]\ndata: {"text":"y abcde"}')).toBe(true)
    expect(t.bot.sent).toEqual([])
    t.cleanup()
  })

  test('A5 plain message takes the normal path', async () => {
    const t = setup()
    await t.run(act({ type: 'message', name: undefined, text: 'hello', value: undefined }))
    expect(t.events.length).toBe(1)
    expect(t.events[0].text).toBe('hello')
    expect(t.events[0].cardAction).toBeUndefined()
    t.cleanup()
  })

  test('A6 non-card invoke is ignored', async () => {
    const t = setup()
    await t.run(act({ name: 'composeExtension/query' }))
    expect(t.events).toEqual([])
    expect(t.bot.sent).toEqual([])
    t.cleanup()
  })

  test('A10 a Submit click from a sender not on the allowlist is dropped', async () => {
    const t = setup()
    await t.run(act({ type: 'message', name: undefined, text: '', value: { text: 'y abcde' }, from: { id: 'u2', name: 'Eve', aadObjectId: UNKNOWN } }))
    expect(t.events).toEqual([])
    expect(t.bot.sent).toEqual([])
    expect(t.pending.list?.() ?? []).toEqual([])
    expect(t.refs.get('conv-abc')).toBeUndefined()
    t.cleanup()
  })
})

describe('adapter: every Action.Execute invoke is answered', () => {
  test('A7 an invoke without an aadObjectId, or from another tenant, gets Not allowed and no event', async () => {
    const t = setup()
    await t.run(act({ from: { id: 'user-1', name: 'X' } }))
    await t.run(act({ conversation: { id: 'conv-abc', tenantId: '99999999-0000-4000-8000-000000000009' } }))
    expect(t.events).toEqual([])
    const answers = t.bot.sent.filter((x: any) => x.type === 'invokeResponse').map((x: any) => x.value.body.value)
    expect(answers).toEqual(['Not allowed.', 'Not allowed.'])
    t.cleanup()
  })
  test('A8 an automatic refresh invoke is not a click', async () => {
    const t = setup()
    await t.run(act({ value: { trigger: 'automatic', action: { type: 'Action.Execute', verb: 'confirm_delete', title: 'Yes, delete' } } }))
    expect(t.events).toEqual([])
    const answers = t.bot.sent.filter((x: any) => x.type === 'invokeResponse').map((x: any) => x.value.body.value)
    expect(answers).toEqual(['Ignored.'])
    t.cleanup()
  })
  test('A9 a manual trigger is a click', async () => {
    const t = setup()
    await t.run(act({ value: { trigger: 'manual', action: { type: 'Action.Execute', verb: 'go', title: 'Go', data: {} } } }))
    expect(t.events.length).toBe(1)
    t.cleanup()
  })
})

describe('cardAction helpers', () => {
  test('U1', () => {
    const c = cardActionFrom({
      type: 'invoke', name: 'adaptiveCard/action',
      value: { action: { title: 'a "q" <b>x</b>\nline', verb: 'v'.repeat(300), data: 'd'.repeat(3000) } },
    })!
    expect(c.title).toBe('a q b x /b line')
    expect(c.verb).toBe('v'.repeat(100))
    expect(c.dataJson.length).toBeLessThanOrEqual(2000)
    const parsed = JSON.parse(c.dataJson)
    expect(parsed.truncated).toBe(true)
    expect(JSON.stringify('d'.repeat(3000)).startsWith(parsed.start)).toBe(true)
    expect(parsed.start.length).toBeGreaterThan(1900)
    const quotes = cardActionFrom({ type: 'invoke', name: 'adaptiveCard/action', value: { action: { data: '"'.repeat(3000) } } })!
    expect(quotes.dataJson.length).toBeLessThanOrEqual(2000)
    expect(JSON.parse(quotes.dataJson).truncated).toBe(true)
    expect(cleanLabel('\u200fGo\u200e\u061c')).toBe('Go')

    const none = cardActionFrom({ type: 'invoke', name: 'adaptiveCard/action' })!
    expect(none).toEqual({ kind: 'execute', title: undefined, verb: undefined, dataJson: 'null' })
    expect(cardActionText(none)).toBe('[Card action]\ndata: null')

    expect(cardActionFrom({ type: 'message', value: [1] })).toBeUndefined()
    expect(cardActionFrom({ type: 'message', value: 'x' })).toBeUndefined()
    expect(cardActionText({ kind: 'execute', title: 'Go', verb: 'go', dataJson: '{"tender":"T-1"}' })).toBe(
      '[Card action] title="Go" verb="go"\ndata: {"tender":"T-1"}',
    )
  })
})

describe('adapter: card actions behind the early-ack queue', () => {
  // A message turn runs with its activity only (makeEarlyAckTurn); an invoke keeps the whole context, so it is answered.
  test('E1 an execute invoke through makeEarlyAckTurn is answered', async () => {
    const t = setup()
    const turn = makeEarlyAckTurn(t.turn)
    await turn(new TurnContext(t.bot as any, act()))
    expect(t.events.length).toBe(1)
    expect(t.bot.sent.length).toBe(1)
    expect(t.bot.sent[0].type).toBe('invokeResponse')
    expect(t.bot.sent[0].value.body.value).toBe('Received.')
    t.cleanup()
  })

  test('E2 a submit through makeEarlyAckTurn reaches Claude and sends nothing', async () => {
    const t = setup()
    const turn = makeEarlyAckTurn(t.turn)
    await turn(new TurnContext(t.bot as any, act({ type: 'message', name: undefined, text: '', value: { a: 1 } })))
    expect(t.events.length).toBe(1)
    expect(t.events[0].cardAction.kind).toBe('submit')
    expect(t.bot.sent).toEqual([])
    t.cleanup()
  })

  test('E3 an invoke that arrives with its activity only is not answered, and nothing throws', async () => {
    const t = setup()
    await t.turn({ activity: act() })
    expect(t.events.length).toBe(1)
    expect(t.bot.sent).toEqual([])
    t.cleanup()
  })
})
