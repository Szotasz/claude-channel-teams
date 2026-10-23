import { describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { buildDialogTask, dialogInvokeFrom, MAX_CARD_DATA_CHARS } from '../src/teams/cardAction'
import { createAllowlist } from '../src/pairing/allowlist'
import { createPendingStore } from '../src/pairing/pair'
import { createConversationRefStore } from '../src/teams/conversationRefs'

const CARD = { type: 'AdaptiveCard', version: '1.5', body: [{ type: 'TextBlock', text: 'Hi' }] }
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
    name: 'task/fetch',
    id: 'inv-1',
    replyToId: 'card-msg-1',
    timestamp: new Date('2026-05-18T19:43:00Z'),
    from: { id: 'user-1', name: 'Test User', aadObjectId: SENDER },
    conversation: { id: 'conv-abc', tenantId: TENANT },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    value: { data: { msteams: { type: 'task/fetch' }, dialog: { card: CARD } } },
    ...overrides,
  } as unknown as Activity
}

function setup(sender = SENDER, extra: Record<string, unknown> = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-dialog-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(SENDER)
  const pending = createPendingStore(join(dir, 'pending.json'))
  const events: any[] = []
  const bot = new RecordingAdapter()
  const turn = makeTurnHandler({
    config: {
      stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'a', appPassword: 'p', appType: 'SingleTenant',
      tenantId: TENANT, bindHost: '127.0.0.1', port: 0, receivedFilesDir: '/tmp/unused',
    } as any,
    adapter: {} as any,
    allowlist,
    refs: createConversationRefStore(),
    onEvent: e => events.push(e),
    pending,
    ...extra,
  })
  const run = (a: Activity) => turn(new TurnContext(bot as any, a))
  return { run, events, bot, pending, cleanup: () => rmSync(dir, { recursive: true }) }
}

const body = (t: any, i = 0) => t.bot.sent[i].value.body
const submit = (data: unknown, o: Record<string, unknown> = {}) => act({ name: 'task/submit', value: { data }, ...o })

describe('dialogInvokeFrom', () => {
  test('fetch / submit / other', () => {
    expect(dialogInvokeFrom({ type: 'invoke', name: 'task/fetch', value: { data: { dialog: { a: 1 } } } })).toEqual({ kind: 'fetch', dialog: { a: 1 } })
    expect(dialogInvokeFrom({ type: 'invoke', name: 'task/fetch', value: 'x' })).toEqual({ kind: 'fetch', dialog: undefined })
    expect(dialogInvokeFrom({ type: 'invoke', name: 'task/submit', value: { data: { q: 'a' } } })).toEqual({ kind: 'submit', dataJson: '{"q":"a"}' })
    expect(dialogInvokeFrom({ type: 'invoke', name: 'adaptiveCard/action', value: {} })).toBeUndefined()
    expect(dialogInvokeFrom({ type: 'message', name: 'task/fetch' })).toBeUndefined()
  })
})

describe('buildDialogTask', () => {
  test('valid card, defaults', () => {
    expect(buildDialogTask({ card: CARD, title: ' My  dialog ' })).toEqual({
      ok: true,
      task: { title: 'My dialog', width: 'medium', height: 'medium', card: { contentType: 'application/vnd.microsoft.card.adaptive', content: CARD } },
    })
    const r: any = buildDialogTask({ card: CARD })
    expect('title' in r.task).toBe(false)
  })
  test('title cleaned and cut; sizes', () => {
    const r: any = buildDialogTask({ card: CARD, title: 'a\u0007b<c>' + 'x'.repeat(200), width: 'large' })
    expect(r.task.title.length).toBeLessThanOrEqual(100)
    expect(r.task.title.startsWith('a b c xxx')).toBe(true)
    expect(r.task.width).toBe('large')
    expect(buildDialogTask({ card: CARD, width: 500 }).ok).toBe(false)
    expect(buildDialogTask({ card: CARD, height: 'huge' }).ok).toBe(false)
  })
  test('invalid input', () => {
    expect(buildDialogTask({ card: { ...CARD, version: '2.0' } }).ok).toBe(false)
    expect(buildDialogTask({ card: { ...CARD, body: [{ type: 'Image', url: 'http://x.example/a.png' }] } }).ok).toBe(false)
    expect(buildDialogTask('x').ok).toBe(false)
    expect(buildDialogTask(null).ok).toBe(false)
    expect(buildDialogTask({ title: 't' }).ok).toBe(false)
  })
})

describe('adapter: dialogs', () => {
  test('fetch opens the dialog, no event', async () => {
    const t = setup()
    await t.run(act())
    expect(t.events.length).toBe(0)
    expect(t.bot.sent.length).toBe(1)
    expect(t.bot.sent[0].type).toBe('invokeResponse')
    expect(t.bot.sent[0].value.status).toBe(200)
    expect(body(t).task.type).toBe('continue')
    expect(body(t).task.value.card.content).toEqual(CARD)
    t.cleanup()
  })
  test('fetch with an invalid dialog', async () => {
    const t = setup()
    await t.run(act({ value: { data: { dialog: { card: { type: 'AdaptiveCard', version: '2.0' } } } } }))
    expect(t.events.length).toBe(0)
    expect(t.bot.sent.length).toBe(1)
    expect(body(t)).toEqual({ task: { type: 'message', value: 'This dialog cannot be opened.' } })
    t.cleanup()
  })
  test('fetch / submit from a sender not on the allowlist', async () => {
    const t = setup()
    const from = { id: 'u', name: 'X', aadObjectId: UNKNOWN }
    await t.run(act({ from }))
    await t.run(submit({ a: 1 }, { from }))
    expect(t.events.length).toBe(0)
    expect(t.bot.sent.length).toBe(2)
    for (const i of [0, 1]) expect(body(t, i)).toEqual({ task: { type: 'message', value: 'Not allowed.' } })
    t.cleanup()
  })
  test('submit sends one event', async () => {
    const t = setup()
    await t.run(submit({ name: 'Bob', n: 2 }))
    expect(t.events.length).toBe(1)
    expect(t.events[0].text).toBe('[Dialog submit]\ndata: {"name":"Bob","n":2}')
    expect(t.events[0].cardAction).toEqual({ kind: 'dialog' })
    expect(t.events[0].messageId).toBe('card-msg-1')
    expect(t.bot.sent.length).toBe(1)
    expect(body(t)).toEqual({ task: { type: 'message', value: 'Sent.' } })
    t.cleanup()
  })
  test('oversized submit data is truncated', async () => {
    const t = setup()
    await t.run(submit({ big: 'y'.repeat(5000) }))
    const line = t.events[0].text.split('\ndata: ')[1]
    expect(line.length).toBeLessThanOrEqual(MAX_CARD_DATA_CHARS)
    expect(JSON.parse(line).truncated).toBe(true)
    t.cleanup()
  })
  test('early drop still gets a task answer', async () => {
    const t = setup()
    await t.run(act({ from: { id: 'u', name: 'X' } }))
    expect(t.events.length).toBe(0)
    expect(t.bot.sent.length).toBe(1)
    expect(body(t)).toEqual({ task: { type: 'message', value: 'Not allowed.' } })
    t.cleanup()
  })
  test('throwing onEvent leaves one task answer', async () => {
    const t = setup(SENDER, { onEvent: () => { throw new Error('boom') } })
    await expect(t.run(submit({ a: 1 }))).rejects.toThrow('boom')
    expect(t.bot.sent.length).toBe(1)
    expect(body(t)).toEqual({ task: { type: 'message', value: 'Not allowed.' } })
    t.cleanup()
  })
  test('adaptiveCard/action keeps the card-format answer', async () => {
    const t = setup()
    await t.run(act({ name: 'adaptiveCard/action', value: { action: { type: 'Action.Execute', data: {} } } }))
    expect(t.bot.sent.length).toBe(1)
    expect(body(t).value).toBe('Received.')
    t.cleanup()
  })
})
