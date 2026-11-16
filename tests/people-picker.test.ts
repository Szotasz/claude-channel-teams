import { describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { validateAdaptiveCard } from '../src/teams/cards'
import { createAllowlist } from '../src/pairing/allowlist'
import { createPendingStore } from '../src/pairing/pair'
import { createConversationRefStore } from '../src/teams/conversationRefs'

const SENDER = '00000000-0000-4000-8000-000000000001'
const ID1 = 'aaaaaaaa-0000-4000-8000-000000000011'
const ID2 = 'bbbbbbbb-0000-4000-8000-000000000012'
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
    from: { id: 'user-1', name: 'Test User', aadObjectId: SENDER },
    conversation: { id: 'conv-abc', tenantId: TENANT },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    value: { action: { type: 'Action.Execute', title: 'Go', verb: 'go', data: { tender: 'T-1' } } },
    ...overrides,
  } as unknown as Activity
}

function setup(extra: Record<string, unknown> = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-picker-'))
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

const picker = (dataset: string, action: Record<string, unknown>) => ({
  type: 'AdaptiveCard', version: '1.5',
  body: [{ type: 'Input.ChoiceSet', id: 'assignee', choices: [], 'choices.data': { type: 'Data.Query', dataset }, isMultiSelect: true }],
  actions: [action],
})

const executeValue = { action: { type: 'Action.Execute', title: 'Assign', verb: 'assign', data: { assignee: `${ID1},${ID2}` } } }

describe('people picker card', () => {
  test('P1 the card check accepts the picker card', () => {
    expect(validateAdaptiveCard(picker('graph.microsoft.com/users', { type: 'Action.Submit', title: 'Assign' })).ok).toBe(true)
    expect(
      validateAdaptiveCard(picker('graph.microsoft.com/users?scope=currentContext', { type: 'Action.Execute', title: 'Assign', verb: 'assign' })).ok,
    ).toBe(true)
  })

  test('P2 an Action.Execute click carries the picked ids verbatim', async () => {
    const t = setup()
    await t.run(act({ value: executeValue }))
    expect(t.events.length).toBe(1)
    expect(t.events[0].cardAction).toEqual({ kind: 'execute', verb: 'assign' })
    expect(t.events[0].text).toBe(`[Card action] title="Assign" verb="assign"\ndata: {"assignee":"${ID1},${ID2}"}`)
    t.cleanup()
  })

  test('P2 an Action.Submit click carries the picked ids verbatim', async () => {
    const t = setup()
    await t.run(act({ type: 'message', name: undefined, text: undefined, value: { assignee: `${ID1},${ID2}` } }))
    expect(t.events.length).toBe(1)
    expect(t.events[0].cardAction).toEqual({ kind: 'submit' })
    expect(t.events[0].text).toBe(`[Card action]\ndata: {"assignee":"${ID1},${ID2}"}`)
    t.cleanup()
  })

  test('a click from a sender not on the allowlist delivers nothing', async () => {
    const t = setup()
    await t.run(act({ value: executeValue, from: { id: 'u2', name: 'Test User', aadObjectId: UNKNOWN } }))
    expect(t.events).toEqual([])
    const answers = t.bot.sent.filter((x: any) => x.type === 'invokeResponse').map((x: any) => x.value.body.value)
    expect(answers).toEqual(['Not allowed.'])
    t.cleanup()
  })
})
