import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { feedbackEventText, isFeedbackInvoke, MAX_FEEDBACK_CHARS, parseFeedback } from '../src/teams/feedback'

const USER = '00000000-0000-4000-8000-000000000001'
// Has hex letters, so a missing lower-casing is visible (USER is digits only).
const LETTERS = 'abcdef01-0000-4000-8000-0000000000ab'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const STRANGER = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'

const invoke = (o: Record<string, unknown> = {}) => ({
  type: 'invoke',
  name: 'message/submitAction',
  value: { actionName: 'feedback', actionValue: { reaction: 'like', feedback: '{"feedbackText":"Nice!"}' } },
  ...o,
})
const withValue = (actionValue: unknown) => ({ actionName: 'feedback', actionValue })

describe('isFeedbackInvoke', () => {
  test('shapes', () => {
    expect(isFeedbackInvoke(invoke())).toBe(true)
    expect(isFeedbackInvoke(invoke({ name: 'other' }))).toBe(false)
    expect(isFeedbackInvoke(invoke({ value: { actionName: 'other' } }))).toBe(false)
    expect(isFeedbackInvoke(invoke({ type: 'message' }))).toBe(false)
    expect(isFeedbackInvoke(invoke({ value: undefined }))).toBe(false)
  })
})

describe('parseFeedback', () => {
  const fb = (text: unknown) => parseFeedback(withValue({ reaction: 'like', feedback: JSON.stringify({ feedbackText: text }) }))
  test('cases', () => {
    expect(parseFeedback(invoke().value)).toEqual({ reaction: 'like', comment: 'Nice!' })
    expect(parseFeedback(withValue({ reaction: 'dislike' }))).toEqual({ reaction: 'dislike', comment: '' })
    expect(parseFeedback(withValue({ reaction: 'like', feedback: '{nope' }))).toEqual({ reaction: 'like', comment: '' })
    expect(fb(5)).toEqual({ reaction: 'like', comment: '' })
    expect(parseFeedback(withValue({ reaction: 'love' }))).toBeUndefined()
    expect(parseFeedback({ actionName: 'feedback' })).toBeUndefined()
    expect(fb('a\u0007b\nc\u0085d')?.comment).toBe('a b\nc d')
    expect(fb('x'.repeat(2500))?.comment.length).toBe(MAX_FEEDBACK_CHARS)
    expect(fb('  hi  ')?.comment).toBe('hi')
    expect(fb('a\u2028b\u202ec\u2066d')?.comment).toBe('a b c d')
    // The cap counts code points, so an emoji at the edge is never cut in half.
    expect(fb('x'.repeat(1999) + '\u{1F600}\u{1F600}')?.comment).toBe('x'.repeat(1999) + '\u{1F600}')
  })
})

describe('feedbackEventText', () => {
  test('strings', () => {
    expect(feedbackEventText({ reaction: 'like', comment: '' })).toBe('[Feedback] like')
    expect(feedbackEventText({ reaction: 'dislike', comment: 'Too long' })).toBe('[Feedback] dislike\nToo long')
  })
})

class RecordingAdapter extends BotAdapter {
  sent: any[] = []
  async sendActivities(_c: any, acts: any[]): Promise<any> { this.sent.push(...acts); return acts.map(() => ({ id: 'x' })) }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function setup(feedback: boolean) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-fb-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(USER)
  allowlist.addEntry(LETTERS)
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
    feedback,
  } as any)
  const bot = new RecordingAdapter()
  const run = (a: Activity) => turn(new TurnContext(bot as any, a))
  return { run, events, pairing, bot }
}

function act(o: Record<string, unknown> = {}): Activity {
  return {
    ...invoke(),
    id: 'inv-1',
    replyToId: 'bot-msg-1',
    timestamp: new Date('2026-05-18T19:43:00Z'),
    from: { id: 'user-1', name: 'Test User', aadObjectId: USER },
    conversation: { id: 'a:personal1', tenantId: TENANT_ID, conversationType: 'personal' },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    ...o,
  } as unknown as Activity
}

const answered = (bot: RecordingAdapter) => {
  expect(bot.sent.length).toBe(1)
  expect({ type: bot.sent[0].type, value: bot.sent[0].value }).toEqual({ type: 'invokeResponse', value: { status: 200 } })
}

describe('feedback invoke in the adapter', () => {
  test('delivers an allowlisted like', async () => {
    const s = setup(true)
    await s.run(act())
    expect(s.events).toEqual([{
      text: '[Feedback] like\nNice!',
      conversationId: 'a:personal1',
      aadObjectId: USER,
      tenantId: TENANT_ID,
      fromName: 'Test User',
      messageId: 'bot-msg-1',
      ts: '2026-05-18T19:43:00.000Z',
      feedback: 'like',
    }])
    answered(s.bot)
  })

  test('flag off: answered, not delivered', async () => {
    const s = setup(false)
    await s.run(act())
    expect(s.events.length).toBe(0)
    answered(s.bot)
  })

  const dropped: [string, Record<string, unknown>][] = [
    ['sender not on the allowlist', { from: { id: 'u', name: 'X', aadObjectId: STRANGER } }],
    ['another tenant', { conversation: { id: 'a:personal1', tenantId: 'other-tenant', conversationType: 'personal' } }],
    ['unknown reaction', { value: withValue({ reaction: 'love' }) }],
    ['no aadObjectId', { from: { id: 'u', name: 'X' } }],
    ['no replyToId (no rated message)', { replyToId: undefined }],
    ['no conversation id', { conversation: { tenantId: TENANT_ID, conversationType: 'personal' } }],
  ]
  for (const [name, o] of dropped) {
    test(`${name}: answered, not delivered`, async () => {
      const s = setup(true)
      await s.run(act(o))
      expect(s.events.length).toBe(0)
      answered(s.bot)
    })
  }

  test('an upper-case AAD id is delivered lower-cased', async () => {
    const s = setup(true)
    await s.run(act({ from: { id: 'user-1', name: 'Test User', aadObjectId: LETTERS.toUpperCase() } }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].aadObjectId).toBe(LETTERS)
    answered(s.bot)
  })

  test('a pairing code as the comment is not a pairing', async () => {
    const s = setup(true)
    await s.run(act({
      from: { id: 'u', name: 'X', aadObjectId: STRANGER },
      value: withValue({ reaction: 'like', feedback: '{"feedbackText":"123456"}' }),
    }))
    expect(s.pairing).toEqual([])
    expect(s.events.length).toBe(0)
    answered(s.bot)
  })
})
