import { describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { createReplySender } from '../src/teams/reply'
import { createAllowlist } from '../src/pairing/allowlist'
import { createPendingStore } from '../src/pairing/pair'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import {
  USER_VIEW_VERB,
  createKnownUsers,
  createUserViewStore,
  userViewInvokeFrom,
  validateUserCards,
  withRefresh,
} from '../src/teams/userViews'

const SENDER = '00000000-0000-4000-8000-000000000001'
const OTHER = '00000000-0000-4000-8000-000000000003'
const TENANT = '00000000-0000-4000-8000-000000000002'
// Has hex letters, so a missing lower-casing is visible (the ids above are digits only).
const LETTERS = 'abcdef01-0000-4000-8000-0000000000ab'
const CT = 'application/vnd.microsoft.card.adaptive'
const C = (t: string) => ({ type: 'AdaptiveCard', version: '1.5', body: [{ type: 'TextBlock', text: t }] })
const entry = (id = SENDER, card: unknown = C('mine')) => ({ aad_object_id: id, card })
const msg = (f: () => unknown) => {
  try {
    f()
  } catch (e) {
    return (e as Error).message
  }
  return undefined
}
const reservedCard = () => ({
  ...C('x'),
  body: [{ type: 'ActionSet', actions: [{ type: 'Action.Execute', verb: USER_VIEW_VERB }] }],
})

describe('validateUserCards', () => {
  test('non-array', () => expect(msg(() => validateUserCards('x', C('b')))).toBe('user_cards must be an array'))
  test('empty', () => expect(msg(() => validateUserCards([], C('b')))).toBe('user_cards must not be empty'))
  test('61 entries', () =>
    expect(msg(() => validateUserCards(Array(61).fill(entry()), C('b')))).toBe('user_cards can list at most 60 people'))
  test('no base card', () =>
    expect(msg(() => validateUserCards([entry()], undefined))).toBe('user_cards needs card (the card everyone else sees)'))
  test('base version 1.3', () =>
    expect(msg(() => validateUserCards([entry()], { ...C('b'), version: '1.3' }))).toBe(
      'card.version must be 1.4 or 1.5 when user_cards is set',
    ))
  test('base reserved verb', () =>
    expect(msg(() => validateUserCards([entry()], reservedCard()))).toBe(`card uses the reserved verb ${USER_VIEW_VERB}`))
  test('entry not an object', () =>
    expect(msg(() => validateUserCards(['x'], C('b')))).toBe('user_cards[0] must be an object'))
  test('bad aad', () =>
    expect(msg(() => validateUserCards([entry(), entry('nope')], C('b')))).toBe(
      'user_cards[1].aad_object_id must be an AAD object id',
    ))
  test('same aad twice', () => {
    expect(msg(() => validateUserCards([entry(LETTERS), entry(LETTERS.toUpperCase())], C('b')))).toBe(
      'user_cards[1].aad_object_id is listed twice',
    )
    expect(msg(() => validateUserCards([entry(LETTERS.toUpperCase()), entry(LETTERS)], C('b')))).toBe(
      'user_cards[1].aad_object_id is listed twice',
    )
  })
  test('60 entries are accepted', () => {
    const ids = Array.from({ length: 60 }, (_, i) => `00000000-0000-4000-8000-${String(i).padStart(12, '0')}`)
    expect(validateUserCards(ids.map(id => entry(id)), C('b')).length).toBe(60)
  })
  test('invalid entry card', () =>
    expect(msg(() => validateUserCards([entry(SENDER, { ...C('x'), type: 'Foo' })], C('b')))).toBe(
      'user_cards[0].card is invalid: card.type must be "AdaptiveCard"',
    ))
  test('entry version 1.0', () =>
    expect(msg(() => validateUserCards([entry(SENDER, { ...C('x'), version: '1.0' })], C('b')))).toBe(
      'user_cards[0].card.version must be 1.4 or 1.5',
    ))
  test('entry reserved verb', () =>
    expect(msg(() => validateUserCards([entry(SENDER, reservedCard())], C('b')))).toBe(
      `user_cards[0].card uses the reserved verb ${USER_VIEW_VERB}`,
    ))
  test('valid list', () => {
    const a = C('a')
    const b = C('b')
    expect(validateUserCards([entry(LETTERS.toUpperCase(), a), entry(OTHER, b)], C('base'))).toEqual([
      { aadObjectId: LETTERS, card: a },
      { aadObjectId: OTHER, card: b },
    ])
  })
})

describe('withRefresh', () => {
  test('exact refresh, input unchanged', () => {
    const card = C('x')
    const out = withRefresh(card, 'v1', ['29:a', '29:b'])
    expect(out).toEqual({
      ...C('x'),
      refresh: {
        action: { type: 'Action.Execute', title: 'Refresh', verb: USER_VIEW_VERB, data: { view: 'v1' } },
        userIds: ['29:a', '29:b'],
      },
    })
    expect(card).toEqual(C('x'))
  })
  test('userIds is a copy', () => {
    const mris = ['29:a']
    const out = withRefresh(C('x'), 'v1', mris)
    mris.push('29:b')
    expect((out.refresh as { userIds: string[] }).userIds).toEqual(['29:a'])
  })
})

describe('createKnownUsers', () => {
  test('lower-casing and empty mri', () => {
    const k = createKnownUsers()
    k.set('ABC', '29:x')
    expect(k.get('abc')).toBe('29:x')
    expect(k.get('AbC')).toBe('29:x')
    k.set('def', '')
    expect(k.get('def')).toBeUndefined()
    expect(k.size()).toBe(1)
  })
  test('evicts the oldest', () => {
    const k = createKnownUsers(2)
    k.set('a', '1')
    k.set('b', '2')
    k.set('c', '3')
    expect(k.get('a')).toBeUndefined()
    expect(k.get('b')).toBe('2')
    expect(k.get('c')).toBe('3')
    expect(k.size()).toBe(2)
  })
  test('set again moves to newest', () => {
    const k = createKnownUsers(2)
    k.set('a', '1')
    k.set('b', '2')
    k.set('a', '1')
    k.set('c', '3')
    expect(k.get('b')).toBeUndefined()
    expect(k.get('a')).toBe('1')
    expect(k.get('c')).toBe('3')
  })
})

describe('createUserViewStore', () => {
  const counter = () => {
    let n = 0
    return () => `v${++n}`
  }
  test('create / cardFor', () => {
    const s = createUserViewStore(500, counter())
    const r = s.create(C('base'), [{ aadObjectId: SENDER, card: C('mine') }], ['29:a'])
    expect(r.id).toBe('v1')
    expect(r.base).toEqual(withRefresh(C('base'), 'v1', ['29:a']))
    expect(s.cardFor('v1', SENDER.toUpperCase())).toEqual(withRefresh(C('mine'), 'v1', ['29:a']))
    expect(s.cardFor('v1', OTHER)).toEqual(withRefresh(C('base'), 'v1', ['29:a']))
    expect(s.cardFor('nope', SENDER)).toBeUndefined()
    expect(s.size()).toBe(1)
    s.create(C('base'), [{ aadObjectId: LETTERS, card: C('theirs') }], ['29:b'])
    expect(s.cardFor('v2', LETTERS.toUpperCase())).toEqual(withRefresh(C('theirs'), 'v2', ['29:b']))
  })
  test('evicts the oldest', () => {
    const s = createUserViewStore(2, counter())
    s.create(C('1'), [], [])
    s.create(C('2'), [], [])
    s.create(C('3'), [], [])
    expect(s.size()).toBe(2)
    expect(s.cardFor('v1', SENDER)).toBeUndefined()
    expect(s.cardFor('v2', SENDER)).toBeDefined()
    expect(s.cardFor('v3', SENDER)).toBeDefined()
  })
})

describe('userViewInvokeFrom', () => {
  const inv = (value: unknown, over: Record<string, unknown> = {}) => ({
    type: 'invoke',
    name: 'adaptiveCard/action',
    value,
    ...over,
  })
  const action = (data: unknown, verb = USER_VIEW_VERB) => ({ type: 'Action.Execute', verb, data })
  test('with and without trigger', () => {
    expect(userViewInvokeFrom(inv({ trigger: 'automatic', action: action({ view: 'v1' }) }))).toEqual({ viewId: 'v1' })
    expect(userViewInvokeFrom(inv({ action: action({ view: 'v1' }) }))).toEqual({ viewId: 'v1' })
  })
  test('bad view', () => {
    expect(userViewInvokeFrom(inv({ action: action({}) }))).toEqual({ viewId: undefined })
    expect(userViewInvokeFrom(inv({ action: action({ view: '' }) }))).toEqual({ viewId: undefined })
    expect(userViewInvokeFrom(inv({ action: action({ view: 5 }) }))).toEqual({ viewId: undefined })
  })
  test('other verb / message', () => {
    expect(userViewInvokeFrom(inv({ action: action({ view: 'v1' }, 'go') }))).toBeUndefined()
    expect(userViewInvokeFrom(inv({ action: action({ view: 'v1' }) }, { type: 'message' }))).toBeUndefined()
    expect(userViewInvokeFrom(inv({ action: action({ view: 'v1' }) }, { name: 'task/fetch' }))).toBeUndefined()
  })
})

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

function setup() {
  const dir = mkdtempSync(join(tmpdir(), 'cct-uv-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(SENDER)
  const pending = createPendingStore(join(dir, 'pending.json'))
  const events: any[] = []
  const bot = new RecordingAdapter()
  const knownUsers = createKnownUsers()
  const userViews = createUserViewStore(500, () => 'view-1')
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
    knownUsers,
    userViews,
  })
  const run = (a: Activity) => turn(new TurnContext(bot as any, a))
  return { run, events, bot, knownUsers, userViews, cleanup: () => rmSync(dir, { recursive: true }) }
}

const viewInvoke = (view: string, trigger: string | undefined = 'automatic') =>
  act({
    value: {
      ...(trigger ? { trigger } : {}),
      action: { type: 'Action.Execute', verb: USER_VIEW_VERB, data: { view } },
    },
  })

describe('adapter: user views', () => {
  const answer = (value: unknown) => ({
    type: 'invokeResponse',
    value: { status: 200, body: { statusCode: 200, type: CT, value } },
  })
  // TurnContext adds the conversation reference fields to every sent activity; compare what the bot chose.
  const sentShape = (t: { bot: RecordingAdapter }) =>
    t.bot.sent.map((x: any) => ({ type: x.type, value: x.value }))

  test('listed sender gets their card', async () => {
    const t = setup()
    const { id } = t.userViews.create(C('base'), [{ aadObjectId: SENDER, card: C('mine') }], ['user-1'])
    await t.run(viewInvoke(id))
    expect(sentShape(t)).toEqual([answer(withRefresh(C('mine'), id, ['user-1']))] as any)
    expect(t.events).toEqual([])
    t.cleanup()
  })

  test('unlisted sender gets the base card', async () => {
    const t = setup()
    const { id } = t.userViews.create(C('base'), [{ aadObjectId: OTHER, card: C('theirs') }], ['u'])
    await t.run(viewInvoke(id))
    expect(sentShape(t)).toEqual([answer(withRefresh(C('base'), id, ['u']))] as any)
    expect(t.events).toEqual([])
    t.cleanup()
  })

  test('unknown view is ignored', async () => {
    const t = setup()
    await t.run(viewInvoke('nope'))
    expect(t.bot.sent.length).toBe(1)
    expect(t.bot.sent[0].type).toBe('invokeResponse')
    expect(t.bot.sent[0].value.body.value).toBe('Ignored.')
    expect(t.events).toEqual([])
    t.cleanup()
  })

  test('a listed person no longer on the allowlist gets the base card', async () => {
    const t = setup()
    const { id } = t.userViews.create(C('base'), [{ aadObjectId: OTHER, card: C('theirs') }], ['u'])
    await t.run({ ...viewInvoke(id), from: { id: 'user-2', name: 'Test User', aadObjectId: OTHER } } as Activity)
    expect(sentShape(t)).toEqual([answer(withRefresh(C('base'), id, ['u']))] as any)
    expect(t.events).toEqual([])
    t.cleanup()
  })

  test('another tenant is not allowed', async () => {
    const t = setup()
    const { id } = t.userViews.create(C('base'), [{ aadObjectId: SENDER, card: C('mine') }], ['user-1'])
    await t.run({ ...viewInvoke(id), conversation: { id: 'conv-abc', tenantId: OTHER } } as Activity)
    expect(t.bot.sent.length).toBe(1)
    expect(t.bot.sent[0].value.body.value).toBe('Not allowed.')
    expect(t.events).toEqual([])
    t.cleanup()
  })

  test('manual refresh (no trigger) is answered with the card', async () => {
    const t = setup()
    const { id } = t.userViews.create(C('base'), [{ aadObjectId: SENDER, card: C('mine') }], ['user-1'])
    await t.run(viewInvoke(id, undefined))
    expect(sentShape(t)).toEqual([answer(withRefresh(C('mine'), id, ['user-1']))] as any)
    expect(t.events).toEqual([])
    t.cleanup()
  })

  test('a normal message records the user MRI', async () => {
    const t = setup()
    await t.run(act({ type: 'message', name: undefined, text: 'hi', value: undefined }))
    expect(t.knownUsers.get(SENDER)).toBe('user-1')
    t.cleanup()
  })

  test('a card click records the user MRI and is still an event', async () => {
    const t = setup()
    await t.run(act())
    expect(t.knownUsers.get(SENDER)).toBe('user-1')
    expect(t.events.length).toBe(1)
    expect(t.events[0].cardAction).toEqual({ kind: 'execute', verb: 'go' })
    t.cleanup()
  })
})

describe('reply: user_cards', () => {
  function rsetup() {
    const dir = mkdtempSync(join(tmpdir(), 'cct-uvreply-'))
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(SENDER)
    const refs = createConversationRefStore()
    refs.put('conv-abc', { conversation: { id: 'conv-abc' } } as any, SENDER)
    const state = { calls: 0, sent: undefined as any }
    const adapter = {
      async continueConversationAsync(_a: string, _r: any, logic: any) {
        state.calls++
        await logic({
          async sendActivity(act: any) {
            state.sent = act
            return { id: 'm1' }
          },
        })
      },
    } as any
    const knownUsers = createKnownUsers()
    knownUsers.set(SENDER, '29:mri-a')
    const userViews = createUserViewStore(500, () => 'view-9')
    const config = {
      stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'a', appPassword: 'p', appType: 'SingleTenant' as const,
      tenantId: TENANT, bindHost: '127.0.0.1', port: 0,
    }
    const sender = createReplySender({ config, adapter, allowlist, refs, knownUsers, userViews } as any)
    return { ...sender, state, userViews }
  }

  test('known person: refresh added, view stored', async () => {
    const s = rsetup()
    await s.sendReply('conv-abc', 'hi', C('base'), { userCards: [entry(SENDER, C('mine'))] })
    const content = s.state.sent.attachments[0].content
    expect(s.state.sent.attachments[0].contentType).toBe(CT)
    expect(content).toEqual(withRefresh(C('base'), 'view-9', ['29:mri-a']))
    expect(content.refresh.userIds).toEqual(['29:mri-a'])
    expect(content.refresh.action.verb).toBe(USER_VIEW_VERB)
    expect(content.refresh.action.data.view).toBe('view-9')
    expect(s.userViews.cardFor('view-9', SENDER)).toEqual(withRefresh(C('mine'), 'view-9', ['29:mri-a']))
  })

  test('unknown person is refused', async () => {
    const s = rsetup()
    await expect(s.sendReply('conv-abc', 'hi', C('base'), { userCards: [entry(OTHER)] })).rejects.toThrow(
      `user_cards: no Teams user id is known yet for ${OTHER} (they have not written to the bot since it started)`,
    )
    expect(s.state.calls).toBe(0)
    expect(s.userViews.size()).toBe(0)
  })

  test('no view is stored when a later check refuses the reply', async () => {
    const s = rsetup()
    await expect(
      s.sendReply('conv-unknown', 'hi', C('base'), { userCards: [entry(SENDER, C('mine'))] }),
    ).rejects.toThrow('unknown conversation_id: conv-unknown')
    await expect(
      s.sendReply('conv-abc', 'hi', C('base'), { userCards: [entry(SENDER, C('mine'))], citations: 'x' as any }),
    ).rejects.toThrow('citations must be an array')
    expect(s.state.calls).toBe(0)
    expect(s.userViews.size()).toBe(0)
  })

  test('a plain card may not use the reserved verb either', async () => {
    const s = rsetup()
    await expect(s.sendReply('conv-abc', 'hi', reservedCard())).rejects.toThrow(
      `card uses the reserved verb ${USER_VIEW_VERB}`,
    )
    expect(s.state.calls).toBe(0)
  })

  test('user_cards without card', async () => {
    const s = rsetup()
    await expect(s.sendReply('conv-abc', 'hi', undefined, { userCards: [entry()] })).rejects.toThrow(
      'user_cards needs card (the card everyone else sees)',
    )
    expect(s.state.calls).toBe(0)
  })

  test('base card with its own refresh is refused', async () => {
    const s = rsetup()
    const base = { ...C('base'), refresh: { action: { type: 'Action.Execute', verb: 'x' } } }
    await expect(s.sendReply('conv-abc', 'hi', base, { userCards: [entry()] })).rejects.toThrow('is a refresh')
    expect(s.state.calls).toBe(0)
    expect(s.userViews.size()).toBe(0)
  })
})
