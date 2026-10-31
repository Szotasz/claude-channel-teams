import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import {
  ConversationNotAllowedError,
  LATEST_PRIVATE_MESSAGE,
  createReplySender,
  parseTargetedMessageId,
} from '../src/teams/reply'
import { createFileSender } from '../src/teams/sendFile'
import {
  MAX_CONVERSATIONS,
  MAX_ENTRIES_PER_CONVERSATION,
  MAX_SENT_IDS,
  TARGETED_TTL_MS,
  createTargetedInbound,
} from '../src/teams/targetedInbound'

const ALLOWED = '00000000-0000-4000-8000-000000000001'
const OTHER = '00000000-0000-4000-8000-000000000003'
const UNKNOWN_ID = 'cccccccc-bbbb-4ccc-8ddd-eeeeeeeeeee3' // unique: the 1:1 hint is sent once per sender per process
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const CHAN = '19:chan1@thread.tacv2;messageid=111'
const CHAN_B = '19:chan2@thread.tacv2;messageid=222'
const KEY = Symbol('ConnectorClientKey')
const MENTION = [{ type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-1', name: 'Bot' } }] as any

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function makeActivity(o: Partial<Activity> = {}): Activity {
  return {
    type: 'message',
    text: '/status',
    id: 'msg-1',
    timestamp: new Date('2026-05-18T19:43:00Z'),
    from: { id: '29:user-1', name: 'Adnan', aadObjectId: ALLOWED },
    conversation: { id: CHAN, tenantId: TENANT_ID, conversationType: 'channel' },
    recipient: { id: 'bot-1', name: 'Bot', isTargeted: true },
    channelId: 'msteams',
    serviceUrl: 'https://smba.example.invalid/teams/',
    ...o,
  } as unknown as Activity
}
const untargeted = (o: Partial<Activity> = {}) =>
  makeActivity({ recipient: { id: 'bot-1', name: 'Bot' } as any, ...o })
const mentionActivity = (o: Partial<Activity> = {}) =>
  untargeted({ text: '<at>Bot</at> hi', entities: MENTION, ...o })

function setup(opts: { targetedMessages?: boolean; status?: number; body?: string; pairing?: boolean } = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-targeted-inbound-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(ALLOWED)
  allowlist.addEntry(OTHER)
  const refs = createConversationRefStore()
  const targetedInbound = createTargetedInbound()
  const events: any[] = []
  const resolved: any[] = []
  const hints: string[] = []
  const config = {
    stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'app-id', appPassword: 'pw',
    appType: 'SingleTenant', tenantId: TENANT_ID, bindHost: '127.0.0.1', port: 0,
    receivedFilesDir: '/tmp/unused', targetedMessages: opts.targetedMessages ?? true,
  } as any
  let now = Date.now()
  const handler = makeTurnHandler({
    config, adapter: {} as any, allowlist, refs, targetedInbound,
    onEvent: e => events.push(e),
    permission: { resolve: (id: string, v: string) => { resolved.push([id, v]); return true } } as any,
    ...(opts.pairing
      ? {
          pending: { recordIncoming: () => { throw new Error('no pairing in shared conversation') } } as any,
          sendPairingDm: async (_r: any, t: string) => { hints.push(t) },
        }
      : {}),
  })
  const run = (a: Activity) => handler(new TurnContext(new NoopAdapter() as any, a))

  const posts: any[] = []
  const publicSends: any[] = []
  const edits: any[] = []
  const client = {
    sendRequest: async (r: any) => {
      posts.push(r)
      return { status: opts.status ?? 200, bodyAsText: opts.body ?? JSON.stringify({ id: 'sent-1' }) }
    },
  }
  const adapter = {
    continueConversationAsync: async (appId: string, _ref: any, logic: any) => {
      await logic({
        adapter: { ConnectorClientKey: KEY },
        turnState: new Map([[KEY, client]]),
        sendActivity: async (act: any) => { publicSends.push(act); return { id: 'pub-1' } },
        updateActivity: async (act: any) => { edits.push(act) },
      })
    },
  } as any
  const replier = createReplySender({ config, adapter, allowlist, refs, targetedInbound, now: () => now })
  return {
    run, events, resolved, hints, refs, allowlist, targetedInbound, replier, posts, publicSends, edits,
    adapter, config, dir, setNow: (n: number) => { now = n }, getNow: () => now,
  }
}

describe('inbound targeted detection', () => {
  test('B1 targeted channel message without @mention is delivered, targeted + mentioned', async () => {
    const s = setup()
    await s.run(makeActivity())
    expect(s.events.length).toBe(1)
    expect(s.events[0].targeted).toBe(true)
    expect(s.events[0].mentioned).toBe(true)
    expect(s.events[0].text).toBe('/status')
  })

  test('B2 same message without isTargeted is dropped', async () => {
    const s = setup()
    await s.run(untargeted())
    expect(s.events.length).toBe(0)
  })

  test('B3 personal message with isTargeted is handled as today', async () => {
    const s = setup()
    await s.run(makeActivity({ conversation: { id: 'a:1xyz', tenantId: TENANT_ID, conversationType: 'personal' } as any }))
    expect(s.events.length).toBe(1)
    expect('targeted' in s.events[0]).toBe(false)
    expect(s.events[0].mentioned).toBeUndefined()
    expect(s.targetedInbound.get('a:1xyz', 'msg-1', s.getNow())).toBeUndefined()
    expect(s.targetedInbound.latestIsTargeted('a:1xyz', s.getNow())).toBe(false)
  })

  test('B4 non-allowlisted sender: same as their @mention, nothing remembered', async () => {
    const stranger = { id: '29:u2', name: 'X', aadObjectId: UNKNOWN_ID } as any
    const t = setup({ pairing: true })
    await t.run(makeActivity({ from: stranger }))
    const m = setup({ pairing: true })
    // (a different stranger: the one-hint-per-sender memory is process-wide)
    await m.run(mentionActivity({ from: { ...stranger, aadObjectId: 'bbbbbbbb-bbbb-cccc-dddd-eeeeeeeeeeee' } }))
    expect(t.events.length).toBe(0)
    expect(t.hints).toEqual(['This bot is gated. To pair, send me a direct (1:1) message.'])
    expect(t.hints).toEqual(m.hints)
    expect(t.events).toEqual(m.events)
    expect(t.targetedInbound.get(CHAN, 'msg-1', t.getNow())).toBeUndefined()
    expect(t.targetedInbound.latestIsTargeted(CHAN, t.getNow())).toBe(false)
  })

  test('B11 targeted "yes <id>" is a permission verdict', async () => {
    const s = setup()
    await s.run(makeActivity({ text: 'yes abcde' }))
    expect(s.resolved).toEqual([['abcde', 'allow']])
    expect(s.events.length).toBe(0)
  })
})

describe('targeted reply', () => {
  for (const [name, on] of [['B5', true], ['B12 (switch off)', false]] as const) {
    test(`${name} reply with targeted_message_id sends one targeted POST`, async () => {
      const s = setup({ targetedMessages: on })
      await s.run(makeActivity())
      const id = await s.replier.sendReply(CHAN, 'answer', undefined, { targetedMessageId: 'msg-1' })
      expect(id).toBe('sent-1')
      expect(s.publicSends.length).toBe(0)
      expect(s.posts.length).toBe(1)
      const req = s.posts[0]
      expect(req.method).toBe('POST')
      expect(req.url).toBe(
        'https://smba.example.invalid/teams/v3/conversations/' + encodeURIComponent(CHAN) + '/activities?isTargetedActivity=true',
      )
      const body = JSON.parse(req.body)
      expect(body.recipient).toEqual({ id: '29:user-1', name: 'Adnan', aadObjectId: ALLOWED, isTargeted: true })
      expect(body.entities).toEqual([{ type: 'targetedMessageInfo', messageId: 'msg-1' }])
      expect(body.textFormat).toBe('markdown')
      expect(body.text).toBe('answer')
    })
  }

  test('B5 card goes into attachments', async () => {
    const s = setup()
    await s.run(makeActivity())
    const card = { type: 'AdaptiveCard', version: '1.5', body: [{ type: 'TextBlock', text: 'x' }] }
    await s.replier.sendReply(CHAN, 'answer', card, { targetedMessageId: 'msg-1' })
    const body = JSON.parse(s.posts[0].body)
    expect(body.attachments.length).toBe(1)
    expect(body.attachments[0].content).toEqual(card)
  })

  test('B6 unknown, other-conversation and expired ids throw, nothing sent', async () => {
    const s = setup()
    await s.run(makeActivity())
    await s.run(makeActivity({ id: 'msg-b', conversation: { id: CHAN_B, tenantId: TENANT_ID, conversationType: 'channel' } as any }))
    const msg = 'unknown or expired targeted_message_id for this conversation'
    await expect(s.replier.sendReply(CHAN, 'x', undefined, { targetedMessageId: 'nope' })).rejects.toThrow(msg)
    await expect(s.replier.sendReply(CHAN, 'x', undefined, { targetedMessageId: 'msg-b' })).rejects.toThrow(msg)
    s.setNow(s.getNow() + TARGETED_TTL_MS + 60_000)
    await expect(s.replier.sendReply(CHAN, 'x', undefined, { targetedMessageId: 'msg-1' })).rejects.toThrow(msg)
    expect(s.posts.length).toBe(0)
    expect(s.publicSends.length).toBe(0)
  })

  test('B7 reply after a targeted inbound is refused without targeted_message_id; public: true posts', async () => {
    const s = setup()
    await s.run(makeActivity())
    await expect(s.replier.sendReply(CHAN, 'x')).rejects.toThrow(LATEST_PRIVATE_MESSAGE)
    expect(s.publicSends.length).toBe(0)
    expect(s.posts.length).toBe(0)
    await s.replier.sendReply(CHAN, 'x', undefined, { public: true })
    expect(s.publicSends.length).toBe(1)
    expect(s.posts.length).toBe(0)
  })

  test('B8 a normal @mention after the targeted one re-allows public replies', async () => {
    const s = setup()
    await s.run(makeActivity())
    await s.run(mentionActivity({ id: 'msg-2' }))
    await s.replier.sendReply(CHAN, 'x')
    expect(s.publicSends.length).toBe(1)
  })

  test('B9 sender removed from the allowlist after the inbound', async () => {
    const s = setup()
    await s.run(makeActivity())
    await s.run(mentionActivity({ id: 'msg-2', from: { id: '29:user-3', name: 'Z', aadObjectId: OTHER } as any }))
    s.allowlist.removeEntry(ALLOWED)
    await expect(s.replier.sendReply(CHAN, 'x', undefined, { targetedMessageId: 'msg-1' })).rejects.toBeInstanceOf(
      ConversationNotAllowedError,
    )
    expect(s.posts.length).toBe(0)
    expect(s.publicSends.length).toBe(0)
  })

  test('B10 HTTP 403 fails the call, no fallback', async () => {
    const s = setup({ status: 403 })
    await s.run(makeActivity())
    await expect(s.replier.sendReply(CHAN, 'x', undefined, { targetedMessageId: 'msg-1' })).rejects.toThrow('HTTP 403')
    expect(s.posts.length).toBe(1)
    expect(s.publicSends.length).toBe(0)
  })

  test('B13 send_file is refused right after a targeted inbound unless public', async () => {
    const s = setup()
    await s.run(makeActivity())
    const sendFile = createFileSender({
      config: s.config, adapter: s.adapter, allowlist: s.allowlist, refs: s.refs,
      outbox: { isPathAllowed: () => true, mint: async (x: any) => ({ token: 'tok-12345678', filename: x.filename, sizeBytes: x.content.length, mime: x.mime }) } as any,
      downloadBaseUrl: 'https://files.example.invalid/files',
      targetedInbound: s.targetedInbound,
      now: s.getNow,
    })
    const input = { conversationId: CHAN, content: Buffer.from('hi'), filename: 'a.txt' }
    await expect(sendFile(input)).rejects.toThrow('the latest message in this conversation was private (targeted)')
    expect(s.publicSends.length).toBe(0)
    await sendFile({ ...input, public: true })
    expect(s.publicSends.length).toBe(1)
  })

  test('B14 edit_message on a targeted reply id is refused', async () => {
    const s = setup()
    await s.run(makeActivity())
    const id = await s.replier.sendReply(CHAN, 'x', undefined, { targetedMessageId: 'msg-1' })
    await expect(s.replier.editReply(CHAN, id!, 'y')).rejects.toThrow('editing a targeted message is not supported')
    expect(s.edits.length).toBe(0)
  })

  test('B15 targeted_message_id together with public throws', async () => {
    const s = setup()
    await s.run(makeActivity())
    await expect(
      s.replier.sendReply(CHAN, 'x', undefined, { targetedMessageId: 'msg-1', public: true }),
    ).rejects.toThrow('targeted_message_id and public cannot be combined')
    expect(s.posts.length).toBe(0)
    expect(s.publicSends.length).toBe(0)
  })

  test('B19 targeted_message_id together with user_cards throws: the private post would drop the views', async () => {
    const s = setup()
    await s.run(makeActivity())
    const card = { type: 'AdaptiveCard', version: '1.5', body: [{ type: 'TextBlock', text: 'base' }] }
    await expect(
      s.replier.sendReply(CHAN, 'x', card, { targetedMessageId: 'msg-1', userCards: [{ aad_object_id: ALLOWED, card }] }),
    ).rejects.toThrow('targeted_message_id and user_cards cannot be combined')
    expect(s.posts.length).toBe(0)
    expect(s.publicSends.length).toBe(0)
  })
})

describe('targetedInbound memory', () => {
  const ref = { user: { id: 'u' } } as any
  test('201st entry drops the oldest', () => {
    const t = createTargetedInbound()
    for (let i = 0; i <= MAX_ENTRIES_PER_CONVERSATION; i++) t.remember('c', `m${i}`, ref, 'aad', 1000)
    expect(t.get('c', 'm0', 1000)).toBeUndefined()
    expect(t.get('c', 'm1', 1000)).toBeDefined()
    expect(t.get('c', `m${MAX_ENTRIES_PER_CONVERSATION}`, 1000)).toBeDefined()
  })
  test('1001st conversation drops the oldest', () => {
    const t = createTargetedInbound()
    for (let i = 0; i <= MAX_CONVERSATIONS; i++) t.remember(`c${i}`, 'm', ref, 'aad', 1000)
    expect(t.get('c0', 'm', 1000)).toBeUndefined()
    expect(t.get('c1', 'm', 1000)).toBeDefined()
  })
  test('sent ids are bounded', () => {
    const t = createTargetedInbound()
    for (let i = 0; i <= MAX_SENT_IDS; i++) t.rememberSent(`s${i}`)
    expect(t.isSentTargeted('s0')).toBe(false)
    expect(t.isSentTargeted('s1')).toBe(true)
  })
  test('24 h expiry', () => {
    const t = createTargetedInbound()
    t.remember('c', 'm', ref, 'aad', 1000)
    t.noteDelivered('c', true, 1000)
    expect(t.get('c', 'm', 1000 + TARGETED_TTL_MS)).toBeDefined()
    expect(t.get('c', 'm', 1000 + TARGETED_TTL_MS + 1)).toBeUndefined()
    expect(t.latestIsTargeted('c', 1000 + TARGETED_TTL_MS)).toBe(true)
    expect(t.latestIsTargeted('c', 1000 + TARGETED_TTL_MS + 1)).toBe(false)
  })
  test('latest flag follows the most recent delivered message', () => {
    const t = createTargetedInbound()
    t.noteDelivered('c', true, 1)
    t.noteDelivered('c', false, 2)
    expect(t.latestIsTargeted('c', 3)).toBe(false)
    expect(t.get('other', 'm', 3)).toBeUndefined()
  })
})

describe('review follow-ups', () => {
  test('B16 other conversations never evict a fresh private marker; old or public markers make room', () => {
    const t = createTargetedInbound()
    t.noteDelivered('A', true, 1000)
    for (let i = 0; i < MAX_CONVERSATIONS * 2; i++) t.noteDelivered(`c${i}`, false, 1500)
    expect(t.latestIsTargeted('A', 2000)).toBe(true)
  })

  test('B18 the bound is hard: a fresh private marker forced out fails closed until it expires', () => {
    const t = createTargetedInbound()
    for (let i = 0; i < MAX_CONVERSATIONS; i++) t.noteDelivered(`p${i}`, true, 1000 + i)
    expect(t.latestIsTargeted('never-seen', 5000)).toBe(false)
    t.noteDelivered('new', false, 5000) // no room: p0 (the oldest private marker) has to go
    expect(t.latestIsTargeted('p0', 5000)).toBe(true) // untracked now, so private
    expect(t.latestIsTargeted('never-seen', 5000)).toBe(true)
    expect(t.latestIsTargeted('new', 5000)).toBe(false) // tracked and public
    expect(t.latestIsTargeted('never-seen', 1000 + TARGETED_TTL_MS)).toBe(true)
    expect(t.latestIsTargeted('never-seen', 1000 + TARGETED_TTL_MS + 1)).toBe(false)
  })

  test('B17 targeted_message_id: absent is undefined, a non-empty string passes, anything else is refused', () => {
    expect(parseTargetedMessageId(undefined)).toBeUndefined()
    expect(parseTargetedMessageId(null)).toBeUndefined()
    expect(parseTargetedMessageId('1716060180000')).toBe('1716060180000')
    for (const bad of [1716060180000, '', {}, true]) {
      expect(() => parseTargetedMessageId(bad)).toThrow('targeted_message_id must be')
    }
  })
})
