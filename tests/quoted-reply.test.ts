import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { channelEventMeta, inboundQuote, makeTurnHandler, stripQuotePlaceholders } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createReplySender, parseQuoteArg } from '../src/teams/reply'
import { createSeenMessages } from '../src/teams/seenMessages'

const ID = '00000000-0000-4000-8000-000000000001'
const TENANT = '00000000-0000-4000-8000-000000000002'
const ent = (q: object) => [{ type: 'quotedReply', quotedReply: q }]

describe('inboundQuote', () => {
  test('valid entity', () => {
    expect(inboundQuote(ent({ messageId: '1716', senderName: 'Anna Smith', preview: 'Deadline is Friday' }))).toEqual({
      messageId: '1716', senderName: 'Anna Smith', preview: 'Deadline is Friday', deleted: false,
    })
  })
  test('cleans and cuts preview', () => {
    const q = inboundQuote(ent({ messageId: '1', preview: 'a\nb' + 'x'.repeat(400) }))!
    expect(q.preview!.length).toBe(300)
    expect(q.preview!.startsWith('a b')).toBe(true)
  })
  test('deleted', () => {
    expect(inboundQuote(ent({ messageId: '1', isReplyDeleted: true }))!.deleted).toBe(true)
  })
  test('invalid / missing', () => {
    expect(inboundQuote(ent({ messageId: 'a"b' }))).toBeUndefined()
    expect(inboundQuote(ent({ messageId: '' }))).toBeUndefined()
    expect(inboundQuote([{ type: 'mention' }])).toBeUndefined()
    expect(inboundQuote({})).toBeUndefined()
  })
  test('non-whitespace control characters in the preview and the name are replaced', () => {
    const q = inboundQuote([{ type: 'quotedReply', quotedReply: { messageId: '1', senderName: 'K\u0007A', preview: 'a\u0007b\u009bc' } }])
    expect(q?.preview).toBe('a b c')
    expect(q?.senderName).toBe('K A')
  })
})

describe('stripQuotePlaceholders', () => {
  test('strips', () => {
    expect(stripQuotePlaceholders('<quoted messageId="1716"/> hello')).toBe('hello')
    expect(stripQuotePlaceholders('<quoted messageId="1"/> a <quoted messageId="2"/>')).toBe('a')
    expect(stripQuotePlaceholders('plain')).toBe('plain')
  })
})

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

async function run(text: string, entities?: unknown, seenMessages?: ReturnType<typeof createSeenMessages>) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-quote-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(ID, 'seed')
  const events: any[] = []
  const handler = makeTurnHandler({
    config: { stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'a', appPassword: 'p', appType: 'SingleTenant', tenantId: TENANT, bindHost: '127.0.0.1', port: 0 } as any,
    adapter: {} as any,
    allowlist,
    refs: createConversationRefStore(),
    onEvent: e => events.push(e),
    seenMessages,
  })
  const act = {
    type: 'message', text, id: 'msg-1', timestamp: new Date(),
    from: { id: 'u', name: 'T', aadObjectId: ID },
    conversation: { id: 'conv-abc', tenantId: TENANT },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams', serviceUrl: 'https://smba.trafficmanager.net/uk/',
    ...(entities ? { entities } : {}),
  } as unknown as Activity
  await handler(new TurnContext(new NoopAdapter() as any, act))
  return events[0]
}

describe('adapter quotes', () => {
  const q = { messageId: '1716', senderName: 'Anna Smith', preview: 'Deadline is Friday' }
  test('quote', async () => {
    const e = await run('<quoted messageId="1716"/> yes, go ahead', ent(q))
    expect(e.text).toBe('[Quote] Anna Smith: Deadline is Friday\nyes, go ahead')
    expect(e.quotedMessageId).toBe('1716')
  })
  test('deleted', async () => {
    const e = await run('<quoted messageId="1716"/> ok', ent({ ...q, isReplyDeleted: true }))
    expect(e.text).toBe('[Quote] Anna Smith: (deleted message)\nok')
  })
  test('no quote', async () => {
    const e = await run('hello')
    expect(e.text).toBe('hello')
    expect('quotedMessageId' in e).toBe(false)
  })
})

function setup(opts: { store?: boolean; fakeId?: string } = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-quote-r-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(ID)
  const refs = createConversationRefStore()
  refs.put('conv-abc', { conversation: { id: 'conv-abc' } } as any, ID)
  const sent: any[] = []
  const adapter = {
    async continueConversationAsync(_a: string, _r: any, logic: any) {
      await logic({ async sendActivity(act: any) { sent.push(act); return opts.fakeId ? { id: opts.fakeId } : undefined } })
    },
  } as any
  const config = { appId: 'a' } as any
  const seen = createSeenMessages()
  seen.remember('conv-abc', '1716')
  seen.remember('conv-other', 'foreign-1')
  const seenMessages = opts.store === false ? undefined : seen
  return { sent, seen, sender: createReplySender({ config, adapter, allowlist, refs, seenMessages }) }
}

describe('sendReply quoting', () => {
  test('with quote', async () => {
    const { sent, sender } = setup()
    await sender.sendReply('conv-abc', 'ok', undefined, { quoteMessageId: '1716' })
    expect(sent[0].entities).toEqual(ent({ messageId: '1716' }))
    expect(sent[0].text).toBe('<quoted messageId="1716"/> ok')
  })
  test('invalid id', async () => {
    const { sent, sender } = setup()
    await expect(sender.sendReply('conv-abc', 'ok', undefined, { quoteMessageId: 'x"y' })).rejects.toThrow('invalid quote message id')
    await expect(sender.sendReply('conv-abc', 'ok', undefined, { quoteMessageId: '' })).rejects.toThrow('invalid quote message id')
    expect(sent.length).toBe(0)
  })
  test('without', async () => {
    const { sent, sender } = setup()
    await sender.sendReply('conv-abc', 'ok')
    expect('entities' in sent[0]).toBe(false)
    expect(sent[0].text).toBe('ok')
  })
})

describe('parseQuoteArg', () => {
  test('omitted or null is no quote, a string goes on (sendReply rejects an invalid one), other types throw', () => {
    expect(parseQuoteArg(undefined)).toBeUndefined()
    expect(parseQuoteArg(null)).toBeUndefined()
    expect(parseQuoteArg('1716')).toBe('1716')
    expect(parseQuoteArg('')).toBe('')
    expect(() => parseQuoteArg(42)).toThrow('quote_message_id must be a string')
  })
})

describe('createSeenMessages', () => {
  test('remember and has, per conversation', () => {
    const s = createSeenMessages()
    s.remember('c1', 'm1')
    expect(s.has('c1', 'm1')).toBe(true)
    expect(s.has('c2', 'm1')).toBe(false)
  })
  test('per-conversation bound drops the oldest; re-remembering makes an id newest', () => {
    const s = createSeenMessages(3)
    for (const id of ['a', 'b', 'c', 'd']) s.remember('c', id)
    expect(s.has('c', 'a')).toBe(false)
    expect(s.has('c', 'b')).toBe(true)
    const t = createSeenMessages(3)
    for (const id of ['a', 'b', 'c']) t.remember('c', id)
    t.remember('c', 'a')
    t.remember('c', 'd')
    expect(t.has('c', 'a')).toBe(true)
    expect(t.has('c', 'b')).toBe(false)
  })
  test('conversation bound drops the least recently used; has does not refresh', () => {
    const s = createSeenMessages(500, 2)
    s.remember('x', '1'); s.remember('y', '1'); s.remember('z', '1')
    expect(s.has('x', '1')).toBe(false)
    expect(s.has('y', '1')).toBe(true)
    const t = createSeenMessages(500, 2)
    t.remember('x', '1'); t.remember('y', '1')
    t.has('x', '1')
    t.remember('z', '1')
    expect(t.has('x', '1')).toBe(false)
    expect(t.has('y', '1')).toBe(true)
    const u = createSeenMessages(500, 2)
    u.remember('x', '1'); u.remember('y', '1'); u.remember('x', '2')
    u.remember('z', '1')
    expect(u.has('x', '1')).toBe(true)
    expect(u.has('y', '1')).toBe(false)
  })
})

describe('outbound quote restriction', () => {
  const msg = 'quote_message_id is not a message seen in this conversation'
  test('id seen only in another conversation is refused; no store refuses; seen id is sent', async () => {
    const a = setup()
    await expect(a.sender.sendReply('conv-abc', 'ok', undefined, { quoteMessageId: 'foreign-1' })).rejects.toThrow(msg)
    expect(a.sent.length).toBe(0)
    const b = setup({ store: false })
    await expect(b.sender.sendReply('conv-abc', 'ok', undefined, { quoteMessageId: '1716' })).rejects.toThrow(msg)
    expect(b.sent.length).toBe(0)
    await a.sender.sendReply('conv-abc', 'ok', undefined, { quoteMessageId: '1716' })
    expect(a.sent.length).toBe(1)
  })
  test('the id of a sent message can be quoted afterwards', async () => {
    const { sender, sent } = setup({ fakeId: 'bot-1' })
    await sender.sendReply('conv-abc', 'first')
    await sender.sendReply('conv-abc', 'second', undefined, { quoteMessageId: 'bot-1' })
    expect(sent[1].entities).toEqual(ent({ messageId: 'bot-1' }))
  })
})

describe('adapter remembers seen ids', () => {
  test('inbound id is remembered, a quoted id is not', async () => {
    const s = createSeenMessages()
    await run('hello', undefined, s)
    expect(s.has('conv-abc', 'msg-1')).toBe(true)
    const t = createSeenMessages()
    await run('hi', ent({ messageId: 'q-9' }), t)
    expect(t.has('conv-abc', 'msg-1')).toBe(true)
    expect(t.has('conv-abc', 'q-9')).toBe(false)
  })
})

describe('quote wiring', () => {
  test('channelEventMeta carries quoted_message_id', () => {
    const event = { text: 't', conversationId: 'c', aadObjectId: ID, tenantId: TENANT, fromName: 'n' }
    expect(channelEventMeta({ ...event, quotedMessageId: '1:abc' }).quoted_message_id).toBe('1:abc')
    expect('quoted_message_id' in channelEventMeta(event)).toBe(false)
  })
  test('server.ts parses quote_message_id and shares one seenMessages between the replier and the listener', () => {
    const server = readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8')
    const start = server.indexOf("      case 'reply': {")
    const reply = server.slice(start, server.indexOf("\n      case '", start + 1))
    expect(reply).toContain('const quoteMessageId = parseQuoteArg(args.quote_message_id)')
    expect(reply).toContain('...(quoteMessageId !== undefined ? { quoteMessageId } : {}),')
    expect(server).toContain('const seenMessages = createSeenMessages()')
    expect(server).toMatch(/const replier = createReplySender\(\{[^}]*\bseenMessages\b[^}]*\}\)/)
    const lstart = server.indexOf('const listener = startListener({')
    expect(server.slice(lstart, server.indexOf('\n})', lstart))).toMatch(/\n  seenMessages,\n/)
  })
})
