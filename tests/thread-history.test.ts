import { describe, expect, test } from 'bun:test'
import { BotAdapter, TeamsInfo, TurnContext, type Activity } from 'botbuilder'
import {
  createTeamGroupStore,
  fetchThreadHistory,
  parseChannelThread,
} from '../src/teams/threadHistory'
import { baseConversationId, makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

const J = (b: unknown, status = 200) =>
  new Response(JSON.stringify(b), { status, headers: { 'content-type': 'application/json' } })

type Call = { url: string; init?: RequestInit }
function fakeFetch(handler: (url: string, init?: RequestInit) => Response | Promise<Response>) {
  const calls: Call[] = []
  const fn = async (url: string, init?: RequestInit) => {
    calls.push({ url, init })
    return handler(url, init)
  }
  return { fn, calls }
}

const msg = (id: string, t: string, extra: Record<string, unknown> = {}) => ({
  id,
  replyToId: null,
  createdDateTime: t,
  messageType: 'message',
  deletedDateTime: null,
  from: { user: { id: 'AAA-1', displayName: 'Ann' } },
  body: { contentType: 'text', content: `text ${id}` },
  ...extra,
})
const G = 'https://graph.microsoft.com/v1.0/'
const P = { teamId: 'team1', channelId: 'c1', rootMessageId: 'r1' }
const tok = async () => 'TOKEN'

describe('fetchThreadHistory', () => {
  test('H3 order and URLs', async () => {
    const f = fakeFetch(url =>
      url.includes('/replies')
        ? J({ value: [msg('r4', '2026-01-01T00:00:04Z'), msg('r2', '2026-01-01T00:00:02Z'), msg('r3', '2026-01-01T00:00:03Z')] })
        : J(msg('r1', '2026-01-01T00:00:00Z')),
    )
    const out = await fetchThreadHistory(
      { teamId: 'g 1', channelId: '19:a b@thread.tacv2', rootMessageId: '123' },
      { getToken: tok, fetch: f.fn },
    )
    expect(out.map(m => m.id)).toEqual(['r1', 'r2', 'r3', 'r4'])
    const base = `${G}teams/g%201/channels/19%3Aa%20b%40thread.tacv2/messages/123`
    expect(f.calls.map(c => c.url)).toEqual([base, `${base}/replies?$top=30`])
    expect((f.calls[0]!.init!.headers as Record<string, string>).Authorization).toBe('Bearer TOKEN')
    expect(f.calls[1]!.init!.redirect).toBe('error')
  })
  test('H4 paging and limit', async () => {
    const next = `${G}next-page`
    const f = fakeFetch(url => {
      if (url === next) return J({ value: [msg('r3', '3'), msg('r2', '2')], '@odata.nextLink': `${G}more` })
      if (url.includes('/replies')) return J({ value: [msg('r5', '5'), msg('r4', '4')], '@odata.nextLink': next })
      return J(msg('r1', '0'))
    })
    const out = await fetchThreadHistory({ ...P, limit: 3 }, { getToken: tok, fetch: f.fn })
    expect(out.map(m => m.id)).toEqual(['r1', 'r3', 'r4', 'r5'])
    expect(f.calls.length).toBe(3)
    expect(f.calls[1]!.url).toContain('$top=3')
  })
  test('$top never exceeds 50', async () => {
    const f = fakeFetch(url => (url.includes('/replies') ? J({ value: [] }) : J(msg('r1', '0'))))
    await fetchThreadHistory({ ...P, limit: 100 }, { getToken: tok, fetch: f.fn })
    expect(f.calls[1]!.url).toEndWith('/replies?$top=50')
  })
  test('H5 foreign next link', async () => {
    const f = fakeFetch(url =>
      url.includes('/replies') ? J({ value: [msg('a', '1')], '@odata.nextLink': 'https://evil.example/x' }) : J(msg('r1', '0')),
    )
    await expect(fetchThreadHistory(P, { getToken: tok, fetch: f.fn })).rejects.toThrow('unexpected Graph next link')
    expect(f.calls.some(c => c.url.includes('evil'))).toBe(false)
  })
  test('a next link that leaves /v1.0/ through dot segments is refused', async () => {
    for (const link of [`${G}../beta/teams/x`, `${G}%2e%2e/beta/teams/x`]) {
      const f = fakeFetch(url =>
        url.includes('/replies') ? J({ value: [msg('a', '1')], '@odata.nextLink': link }) : J(msg('r1', '0')),
      )
      await expect(fetchThreadHistory(P, { getToken: tok, fetch: f.fn })).rejects.toThrow('unexpected Graph next link')
      expect(f.calls.length).toBe(2)
    }
  })
  test('H6 errors', async () => {
    const run = (status: number, onlyRoot = true) =>
      fetchThreadHistory(P, { getToken: tok, fetch: fakeFetch(() => J({}, status)).fn }).catch(e => (e as Error).message)
    expect(await run(403)).toContain('ChannelMessage.Read.Group grant')
    expect(await run(404)).toBe('thread not found')
    expect(await run(500)).toBe('Graph request failed (HTTP 500)')
    const repliesGone = fakeFetch(url => (url.includes('/replies') ? J({}, 404) : J(msg('r1', '0'))))
    expect(await fetchThreadHistory(P, { getToken: tok, fetch: repliesGone.fn }).catch(e => (e as Error).message)).toBe(
      'Graph request failed (HTTP 404)',
    )
  })
  test('network errors name only a plain error code', async () => {
    const failWith = (code: string) =>
      fetchThreadHistory(P, {
        getToken: tok,
        fetch: fakeFetch(() => {
          throw Object.assign(new Error('boom https://x/?secret=1'), { code })
        }).fn,
      }).catch(e => (e as Error).message)
    expect(await failWith('ECONNRESET')).toBe('Graph request failed (ECONNRESET)')
    expect(await failWith('https://x/?secret=1')).toBe('Graph request failed (network error)')
  })
  test('H7 system and deleted skipped', async () => {
    const f = fakeFetch(url =>
      url.includes('/replies')
        ? J({ value: [msg('s', '3', { messageType: 'systemEventMessage' }), msg('d', '2', { deletedDateTime: '2026' }), msg('ok', '1')] })
        : J(msg('r1', '0')),
    )
    const out = await fetchThreadHistory(P, { getToken: tok, fetch: f.fn })
    expect(out.map(m => m.id)).toEqual(['r1', 'ok'])
    const sysRoot = fakeFetch(url =>
      url.includes('/replies') ? J({ value: [msg('ok', '1')] }) : J(msg('r1', '0', { messageType: 'systemEventMessage' })),
    )
    const out2 = await fetchThreadHistory(P, { getToken: tok, fetch: sysRoot.fn })
    expect(out2.map(m => m.id)).toEqual(['ok'])
  })
  test('H8 html and truncation', async () => {
    const html = '<p>a &amp; b &lt;x&gt; &quot;q&quot; it&#39;s</p><div>x&nbsp;y<br>z</div><b>bold</b>\n\n\n\nend'
    const f = fakeFetch(url =>
      url.includes('/replies')
        ? J({ value: [msg('l', '2', { body: { contentType: 'text', content: '😀'.repeat(5000) } })] })
        : J(msg('r1', '0', { body: { contentType: 'html', content: html } })),
    )
    const out = await fetchThreadHistory(P, { getToken: tok, fetch: f.fn })
    expect(out[0]!.text).toBe('a & b <x> "q" it\'s\nx y\nz\nbold\n\nend')
    expect(out[1]!.text).toBe('😀'.repeat(3994) + ' [...]')
    expect(Array.from(out[1]!.text).length).toBe(4000)
  })
  test('H9 allowlisted', async () => {
    const f = fakeFetch(url =>
      url.includes('/replies')
        ? J({ value: [msg('app', '2', { from: { application: { displayName: 'Bot' } } }), msg('bob', '1', { from: { user: { id: 'BBB', displayName: 'Bob' } } })] })
        : J(msg('r1', '0')),
    )
    const out = await fetchThreadHistory(P, { getToken: tok, fetch: f.fn, isAllowed: id => id === 'aaa-1' })
    expect(out[0]).toMatchObject({ from: 'Ann', fromAadObjectId: 'aaa-1', allowlisted: true })
    expect(out[1]).toMatchObject({ id: 'bob', allowlisted: false, fromAadObjectId: 'bbb' })
    expect(out[2]).toMatchObject({ from: 'Bot', fromAadObjectId: null, allowlisted: false })
  })
  test('H10 bad limits', async () => {
    const f = fakeFetch(() => J({}))
    for (const limit of [0, 101, 1.5]) {
      await expect(fetchThreadHistory({ ...P, limit }, { getToken: tok, fetch: f.fn })).rejects.toThrow()
    }
    expect(f.calls.length).toBe(0)
  })
})

test('H14 parseChannelThread', () => {
  const id = '19:abc@thread.tacv2;messageid=1700000000000'
  expect(parseChannelThread(id)).toEqual({ channelId: '19:abc@thread.tacv2', rootMessageId: '1700000000000' })
  expect(parseChannelThread(id).channelId).toBe(baseConversationId(id))
  expect(() => parseChannelThread('19:abc@thread.tacv2')).toThrow('not a channel thread (no ;messageid=)')
  expect(() => parseChannelThread('19:abc@thread.tacv2;messageid=')).toThrow('not a channel thread (no ;messageid=)')
  expect(() => parseChannelThread('19:chat@thread.v2;messageid=1')).toThrow('only available for channel threads')
  expect(() => parseChannelThread('a:1personal')).toThrow('only available for channel threads')
  expect(() => parseChannelThread('19:x;messageid=1@thread.tacv2')).toThrow('only available for channel threads')
  expect(() => parseChannelThread('19:x@thread.tacv2-junk;messageid=1')).toThrow('only available for channel threads')
})

test('H11 team group store bound', () => {
  const s = createTeamGroupStore()
  for (let i = 0; i < 1001; i++) s.set(`c${i}`, `g${i}`)
  expect(s.get('c0')).toBeUndefined()
  expect(s.get('c1')).toBe('g1')
  expect(s.get('c1000')).toBe('g1000')
})

describe('H12 adapter', () => {
  const ALLOWED = '00000000-0000-4000-8000-000000000001'
  const TENANT = '00000000-0000-4000-8000-000000000002'
  class Noop extends BotAdapter {
    async sendActivities(): Promise<any> { return [] }
    async updateActivity(): Promise<void> {}
    async deleteActivity(): Promise<void> {}
    async continueConversation(): Promise<void> {}
  }
  function setup(aad: string, withStore = true, channelData: unknown = { team: { aadGroupId: 'group-1' } }) {
    let events = 0
    const dir = mkdtempSync(join(tmpdir(), 'th-'))
    const allowlist = createAllowlist(join(dir, 'allow.json'))
    allowlist.addEntry(ALLOWED, 'x')
    const store = createTeamGroupStore()
    const turn = makeTurnHandler({
      config: { stateDir: dir, allowlistFile: join(dir, 'allow.json'), appId: 'a', appPassword: 'p', appType: 'SingleTenant', tenantId: TENANT, bindHost: '127.0.0.1', port: 0 } as any,
      adapter: {} as any,
      allowlist,
      refs: createConversationRefStore(),
      onEvent: () => {
        events++
      },
      ...(withStore ? { teamGroups: store } : {}),
    })
    const act = {
      type: 'message',
      text: '<at>Bot</at> hi',
      id: 'm1',
      from: { id: 'u', name: 'U', aadObjectId: aad },
      conversation: { id: '19:abc@thread.tacv2;messageid=123', conversationType: 'channel', tenantId: TENANT },
      recipient: { id: 'bot-1', name: 'Bot' },
      channelId: 'msteams',
      serviceUrl: 'https://smba.trafficmanager.net/uk/',
      entities: [{ type: 'mention', mentioned: { id: 'bot-1' }, text: '<at>Bot</at>' }],
      channelData,
    } as unknown as Activity
    return { store, events: () => events, run: () => turn(new TurnContext(new Noop(), act)) }
  }
  test('stores for allowlisted sender', async () => {
    const s = setup(ALLOWED)
    await s.run()
    expect(s.store.get('19:abc@thread.tacv2')).toBe('group-1')
    expect(s.events()).toBe(1)
  })
  test('nothing for non-allowlisted', async () => {
    const s = setup('aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee')
    await s.run()
    expect(s.store.get('19:abc@thread.tacv2')).toBeUndefined()
  })
  test('no store dep is fine', async () => {
    const s = setup(ALLOWED, false)
    await s.run()
    expect(s.events()).toBe(1)
  })
})

describe('H13 team details lookup (channelData.team.id only)', () => {
  const ALLOWED = '00000000-0000-4000-8000-000000000001'
  const TENANT = '00000000-0000-4000-8000-000000000002'
  class Noop extends BotAdapter {
    async sendActivities(): Promise<any> { return [] }
    async updateActivity(): Promise<void> {}
    async deleteActivity(): Promise<void> {}
    async continueConversation(): Promise<void> {}
  }
  const original = TeamsInfo.getTeamDetails
  function setup(lookup: () => Promise<unknown>) {
    let lookups = 0
    let events = 0
    TeamsInfo.getTeamDetails = (async () => {
      lookups++
      return lookup()
    }) as any
    const dir = mkdtempSync(join(tmpdir(), 'th-'))
    const allowlist = createAllowlist(join(dir, 'allow.json'))
    allowlist.addEntry(ALLOWED, 'x')
    const store = createTeamGroupStore()
    const turn = makeTurnHandler({
      config: { stateDir: dir, allowlistFile: join(dir, 'allow.json'), appId: 'a', appPassword: 'p', appType: 'SingleTenant', tenantId: TENANT, bindHost: '127.0.0.1', port: 0 } as any,
      adapter: {} as any,
      allowlist,
      refs: createConversationRefStore(),
      onEvent: () => {
        events++
      },
      teamGroups: store,
    })
    const act = (id: string, channelData: unknown) =>
      ({
        type: 'message',
        text: '<at>Bot</at> hi',
        id,
        from: { id: 'u', name: 'U', aadObjectId: ALLOWED },
        conversation: { id: '19:abc@thread.tacv2;messageid=123', conversationType: 'channel', tenantId: TENANT },
        recipient: { id: 'bot-1', name: 'Bot' },
        channelId: 'msteams',
        serviceUrl: 'https://smba.trafficmanager.net/uk/',
        entities: [{ type: 'mention', mentioned: { id: 'bot-1' }, text: '<at>Bot</at>' }],
        channelData,
      }) as unknown as Activity
    const run = (id = 'm1', channelData: unknown = { team: { id: '19:team@thread.tacv2' } }) =>
      Promise.race([
        turn(new TurnContext(new Noop(), act(id, channelData))).then(() => 'done'),
        new Promise(r => setTimeout(() => r('blocked'), 1000)),
      ])
    return { store, run, events: () => events, lookups: () => lookups }
  }
  const settle = () => new Promise(r => setTimeout(r, 10))

  test('a lookup that never answers does not delay delivery', async () => {
    try {
      const s = setup(() => new Promise(() => {}))
      expect(await s.run()).toBe('done')
      expect(s.events()).toBe(1)
      expect(s.store.get('19:abc@thread.tacv2')).toBeUndefined()
    } finally {
      TeamsInfo.getTeamDetails = original
    }
  })
  test('the looked-up id is stored, and a known channel is not looked up again', async () => {
    try {
      const s = setup(async () => ({ aadGroupId: 'group-2' }))
      expect(await s.run('m1')).toBe('done')
      await settle()
      expect(s.store.get('19:abc@thread.tacv2')).toBe('group-2')
      expect(await s.run('m2')).toBe('done')
      await settle()
      expect(s.lookups()).toBe(1)
      expect(s.events()).toBe(2)
    } finally {
      TeamsInfo.getTeamDetails = original
    }
  })
  test('a late lookup answer does not overwrite an id that arrived meanwhile', async () => {
    try {
      let answer: (v: unknown) => void = () => {}
      const s = setup(() => new Promise(r => (answer = r)))
      expect(await s.run('m1')).toBe('done')
      expect(await s.run('m2', { team: { id: '19:team@thread.tacv2', aadGroupId: 'group-new' } })).toBe('done')
      answer({ aadGroupId: 'group-old' })
      await settle()
      expect(s.store.get('19:abc@thread.tacv2')).toBe('group-new')
    } finally {
      TeamsInfo.getTeamDetails = original
    }
  })
  test('a failing lookup is logged and delivery goes on', async () => {
    try {
      const s = setup(async () => {
        throw new Error('boom')
      })
      expect(await s.run()).toBe('done')
      await settle()
      expect(s.events()).toBe(1)
      expect(s.store.get('19:abc@thread.tacv2')).toBeUndefined()
    } finally {
      TeamsInfo.getTeamDetails = original
    }
  })
})
