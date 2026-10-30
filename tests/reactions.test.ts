import { afterAll, describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { ConversationNotAllowedError, createReplySender } from '../src/teams/reply'
import { channelEventMeta, makeTurnHandler, reactionEventText } from '../src/teams/adapter'
// Every temp dir of this file, removed after the last test (they used to stay in the system temp dir).
const tmpDirs: string[] = []
function tempDir(prefix: string): string {
  const dir = mkdtempSync(join(tmpdir(), prefix))
  tmpDirs.push(dir)
  return dir
}
afterAll(() => {
  for (const d of tmpDirs) rmSync(d, { recursive: true, force: true })
})

const ID = '00000000-0000-4000-8000-000000000001'
const UNKNOWN = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
const TENANT = '00000000-0000-4000-8000-000000000002'
const KEY = Symbol('connector')
const CONV = '19:abc@thread.tacv2;messageid=1'
const config: any = { appId: 'app', appPassword: 'pw', appType: 'SingleTenant', tenantId: TENANT, bindHost: '127.0.0.1', port: 0 }

function setup(status = 200, allowed = true) {
  const dir = tempDir('cct-react-')
  const allowlist = createAllowlist(join(dir, 'a.json'))
  allowlist.addEntry(ID)
  const refs = createConversationRefStore()
  refs.put(
    CONV,
    { serviceUrl: 'https://smba.trafficmanager.net/emea/', conversation: { id: CONV } } as any,
    allowed ? ID : UNKNOWN,
  )
  const requests: any[] = []
  let calls = 0
  const client = { sendRequest: async (r: any) => { requests.push(r); return { status } } }
  const adapter: any = {
    continueConversationAsync: async (_a: string, _r: any, logic: any) => {
      calls++
      await logic({ adapter: { ConnectorClientKey: KEY }, turnState: new Map([[KEY, client]]) })
    },
  }
  const r = createReplySender({ config, adapter, allowlist, refs })
  return { r, requests, calls: () => calls }
}

describe('react', () => {
  test('R1 PUT and DELETE urls', async () => {
    const { r, requests } = setup()
    await r.react(CONV, 'm:1', '1f440_eyes')
    await r.react(CONV, 'm:1', '1f440_eyes', true)
    expect(requests[0].method).toBe('PUT')
    expect(requests[0].url).toBe(
      'https://smba.trafficmanager.net/emea/v3/conversations/19%3Aabc%40thread.tacv2%3Bmessageid%3D1/activities/m%3A1/reactions/1f440_eyes',
    )
    expect(requests[1].method).toBe('DELETE')
  })

  test('R2 validation, gate, http error', async () => {
    const s = setup()
    for (const bad of ['Eyes!', '', 'a'.repeat(41)]) {
      await expect(s.r.react(CONV, 'm:1', bad)).rejects.toThrow('invalid reaction id')
    }
    for (const bad of ['', 'x'.repeat(513)]) {
      await expect(s.r.react(CONV, bad, 'like')).rejects.toThrow('invalid message id')
    }
    expect(s.calls()).toBe(0)
    await expect(setup(200, false).r.react(CONV, 'm:1', 'like')).rejects.toBeInstanceOf(ConversationNotAllowedError)
    await expect(setup(404).r.react(CONV, 'm:1', 'like')).rejects.toThrow('react failed (HTTP 404)')
  })
})

describe('reactionEventText', () => {
  test('R3', () => {
    expect(reactionEventText({ reactionsAdded: [{ type: 'like' }, { type: '1f440_eyes' }] })).toBe('[Reaction] added: like, 1f440_eyes')
    expect(reactionEventText({ reactionsRemoved: [{ type: 'heart' }] })).toBe('[Reaction] removed: heart')
    expect(reactionEventText({ reactionsAdded: [{ type: 'BAD TYPE' }, { type: 5 }, null] })).toBeUndefined()
    expect(reactionEventText({ reactionsAdded: [], reactionsRemoved: [] })).toBeUndefined()
    const many = Array.from({ length: 12 }, (_, i) => ({ type: `r${i}` }))
    expect(reactionEventText({ reactionsAdded: many })!.split(': ')[1]!.split(', ').length).toBe(10)
  })
})

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function act(o: Partial<Activity> = {}): Activity {
  return {
    type: 'messageReaction',
    id: 'react-1',
    replyToId: 'bot-msg-1',
    reactionsAdded: [{ type: 'like' }],
    timestamp: new Date('2026-05-17T19:30:00Z'),
    from: { id: 'u', name: 'U', aadObjectId: ID },
    conversation: { id: 'conv-abc', tenantId: TENANT },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    ...o,
  } as unknown as Activity
}

describe('adapter reactions', () => {
  function build() {
    const dir = tempDir('cct-react-ad-')
    const allowlist = createAllowlist(join(dir, 'a.json'))
    allowlist.addEntry(ID)
    const events: any[] = []
    const sent: any[] = []
    const entries: any[] = []
    const handler = makeTurnHandler({
      config,
      adapter: {} as any,
      allowlist,
      refs: createConversationRefStore(),
      onEvent: e => events.push(e),
      pending: { recordIncoming: (x: any) => { entries.push(x); return { action: 'send_new', entry: { code: 'X', pair_id: 'p' } } } } as any,
      sendPairingDm: async (...a: any[]) => { sent.push(a) },
    })
    return { handler, events, sent, entries }
  }
  const run = (h: any, a: Activity) => h(new TurnContext(new NoopAdapter() as any, a))

  test('R4 allowlisted sender emits event', async () => {
    const b = build()
    await run(b.handler, act())
    expect(b.events.length).toBe(1)
    expect(b.events[0].text).toBe('[Reaction] added: like')
    expect(b.events[0].messageId).toBe('bot-msg-1')
    // marked like an edit: an update about an earlier message, never part of a text burst
    expect(b.events[0].messageEvent).toBe('reaction')
    expect(channelEventMeta(b.events[0]).message_event).toBe('reaction')
  })

  test('R4 unknown sender: no event, no pending', async () => {
    const b = build()
    await run(b.handler, act({ from: { id: 'x', name: 'X', aadObjectId: UNKNOWN } as any }))
    expect(b.events).toEqual([])
    expect(b.entries).toEqual([])
    expect(b.sent).toEqual([])
  })

  test('R4 team channel without mention still emits', async () => {
    const b = build()
    await run(b.handler, act({ conversation: { id: '19:x@thread.tacv2', tenantId: TENANT, conversationType: 'channel' } as any }))
    expect(b.events.length).toBe(1)
  })
})
