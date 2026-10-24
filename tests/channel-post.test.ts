import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync, rmSync, statSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createChannelRefStore } from '../src/teams/channelRefs'
import { createChannelPoster, ChannelNotEnabledError, UnknownChannelError } from '../src/teams/channelPost'
import { ConversationNotAllowedError } from '../src/teams/reply'
import { proactiveChannelsFrom } from '../src/config'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const OTHER_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const CHAN = '19:chan1@thread.tacv2'
const STORED_REF = { conversation: { id: `${CHAN};messageid=111` }, serviceUrl: 'https://smba.example/' }

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function makeConfig(overrides = {}) {
  return {
    stateDir: '/tmp/unused',
    allowlistFile: '/tmp/unused',
    appId: 'app-id',
    appPassword: 'app-pw',
    appType: 'SingleTenant' as const,
    tenantId: TENANT_ID,
    bindHost: '127.0.0.1',
    port: 0,
    receivedFilesDir: '/tmp/unused',
    proactiveChannels: new Set([CHAN]),
    ...overrides,
  }
}

describe('post_to_channel', () => {
  let dir: string
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cct-chanpost-'))
  })

  function setup(sendResult: any = [{ conversation: { id: `${CHAN};messageid=999` }, serviceUrl: 'https://smba.example/' }, 'act-1']) {
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(FIXTURE_ID)
    const refs = createConversationRefStore()
    const channelRefs = createChannelRefStore(join(dir, 'channel-refs.json'))
    const calls: any[] = []
    const sends: any[] = []
    const adapter = {
      continueConversationAsync: async (appId: string, ref: any, cb: any) => {
        calls.push({ appId, ref })
        await cb({} as any)
      },
    }
    const poster = createChannelPoster({
      config: makeConfig() as any,
      adapter: adapter as any,
      allowlist,
      refs,
      channelRefs,
      sendToChannel: async (_ctx, activity, channelId, appId) => {
        sends.push({ activity, channelId, appId })
        return sendResult
      },
    })
    return { poster, refs, channelRefs, calls, sends }
  }

  test('rejects channel not enabled', async () => {
    const { poster, calls } = setup()
    await expect(poster.postToChannel('19:other@thread.tacv2', 'hi')).rejects.toBeInstanceOf(ChannelNotEnabledError)
    expect(calls.length).toBe(0)
  })

  test('rejects enabled channel without stored ref', async () => {
    const { poster } = setup()
    await expect(poster.postToChannel(CHAN, 'hi')).rejects.toBeInstanceOf(UnknownChannelError)
  })

  test('rejects stored ref whose user is not allowlisted', async () => {
    const { poster, channelRefs } = setup()
    channelRefs.put(CHAN, STORED_REF as any, OTHER_ID)
    await expect(poster.postToChannel(CHAN, 'hi')).rejects.toBeInstanceOf(ConversationNotAllowedError)
  })

  test('happy path', async () => {
    const { poster, channelRefs, refs, calls, sends } = setup()
    channelRefs.put(CHAN, STORED_REF as any, FIXTURE_ID)
    const r = await poster.postToChannel(CHAN, 'hello')
    // The message id is the ';messageid=' part of the new thread id, not the
    // second value of the send (CloudAdapter returns the conversation id there).
    expect(r).toEqual({ conversationId: `${CHAN};messageid=999`, messageId: '999' })
    expect(calls.length).toBe(1)
    expect(calls[0].appId).toBe('app-id')
    expect(calls[0].ref).toEqual(STORED_REF)
    expect(sends[0].channelId).toBe(CHAN)
    expect(sends[0].appId).toBe('app-id')
    expect(sends[0].activity).toEqual({ type: 'message', text: 'hello', textFormat: 'markdown' })
    expect(refs.get(`${CHAN};messageid=999`)).toBeDefined()
  })

  // Through CloudAdapter the second value is the conversation id again, so it is
  // never reported as a message id; the post itself succeeded, so no throw.
  test('a thread id without ;messageid= returns no message id (and does not throw)', async () => {
    const { poster, channelRefs, refs } = setup([{ conversation: { id: 'conv-x' } }, 'conv-x'])
    channelRefs.put(CHAN, STORED_REF as any, FIXTURE_ID)
    expect(await poster.postToChannel(CHAN, 'hi')).toEqual({ conversationId: 'conv-x' })
    expect(refs.get('conv-x')).toBeDefined()
  })

  test('an unchanged route is not rewritten', () => {
    const path = join(dir, 'channel-refs.json')
    const store = createChannelRefStore(path)
    store.put(CHAN, STORED_REF as any, FIXTURE_ID)
    const first = store.get(CHAN)!.savedAt
    const ino = statSync(path).ino
    store.put(CHAN, { ...STORED_REF } as any, FIXTURE_ID.toUpperCase())
    expect(store.get(CHAN)!.savedAt).toBe(first)
    expect(statSync(path).ino).toBe(ino)
    store.put(CHAN, { ...STORED_REF, serviceUrl: 'https://smba.other/' } as any, FIXTURE_ID)
    expect(store.get(CHAN)!.ref.serviceUrl).toBe('https://smba.other/')
  })

  test("proactiveChannelsFrom drops '*' and keeps listed ids", () => {
    expect(proactiveChannelsFrom(` ${CHAN} , * `)).toEqual(new Set([CHAN]))
    expect(proactiveChannelsFrom(undefined).size).toBe(0)
  })

  test('empty Teams result rejects', async () => {
    const { poster, channelRefs } = setup([{}, ''])
    channelRefs.put(CHAN, STORED_REF as any, FIXTURE_ID)
    await expect(poster.postToChannel(CHAN, 'hi')).rejects.toThrow('returned no conversation')
  })

  test('channelRefs persistence', () => {
    const path = join(dir, 'channel-refs.json')
    expect(createChannelRefStore(path).get(CHAN)).toBeUndefined()
    createChannelRefStore(path).put(CHAN, STORED_REF as any, 'ABCDEF')
    const got = createChannelRefStore(path).get(CHAN)
    expect(got?.aadObjectId).toBe('abcdef')
    expect(got?.ref).toEqual(STORED_REF as any)
    expect(statSync(path).mode & 0o777).toBe(0o600)
    writeFileSync(path, 'not json')
    expect(createChannelRefStore(path).get(CHAN)).toBeUndefined()
  })

  test('adapter captures refs only for enabled channels', async () => {
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(FIXTURE_ID)
    const channelRefs = createChannelRefStore(join(dir, 'channel-refs.json'))
    const turn = makeTurnHandler({
      config: makeConfig() as any,
      adapter: {} as any,
      allowlist,
      refs: createConversationRefStore(),
      onEvent: () => {},
      channelRefs,
    })
    const act = (id: string) =>
      ({
        type: 'message',
        text: '<at>Bot</at> hi',
        id: 'msg-1',
        timestamp: new Date(),
        from: { id: 'user-1', name: 'U', aadObjectId: FIXTURE_ID },
        conversation: { id, tenantId: TENANT_ID, conversationType: 'channel' },
        recipient: { id: 'bot-1', name: 'Bot' },
        entities: [{ type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-1', name: 'Bot' } }],
        channelId: 'msteams',
        serviceUrl: 'https://smba.example/',
      }) as unknown as Activity
    // A non-allowlisted sender in an enabled channel never supplies the route.
    const stranger = act(`${CHAN};messageid=110`)
    ;(stranger.from as any).aadObjectId = OTHER_ID
    await turn(new TurnContext(new NoopAdapter() as any, stranger))
    expect(channelRefs.get(CHAN)).toBeUndefined()
    await turn(new TurnContext(new NoopAdapter() as any, act(`${CHAN};messageid=111`)))
    await turn(new TurnContext(new NoopAdapter() as any, act('19:chan2@thread.tacv2;messageid=1')))
    expect(channelRefs.get(CHAN)).toBeDefined()
    expect(channelRefs.get('19:chan2@thread.tacv2')).toBeUndefined()
    rmSync(dir, { recursive: true })
  })
})
