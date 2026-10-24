import { describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { channelEventMeta, makeTurnHandler, metaText, threadIdFromConversationId } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function makeActivity(overrides: Partial<Activity> = {}): Activity {
  return {
    type: 'message',
    text: 'hello',
    id: 'msg-1',
    timestamp: new Date('2026-05-18T19:43:00Z'),
    from: { id: 'user-1', name: 'Adnan Khan', aadObjectId: FIXTURE_ID },
    conversation: { id: 'conv-abc', tenantId: TENANT_ID },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    ...overrides,
  } as unknown as Activity
}

async function run(overrides: Partial<Activity>): Promise<any> {
  const dir = mkdtempSync(join(tmpdir(), 'cct-inbound-meta-'))
  try {
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(FIXTURE_ID)
    const events: any[] = []
    const turn = makeTurnHandler({
      config: {
        stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'a', appPassword: 'p',
        appType: 'SingleTenant', tenantId: TENANT_ID, bindHost: '127.0.0.1', port: 0,
        receivedFilesDir: '/tmp/unused',
      } as any,
      adapter: {} as any,
      allowlist,
      refs: createConversationRefStore(),
      onEvent: e => events.push(e),
    })
    await turn(new TurnContext(new NoopAdapter() as any, makeActivity(overrides)))
    expect(events.length).toBe(1)
    return events[0]
  } finally {
    rmSync(dir, { recursive: true })
  }
}

const channelConv = {
  id: '19:chan@thread.tacv2;messageid=1716890000000',
  conversationType: 'channel',
  tenantId: TENANT_ID,
}
const botMention = { type: 'mention', mentioned: { id: 'bot-1', name: 'Bot' }, text: '<at>Bot</at>' }
const mention = (id: string, name?: string) => ({ type: 'mention', mentioned: { id, name }, text: `<at>${name}</at>` })
const channel = (overrides: Partial<Activity>) =>
  run({ conversation: channelConv, text: '<at>Bot</at> hi', entities: [botMention], ...overrides } as any)

describe('inbound meta', () => {
  test('M1 personal chat', async () => {
    const e = await run({})
    expect(e.isGroup).toBe(false)
    for (const k of ['teamId', 'teamName', 'teamsChannelId', 'teamsChannelName', 'threadId', 'replyToId', 'mentions']) {
      expect(e[k]).toBeUndefined()
    }
  })

  test('M2 team channel message', async () => {
    const e = await run({
      conversation: channelConv,
      channelData: {
        team: { id: '19:team@thread.tacv2', name: 'Tenders' },
        channel: { id: '19:chan@thread.tacv2', name: 'Solar 2026' },
      },
      replyToId: '1716890000000',
      text: '<at>Bot</at> summarize',
      entities: [botMention, mention('u-2', 'Kiss Anna')],
    } as any)
    expect(e.isGroup).toBe(true)
    expect(e.teamId).toBe('19:team@thread.tacv2')
    expect(e.teamName).toBe('Tenders')
    expect(e.teamsChannelId).toBe('19:chan@thread.tacv2')
    expect(e.teamsChannelName).toBe('Solar 2026')
    expect(e.threadId).toBe('1716890000000')
    expect(e.replyToId).toBe('1716890000000')
    expect(e.mentions).toEqual(['Kiss Anna'])
  })

  test('M8 the notification meta keys', async () => {
    const base = {
      aad_object_id: FIXTURE_ID, conversation_id: 'conv-abc', tenant_id: TENANT_ID, from_name: 'Adnan Khan',
      message_id: 'msg-1', ts: '2026-05-18T19:43:00.000Z',
    }
    expect(channelEventMeta(await run({}))).toEqual({ ...base, is_group: 'false' })
    const e = await channel({
      channelData: { team: { id: 't-1', name: 'Tenders' }, channel: { id: 'c-1', name: 'Solar 2026' } },
      replyToId: 'r-1',
      entities: [botMention, mention('u-2', 'Kiss Anna'), mention('u-3', 'Nagy Bela')],
    } as any)
    expect(channelEventMeta(e)).toEqual({
      ...base, conversation_id: channelConv.id, mentioned: 'true', is_group: 'true', team_id: 't-1', team_name: 'Tenders',
      channel_id: 'c-1', channel_name: 'Solar 2026', thread_id: '1716890000000', reply_to_id: 'r-1',
      mentions: 'Kiss Anna, Nagy Bela',
    })
  })

  test('M3 names are sanitized', async () => {
    const e = await channel({
      channelData: { channel: { name: 'a"b<c>\nd' }, team: { name: 'x'.repeat(300) } },
    } as any)
    expect(e.teamsChannelName).toBe('a b c d')
    expect(e.teamName).toBe('x'.repeat(200))
    const f = await channel({ channelData: { team: { name: '   ' } } } as any)
    expect(f.teamName).toBeUndefined()
  })

  test('M4 malformed channelData does not throw', async () => {
    const e = await channel({ channelData: 'oops' } as any)
    expect(e.teamId).toBeUndefined()
    expect(e.teamsChannelName).toBeUndefined()
    const f = await channel({ channelData: { team: 5, channel: 7 } } as any)
    expect(f.teamId).toBeUndefined()
    expect(f.teamsChannelId).toBeUndefined()
  })

  test('M5 mentions', async () => {
    const dup = await channel({
      entities: [botMention, mention('u-2', 'Kiss Anna'), mention('u-2', 'Kiss Anna')],
    } as any)
    expect(dup.mentions).toEqual(['Kiss Anna'])
    const noName = await channel({ entities: [botMention, mention('u-3')] } as any)
    expect(noName.mentions).toEqual(['u-3'])
    const many = await channel({
      entities: [botMention, ...Array.from({ length: 25 }, (_, i) => mention(`u-${i}`, `P${i}`))],
    } as any)
    expect(many.mentions.length).toBe(20)
  })

  test('M6 threadIdFromConversationId', () => {
    expect(threadIdFromConversationId('19:abc@thread.tacv2;messageid=1716890000000')).toBe('1716890000000')
    expect(threadIdFromConversationId('conv-abc')).toBeUndefined()
    expect(threadIdFromConversationId('x;messageid=123;other=1')).toBe('123')
    expect(threadIdFromConversationId('x;messageid=')).toBeUndefined()
  })

  test('M7 metaText', () => {
    expect(metaText(42)).toBeUndefined()
    expect(metaText('  a  b ')).toBe('a b')
    expect(metaText('x'.repeat(250))!.length).toBe(200)
  })

  test('M9 a personal chat gets no team, channel, thread, reply or mention context', async () => {
    const e = await run({
      entities: [mention('u-2', 'Kiss Anna')],
      replyToId: 'r-1',
      channelData: { team: { id: 't-1', name: 'Tenders' }, channel: { id: 'c-1', name: 'Solar' } },
      conversation: { id: 'a:1xyz;messageid=99', conversationType: 'personal', tenantId: TENANT_ID },
    } as any)
    expect(e.isGroup).toBe(false)
    for (const k of ['teamId', 'teamName', 'teamsChannelId', 'teamsChannelName', 'threadId', 'replyToId', 'mentions']) {
      expect(e[k]).toBeUndefined()
    }
  })

  test('M10 C1 control characters and bidi controls are replaced', () => {
    expect(metaText('Sales\u009b31m\u0085x\u202eAdmin\u2066y')).toBe('Sales 31m x Admin y')
  })

  test('M11 the joined mentions stay within 200 characters, whole names only', async () => {
    const fit = await channel({
      entities: [botMention, mention('u-1', 'a'.repeat(98)), mention('u-2', 'b'.repeat(98)), mention('u-3', 'c')],
    } as any)
    expect(fit.mentions).toEqual(['a'.repeat(98), 'b'.repeat(98)])
    expect(channelEventMeta(fit).mentions!.length).toBe(198)
    const long = await channel({
      entities: [botMention, ...['x', 'y', 'z'].map((c, i) => mention(`u-${i}`, c.repeat(150)))],
    } as any)
    expect(long.mentions).toEqual(['x'.repeat(150)])
  })
})
