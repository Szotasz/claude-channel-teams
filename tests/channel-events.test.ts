import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { channelEventFrom, channelEventText, cleanEventName, MAX_EVENT_NAME } from '../src/teams/channelEvents'
import { makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { channelEventsWarning, loadConfig, parseFlag } from '../src/config'

describe('cleanEventName', () => {
  test('control, C1, line separator and bidi characters become spaces, trimmed', () => {
    expect(cleanEventName('a\u0000b\u001fc\u007fd\u009fe\u2028f\u2029g\u202ah\u202ei\u2066j\u2069k')).toBe('a b c d e f g h i j k')
    expect(cleanEventName('\n\u202e x \u2066')).toBe('x')
  })
  test('capped at 100 code points without splitting an emoji', () => {
    expect(cleanEventName('a'.repeat(150))).toBe('a'.repeat(MAX_EVENT_NAME))
    const s = cleanEventName('a'.repeat(99) + '😀😀')!
    expect(s).toBe('a'.repeat(99) + '😀')
  })
  test('no trailing space when a replaced character sits just before the cap', () => {
    expect(cleanEventName('a'.repeat(99) + '\u0000' + 'b')).toBe('a'.repeat(99))
    // Leading spaces do not count against the cap.
    expect(cleanEventName(' '.repeat(5) + 'a'.repeat(100))).toBe('a'.repeat(100))
  })
  test('empty and non-string', () => {
    expect(cleanEventName('')).toBeUndefined()
    expect(cleanEventName('   ')).toBeUndefined()
    expect(cleanEventName(5)).toBeUndefined()
  })
})

const cd = (eventType: string, extra: Record<string, unknown> = {}) => ({
  type: 'conversationUpdate',
  channelData: {
    eventType,
    channel: { id: '19:c@thread.tacv2', name: 'Project X' },
    team: { id: '19:t@thread.tacv2', name: 'Example Team' },
    ...extra,
  },
})

describe('channelEventFrom', () => {
  for (const [et, kind] of [
    ['channelCreated', 'created'],
    ['channelRenamed', 'renamed'],
    ['channelDeleted', 'deleted'],
    ['channelRestored', 'restored'],
  ] as const) {
    test(et, () => {
      expect(channelEventFrom(cd(et))).toEqual({
        kind,
        channelId: '19:c@thread.tacv2',
        channelName: 'Project X',
        teamId: '19:t@thread.tacv2',
        teamName: 'Example Team',
      })
    })
  }
  test('other cases', () => {
    expect(channelEventFrom(cd('teamMemberAdded'))).toBeUndefined()
    expect(channelEventFrom({ ...cd('channelCreated'), type: 'message' })).toBeUndefined()
    expect(channelEventFrom(cd('channelCreated', { channel: { name: 'x' } }))).toBeUndefined()
    expect(channelEventFrom(cd('channelCreated', { channel: { id: '', name: 'x' } }))).toBeUndefined()
    // An eventType that is a key of Object.prototype is not one of the four.
    expect(channelEventFrom(cd('toString'))).toBeUndefined()
    expect(channelEventFrom(cd('constructor'))).toBeUndefined()
    const emptyTeamId = channelEventFrom(cd('channelCreated', { team: { id: '', name: 'T' } }))!
    expect('teamId' in emptyTeamId).toBe(false)
    // The names go through cleanEventName.
    const dirty = channelEventFrom(cd('channelCreated', { channel: { id: 'c', name: 'a\u202eb\nc' }, team: { id: 't', name: 'd\u2028e' } }))!
    expect(dirty.channelName).toBe('a b c')
    expect(dirty.teamName).toBe('d e')
    const noTeam = channelEventFrom(cd('channelCreated', { team: undefined }))!
    expect(noTeam.teamId).toBeUndefined()
    expect('teamId' in noTeam).toBe(false)
  })
})

describe('channelEventText', () => {
  test('with and without names', () => {
    expect(channelEventText({ kind: 'created', channelId: 'c', channelName: 'Project X', teamName: 'Example Team' })).toBe(
      '[Channel created] Project X (team Example Team)',
    )
    expect(channelEventText({ kind: 'deleted', channelId: 'c', channelName: 'Project X' })).toBe('[Channel deleted] Project X')
    expect(channelEventText({ kind: 'renamed', channelId: 'c' })).toBe('[Channel renamed]')
    expect(channelEventText({ kind: 'restored', channelId: 'c', teamName: 'T' })).toBe('[Channel restored] (team T)')
  })
})

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const CONV = '19:c@thread.tacv2;messageid=1'
const FOREIGN_TENANT = 'ffffffff-0000-4000-8000-000000000009'
const LETTERED_TENANT = 'abcdef00-0000-4000-8000-00000000000a'
const TEAM = '19:t@thread.tacv2'

function act(o: Record<string, unknown> = {}): Activity {
  return {
    type: 'conversationUpdate',
    timestamp: new Date('2026-05-18T19:43:00Z'),
    from: { id: '29:abc' },
    conversation: { id: CONV, tenantId: TENANT_ID, conversationType: 'channel' },
    recipient: { id: 'bot-1' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    channelData: {
      eventType: 'channelCreated',
      channel: { id: '19:c@thread.tacv2', name: 'Project X' },
      team: { id: TEAM, name: 'Example Team' },
      tenant: { id: TENANT_ID },
    },
    ...o,
  } as unknown as Activity
}

function setup(
  channelEvents: boolean,
  tenantId = TENANT_ID,
  teams: { channelEventTeams?: ReadonlySet<string> } = { channelEventTeams: new Set([TEAM]) },
) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-chev-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  const events: any[] = []
  const pairing: string[] = []
  const config = {
    stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'a', appPassword: 'p',
    appType: 'SingleTenant', tenantId, bindHost: '127.0.0.1', port: 0,
    receivedFilesDir: '/tmp/unused',
    ...teams,
  }
  const turn = makeTurnHandler({
    config: config as any,
    adapter: {} as any,
    allowlist,
    refs: createConversationRefStore(),
    onEvent: e => events.push(e),
    pending: { recordIncoming: () => { pairing.push('pending'); return { action: 'suppress', reason: 'x' } } } as any,
    sendPairingDm: async () => { pairing.push('dm') },
    channelEvents,
  } as any)
  const run = (a: Activity) => turn(new TurnContext(new NoopAdapter() as any, a))
  return { run, events, pairing, allowlist }
}

describe('adapter channel events', () => {
  test('delivered when enabled', async () => {
    const s = setup(true)
    await s.run(act())
    expect(s.events).toEqual([
      {
        text: '[Channel created] Project X (team Example Team)',
        conversationId: CONV,
        aadObjectId: '',
        tenantId: TENANT_ID,
        fromName: '',
        ts: '2026-05-18T19:43:00.000Z',
        channelEvent: { kind: 'created', channelId: '19:c@thread.tacv2', channelName: 'Project X', teamId: '19:t@thread.tacv2' },
      },
    ])
  })
  test('the person who changed the channel is not reported as a sender', async () => {
    const s = setup(true)
    await s.run(act({ from: { id: '29:abc', name: 'Someone', aadObjectId: '00000000-0000-4000-8000-0000000000AB' } }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].aadObjectId).toBe('')
    expect(s.events[0].fromName).toBe('')
  })
  test('disabled -> no event', async () => {
    const s = setup(false)
    await s.run(act())
    expect(s.events.length).toBe(0)
    expect(s.pairing.length).toBe(0)
  })
  test('another tenant -> no event', async () => {
    const s = setup(true)
    await s.run(act({ conversation: { id: CONV, tenantId: 'ffffffff-0000-4000-8000-000000000009', conversationType: 'channel' } }))
    expect(s.events.length).toBe(0)
  })
  test('the tenant comes from channelData when the conversation has none', async () => {
    const s = setup(true)
    const data = (tenant: string) => ({ ...(act().channelData as object), tenant: { id: tenant } })
    await s.run(act({ conversation: { id: CONV, conversationType: 'channel' }, channelData: data(FOREIGN_TENANT) }))
    expect(s.events.length).toBe(0)
    await s.run(act({ conversation: { id: CONV, conversationType: 'channel' }, channelData: data(TENANT_ID) }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].tenantId).toBe(TENANT_ID)
  })
  test('an empty or matching conversation tenant does not hide a foreign channelData tenant', async () => {
    const s = setup(true)
    const data = { ...(act().channelData as object), tenant: { id: FOREIGN_TENANT } }
    await s.run(act({ conversation: { id: CONV, tenantId: '', conversationType: 'channel' }, channelData: data }))
    await s.run(act({ conversation: { id: CONV, tenantId: TENANT_ID, conversationType: 'channel' }, channelData: data }))
    expect(s.events.length).toBe(0)
  })
  test('an empty conversation tenant counts as none: a matching channelData tenant is delivered', async () => {
    const s = setup(true)
    const data = { ...(act().channelData as object), tenant: { id: TENANT_ID } }
    await s.run(act({ conversation: { id: CONV, tenantId: '', conversationType: 'channel' }, channelData: data }))
    expect(s.events.length).toBe(1)
  })
  test('a non-string conversation tenant does not throw', async () => {
    const s = setup(true)
    await s.run(act({ conversation: { id: CONV, tenantId: 5, conversationType: 'channel' } }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].tenantId).toBe(TENANT_ID)
  })
  test('tenant ids are compared case-insensitively', async () => {
    const s = setup(true, LETTERED_TENANT)
    const data = { ...(act().channelData as object), tenant: { id: LETTERED_TENANT } }
    await s.run(act({ conversation: { id: CONV, tenantId: LETTERED_TENANT.toUpperCase(), conversationType: 'channel' }, channelData: data }))
    expect(s.events.length).toBe(1)
  })
  test('no tenant on the event -> delivered with the pinned tenant, as for messages', async () => {
    const s = setup(true)
    const { tenant: _t, ...noTenant } = act().channelData as Record<string, unknown>
    await s.run(act({ conversation: { id: CONV, conversationType: 'channel' }, channelData: noTenant }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].tenantId).toBe(TENANT_ID)
  })
  test('no pinned tenant -> the event tenant is not compared', async () => {
    const s = setup(true, '')
    await s.run(act({ conversation: { id: CONV, tenantId: FOREIGN_TENANT, conversationType: 'channel' } }))
    expect(s.events.length).toBe(1)
  })
  test('the kind follows the event type', async () => {
    const s = setup(true)
    await s.run(act({ channelData: { ...(act().channelData as object), eventType: 'channelRenamed' } }))
    expect(s.events[0].text).toBe('[Channel renamed] Project X (team Example Team)')
    expect(s.events[0].channelEvent.kind).toBe('renamed')
  })
  test('no timestamp -> no ts', async () => {
    const s = setup(true)
    await s.run(act({ timestamp: undefined }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].ts).toBeUndefined()
  })
  test('no conversation id -> no event', async () => {
    const s = setup(true)
    await s.run(act({ conversation: undefined }))
    expect(s.events.length).toBe(0)
  })
  test('membersAdded conversationUpdate -> no event, nothing else', async () => {
    const s = setup(true)
    await s.run(act({ membersAdded: [{ id: '29:x' }], channelData: { eventType: 'teamMemberAdded', team: { id: 't' } } }))
    await s.run(act({ membersAdded: [{ id: '29:x' }], channelData: undefined }))
    expect(s.events.length).toBe(0)
    expect(s.pairing.length).toBe(0)
  })
  test('a normal message still takes the normal path', async () => {
    const s = setup(true)
    const USER = '00000000-0000-4000-8000-000000000001'
    s.allowlist.addEntry(USER)
    await s.run(act({
      type: 'message', text: 'hi', id: 'm1', channelData: undefined,
      from: { id: 'u', name: 'Ad', aadObjectId: USER },
      conversation: { id: 'a:p1', tenantId: TENANT_ID, conversationType: 'personal' },
    }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('hi')
    expect(s.events[0].channelEvent).toBeUndefined()
  })
})

describe('adapter channel events: team allowlist', () => {
  const data = (team?: unknown) => {
    const { team: _t, ...rest } = act().channelData as Record<string, unknown>
    return team === undefined ? rest : { ...rest, team }
  }
  test('a listed team -> delivered', async () => {
    const s = setup(true, TENANT_ID)
    await s.run(act())
    expect(s.events.length).toBe(1)
    expect(s.events[0].channelEvent.teamId).toBe(TEAM)
  })
  test('a team that is not listed -> no event', async () => {
    const s = setup(true, TENANT_ID, { channelEventTeams: new Set(['19:other@thread.tacv2']) })
    await s.run(act())
    expect(s.events.length).toBe(0)
  })
  test('an empty list -> no event', async () => {
    const s = setup(true, TENANT_ID, { channelEventTeams: new Set() })
    await s.run(act())
    expect(s.events.length).toBe(0)
  })
  test('no list at all -> no event', async () => {
    const s = setup(true, TENANT_ID, {})
    await s.run(act())
    expect(s.events.length).toBe(0)
  })
  test("'*' in the list matches no team", async () => {
    const s = setup(true, TENANT_ID, { channelEventTeams: new Set(['*']) })
    await s.run(act())
    expect(s.events.length).toBe(0)
  })
  test("'*' does not match a team whose id is literally '*'", async () => {
    const s = setup(true, TENANT_ID, { channelEventTeams: new Set(['*']) })
    await s.run(act({ channelData: data({ id: '*', name: 'Example Team' }) }))
    expect(s.events.length).toBe(0)
  })
  test('an event without a team id -> no event even when the list is set', async () => {
    const s = setup(true, TENANT_ID)
    await s.run(act({ channelData: data() }))
    expect(s.events.length).toBe(0)
  })
  test('the list does not turn the feature on', async () => {
    const s = setup(false, TENANT_ID)
    await s.run(act())
    expect(s.events.length).toBe(0)
  })
})

describe('config channelEvents', () => {
  test('parseFlag', () => {
    expect(parseFlag('1')).toBe(true)
    expect(parseFlag(' TRUE ')).toBe(true)
    expect(parseFlag('true')).toBe(true)
    expect(parseFlag('0')).toBe(false)
    expect(parseFlag('yes')).toBe(false)
    expect(parseFlag('')).toBe(false)
    expect(parseFlag(undefined)).toBe(false)
  })
  test('loadConfig', () => {
    const keys = ['TEAMS_CHANNEL_EVENTS', 'TEAMS_STATE_DIR', 'TEAMS_BOT_APP_ID', 'TEAMS_BOT_APP_PASSWORD', 'TEAMS_BOT_TENANT_ID']
    const saved = Object.fromEntries(keys.map(k => [k, process.env[k]]))
    try {
      process.env.TEAMS_STATE_DIR = mkdtempSync(join(tmpdir(), 'cct-cfg-'))
      process.env.TEAMS_BOT_APP_ID = 'a'
      process.env.TEAMS_BOT_APP_PASSWORD = 'p'
      process.env.TEAMS_BOT_TENANT_ID = TENANT_ID
      process.env.TEAMS_CHANNEL_EVENTS = '1'
      expect(loadConfig().channelEvents).toBe(true)
      process.env.TEAMS_CHANNEL_EVENTS = 'true'
      expect(loadConfig().channelEvents).toBe(true)
      delete process.env.TEAMS_CHANNEL_EVENTS
      expect(loadConfig().channelEvents).toBe(false)
    } finally {
      for (const k of keys) {
        if (saved[k] === undefined) delete process.env[k]
        else process.env[k] = saved[k]
      }
    }
  })
  test('loadConfig channelEventTeams', () => {
    const keys = ['TEAMS_CHANNEL_EVENT_TEAMS', 'TEAMS_STATE_DIR', 'TEAMS_BOT_APP_ID', 'TEAMS_BOT_APP_PASSWORD', 'TEAMS_BOT_TENANT_ID']
    const saved = Object.fromEntries(keys.map(k => [k, process.env[k]]))
    try {
      process.env.TEAMS_STATE_DIR = mkdtempSync(join(tmpdir(), 'cct-cfg-'))
      process.env.TEAMS_BOT_APP_ID = 'a'
      process.env.TEAMS_BOT_APP_PASSWORD = 'p'
      process.env.TEAMS_BOT_TENANT_ID = TENANT_ID
      process.env.TEAMS_CHANNEL_EVENT_TEAMS = ' 19:a@thread.tacv2 , ,19:b@thread.tacv2 '
      expect([...loadConfig().channelEventTeams!].sort()).toEqual(['19:a@thread.tacv2', '19:b@thread.tacv2'])
      delete process.env.TEAMS_CHANNEL_EVENT_TEAMS
      expect(loadConfig().channelEventTeams!.size).toBe(0)
    } finally {
      for (const k of keys) {
        if (saved[k] === undefined) delete process.env[k]
        else process.env[k] = saved[k]
      }
    }
  })
  test('channelEventsWarning', () => {
    expect(channelEventsWarning({ channelEvents: false })).toBeUndefined()
    expect(channelEventsWarning({ channelEvents: false, channelEventTeams: new Set(['*']) })).toBeUndefined()
    expect(channelEventsWarning({ channelEvents: true })).toContain('TEAMS_CHANNEL_EVENT_TEAMS is empty')
    expect(channelEventsWarning({ channelEvents: true, channelEventTeams: new Set() })).toContain('TEAMS_CHANNEL_EVENT_TEAMS is empty')
    expect(channelEventsWarning({ channelEvents: true, channelEventTeams: new Set(['*', '19:a@thread.tacv2']) })).toContain('takes no wildcard')
    expect(channelEventsWarning({ channelEvents: true, channelEventTeams: new Set(['19:a@thread.tacv2']) })).toBeUndefined()
  })
})
