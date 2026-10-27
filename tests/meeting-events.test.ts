import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { channelEventMeta, makeTurnHandler, metaText } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { meetingEventFrom, meetingEventText } from '../src/teams/meetingEvents'

const USER = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const CONVERSATION_ID = '19:meeting@thread.v2;messageid=42'

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

const MEETING_START = 'application/vnd.microsoft.meetingStart'
const MEETING_END = 'application/vnd.microsoft.meetingEnd'

function meetingActivity(o: Record<string, unknown> = {}): Activity {
  return {
    type: 'event',
    channelId: 'msteams',
    name: MEETING_START,
    id: 'ev-1',
    timestamp: new Date('2026-10-23T09:00:05Z'),
    from: { id: '29:1abc' },
    conversation: { id: CONVERSATION_ID, tenantId: TENANT_ID },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelData: { tenant: { id: TENANT_ID }, meeting: { id: 'MEETING-ID' } },
    value: {
      Id: 'VALUE-ID',
      JoinUrl: 'https://join.example/x',
      Title: 'Tender kick-off',
      MeetingType: 'Scheduled',
      StartTime: '2026-10-23T09:00:00Z',
      EndTime: null,
    },
    ...o,
  } as unknown as Activity
}

describe('meetingEventFrom', () => {
  test('P1 a meetingStart: title, type, channelData meeting id and the start time', () => {
    expect(meetingEventFrom(meetingActivity())).toEqual({
      kind: 'started',
      title: 'Tender kick-off',
      meetingType: 'Scheduled',
      meetingId: 'MEETING-ID',
      time: '2026-10-23T09:00:00.000Z',
    })
  })

  test('P2 a meetingEnd: kind ended, the time from EndTime only', () => {
    expect(meetingEventFrom(meetingActivity({
      name: MEETING_END,
      value: { Id: 'VALUE-ID', JoinUrl: null, Title: 'Tender kick-off', MeetingType: 'Scheduled', StartTime: null, EndTime: '2026-10-23T10:30:00Z' },
    }))).toEqual({
      kind: 'ended',
      title: 'Tender kick-off',
      meetingType: 'Scheduled',
      meetingId: 'MEETING-ID',
      time: '2026-10-23T10:30:00.000Z',
    })
  })

  test('P3 no channelData.meeting: no meeting id, and never the id from value', () => {
    const e = meetingEventFrom(meetingActivity({ channelData: { tenant: { id: TENANT_ID } } }))
    expect(e!.meetingId).toBeUndefined()
    expect(JSON.stringify(e)).not.toContain('VALUE-ID')
  })

  test('P4 a start time that is not a date: no time', () => {
    expect(meetingEventFrom(meetingActivity({ value: { Title: 'T', MeetingType: 'Scheduled', StartTime: 'not a date', EndTime: null } }))!.time).toBeUndefined()
  })

  test('P5 StartTime on a meetingEnd is ignored', () => {
    const e = meetingEventFrom(meetingActivity({
      name: MEETING_END,
      value: { Title: 'T', MeetingType: 'Scheduled', StartTime: '2026-10-23T09:00:00Z', EndTime: '2026-10-23T10:30:00Z' },
    }))
    expect(e!.time).toBe('2026-10-23T10:30:00.000Z')
  })

  test('P6 a non-string title: no title', () => {
    expect(meetingEventFrom(meetingActivity({ value: { Title: 5, MeetingType: 'Scheduled', StartTime: '2026-10-23T09:00:00Z', EndTime: null } }))!.title).toBeUndefined()
  })

  test('P7 anything that is not one of the two meeting events', () => {
    expect(meetingEventFrom(meetingActivity({ name: ` ${MEETING_START}` }))).toBeUndefined()
    expect(meetingEventFrom(meetingActivity({ channelId: 'webchat' }))).toBeUndefined()
    expect(meetingEventFrom(meetingActivity({ type: 'message' }))).toBeUndefined()
    expect(meetingEventFrom(meetingActivity({ name: 'application/vnd.microsoft.meetingParticipantJoin' }))).toBeUndefined()
  })

  test('P8 a meeting event whose value is not a plain object is malformed (null)', () => {
    expect(meetingEventFrom(meetingActivity({ value: 'x' }))).toBeNull()
    expect(meetingEventFrom(meetingActivity({ value: null }))).toBeNull()
    expect(meetingEventFrom(meetingActivity({ value: [] }))).toBeNull()
  })
})

describe('meetingEventText', () => {
  test('P9 label plus title, or the label alone', () => {
    expect(meetingEventText({ kind: 'started', title: 'T' })).toBe('[Meeting started] T')
    expect(meetingEventText({ kind: 'ended' })).toBe('[Meeting ended]')
  })
})

/** CONVERSATION_ID without ';messageid=': the form TEAMS_MEETING_CONVERSATIONS takes. */
const BASE_CONVERSATION_ID = '19:meeting@thread.v2'

function setup(
  meetingEvents?: boolean,
  tenantId: string = TENANT_ID,
  // null: no meetingConversations key at all (undefined would take this default).
  meetingConversations: ReadonlySet<string> | null = new Set([BASE_CONVERSATION_ID]),
) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-meeting-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(USER)
  const events: any[] = []
  const config = {
    stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'a', appPassword: 'p',
    appType: 'SingleTenant', tenantId, bindHost: '127.0.0.1', port: 0,
    receivedFilesDir: '/tmp/unused',
    ...(meetingConversations === null ? {} : { meetingConversations }),
  }
  const turn = makeTurnHandler({
    config: config as any,
    adapter: {} as any,
    allowlist,
    refs: createConversationRefStore(),
    onEvent: e => events.push(e),
    ...(meetingEvents === undefined ? {} : { meetingEvents }),
  } as any)
  const run = (a: Activity) => turn(new TurnContext(new NoopAdapter() as any, a))
  return { run, events }
}

describe('meeting events in the turn handler', () => {
  test('E1 a meetingStart is delivered as one event, nothing else in it', async () => {
    const s = setup(true)
    await s.run(meetingActivity())
    expect(s.events.length).toBe(1)
    expect(s.events[0]).toEqual({
      text: '[Meeting started] Tender kick-off',
      conversationId: CONVERSATION_ID,
      aadObjectId: '',
      tenantId: TENANT_ID,
      fromName: '',
      ts: '2026-10-23T09:00:05.000Z',
      meetingEvent: {
        kind: 'started',
        title: 'Tender kick-off',
        meetingType: 'Scheduled',
        meetingId: 'MEETING-ID',
        time: '2026-10-23T09:00:00.000Z',
      },
    })
  })

  test('E2 the title is cleaned before it reaches text and meta', async () => {
    const dirty = `T"i<t>\nl\u202e${'x'.repeat(250)}`
    const s = setup(true)
    await s.run(meetingActivity({ value: { Title: dirty, MeetingType: 'Scheduled', StartTime: '2026-10-23T09:00:00Z', EndTime: null } }))
    const cleaned = `T i t l ${'x'.repeat(192)}`
    expect(s.events[0].text).toBe(`[Meeting started] ${cleaned}`)
    expect(s.events[0].meetingEvent.title).toBe(cleaned)
  })

  test('E3 off by default and off when false', async () => {
    const off = setup()
    await off.run(meetingActivity())
    expect(off.events.length).toBe(0)
    const falsey = setup(false)
    await falsey.run(meetingActivity())
    expect(falsey.events.length).toBe(0)
  })

  test('E4 another tenant is dropped', async () => {
    const s = setup(true)
    await s.run(meetingActivity({
      conversation: { id: CONVERSATION_ID, tenantId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' },
      channelData: { tenant: { id: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' }, meeting: { id: 'MEETING-ID' } },
    }))
    expect(s.events.length).toBe(0)
  })

  test('E5 without a conversation id nothing is delivered', async () => {
    const s = setup(true)
    await s.run(meetingActivity({ conversation: undefined }))
    expect(s.events.length).toBe(0)
  })

  test('E6 a meeting event without a value object is dropped, not delivered', async () => {
    const s = setup(true)
    await s.run(meetingActivity({ value: 'x' }))
    expect(s.events.length).toBe(0)
  })

  test('E7 a normal message still takes the normal path, with no meetingEvent', async () => {
    const s = setup(true)
    await s.run({
      type: 'message',
      text: 'hello',
      id: 'msg-1',
      timestamp: new Date('2026-10-23T09:00:05Z'),
      from: { id: 'user-1', name: 'Adnan', aadObjectId: USER },
      conversation: { id: 'a:personal1', tenantId: TENANT_ID, conversationType: 'personal' },
      recipient: { id: 'bot-1', name: 'Bot' },
      channelId: 'msteams',
      serviceUrl: 'https://smba.trafficmanager.net/uk/',
    } as unknown as Activity)
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('hello')
    expect(s.events[0].meetingEvent).toBeUndefined()
  })

  test('E8 a malformed timestamp is delivered without ts', async () => {
    const s = setup(true)
    await s.run(meetingActivity({ timestamp: 'garbage' }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('[Meeting started] Tender kick-off')
    expect(s.events[0].ts).toBeUndefined()
  })

  test('E9 a foreign tenant named only in channelData is dropped', async () => {
    const foreign = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
    const withEmptyConvTenant = setup(true)
    await withEmptyConvTenant.run(meetingActivity({
      conversation: { id: CONVERSATION_ID, tenantId: '' },
      channelData: { tenant: { id: foreign }, meeting: { id: 'MEETING-ID' } },
    }))
    expect(withEmptyConvTenant.events.length).toBe(0)
    const withoutConvTenant = setup(true)
    await withoutConvTenant.run(meetingActivity({
      conversation: { id: CONVERSATION_ID },
      channelData: { tenant: { id: foreign }, meeting: { id: 'MEETING-ID' } },
    }))
    expect(withoutConvTenant.events.length).toBe(0)
  })

  test('E10 a non-string conversation tenant does not throw', async () => {
    const s = setup(true)
    await s.run(meetingActivity({
      conversation: { id: CONVERSATION_ID, tenantId: 5 },
      channelData: { tenant: { id: TENANT_ID }, meeting: { id: 'MEETING-ID' } },
    }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].tenantId).toBe(TENANT_ID)
  })

  test('E13 an empty channelData tenant id is ignored, not a mismatch', async () => {
    const s = setup(true)
    await s.run(meetingActivity({
      conversation: { id: CONVERSATION_ID, tenantId: TENANT_ID },
      channelData: { tenant: { id: '' }, meeting: { id: 'MEETING-ID' } },
    }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].tenantId).toBe(TENANT_ID)
  })

  test('E14 any one tenant that disagrees drops the event', async () => {
    const foreign = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
    const cases: Array<[unknown, string]> = [[TENANT_ID, foreign], [foreign, TENANT_ID], [5, foreign]]
    for (const [convTenant, dataTenant] of cases) {
      const s = setup(true)
      await s.run(meetingActivity({
        conversation: { id: CONVERSATION_ID, tenantId: convTenant },
        channelData: { tenant: { id: dataTenant }, meeting: { id: 'MEETING-ID' } },
      }))
      expect(s.events.length).toBe(0)
    }
  })

  test('E15 the tenant comparison ignores case', async () => {
    const pinned = 'abcdef00-0000-4000-8000-0000000000aa'
    const s = setup(true, pinned)
    await s.run(meetingActivity({
      conversation: { id: CONVERSATION_ID, tenantId: pinned.toUpperCase() },
      channelData: { tenant: { id: pinned.toUpperCase() }, meeting: { id: 'MEETING-ID' } },
    }))
    expect(s.events.length).toBe(1)
    const upperPinned = setup(true, pinned.toUpperCase())
    await upperPinned.run(meetingActivity({
      conversation: { id: CONVERSATION_ID, tenantId: pinned },
      channelData: { tenant: { id: pinned }, meeting: { id: 'MEETING-ID' } },
    }))
    expect(upperPinned.events.length).toBe(1)
  })

  test('E11 an empty or array conversation id is dropped', async () => {
    // A number or an object already throws in the diagnostic log line above the block; an array does not.
    for (const id of ['', ['19:x']] as unknown[]) {
      const s = setup(true)
      await s.run(meetingActivity({ conversation: { id, tenantId: TENANT_ID } }))
      expect(s.events.length).toBe(0)
    }
  })

  test('E12 the meeting type is cleaned and capped at 40', async () => {
    const meetingType = 'Sch"ed<ul>ed\u202e ' + 'x'.repeat(60)
    const s = setup(true)
    await s.run(meetingActivity({
      value: { Title: 'Tender kick-off', MeetingType: meetingType, StartTime: '2026-10-23T09:00:00Z', EndTime: null },
    }))
    const cleaned = metaText(meetingType, 40)
    expect(s.events.length).toBe(1)
    expect(s.events[0].meetingEvent.meetingType).toBe(cleaned)
    expect(s.events[0].meetingEvent.meetingType.length).toBeLessThanOrEqual(40)
    for (const ch of ['"', '<', '\u202e']) expect(s.events[0].meetingEvent.meetingType).not.toContain(ch)
    expect(channelEventMeta(s.events[0]).meeting_type).toBe(cleaned)
  })
})

describe('channelEventMeta with a meeting event', () => {
  test('M1 the five meeting_ keys, with the cleaned values', () => {
    const meta = channelEventMeta({
      text: '[Meeting started] T',
      conversationId: CONVERSATION_ID,
      aadObjectId: '',
      tenantId: TENANT_ID,
      fromName: '',
      meetingEvent: {
        kind: 'started',
        title: 'Tender kick-off',
        meetingType: 'Scheduled',
        meetingId: 'MEETING-ID',
        time: '2026-10-23T09:00:00.000Z',
      },
    })
    expect(meta).toEqual({
      aad_object_id: '',
      conversation_id: CONVERSATION_ID,
      tenant_id: TENANT_ID,
      from_name: '',
      meeting_event: 'started',
      meeting_title: 'Tender kick-off',
      meeting_type: 'Scheduled',
      meeting_id: 'MEETING-ID',
      meeting_time: '2026-10-23T09:00:00.000Z',
    })
  })

  test('M2 without a meeting event there is no meeting_ key', () => {
    const meta = channelEventMeta({
      text: 'hello',
      conversationId: 'a:personal1',
      aadObjectId: USER,
      tenantId: TENANT_ID,
      fromName: 'Adnan',
    })
    expect(Object.keys(meta).filter(k => k.startsWith('meeting_'))).toEqual([])
  })
})

/** Runs `fn` with process.stderr.write captured; returns what was written. */
async function captureStderr(fn: () => Promise<void>): Promise<string> {
  const original = process.stderr.write.bind(process.stderr)
  let out = ''
  process.stderr.write = ((chunk: any) => { out += String(chunk); return true }) as typeof process.stderr.write
  try { await fn() } finally { process.stderr.write = original }
  return out
}

describe('meeting events only for the conversations in TEAMS_MEETING_CONVERSATIONS', () => {
  test('C1 a conversation that is not on the list is dropped, with one stderr line naming the setting', async () => {
    const s = setup(true, TENANT_ID, new Set(['19:other@thread.v2']))
    const err = await captureStderr(() => s.run(meetingActivity()))
    expect(s.events.length).toBe(0)
    expect(err.split('\n').filter(l => l.startsWith('teams channel: drop'))).toEqual([
      'teams channel: drop — meeting event from a conversation not in TEAMS_MEETING_CONVERSATIONS',
    ])
  })

  test('C2 no list and an empty list deliver nothing', async () => {
    const none = setup(true, TENANT_ID, null)
    await captureStderr(() => none.run(meetingActivity()))
    expect(none.events.length).toBe(0)
    const empty = setup(true, TENANT_ID, new Set())
    await captureStderr(() => empty.run(meetingActivity()))
    expect(empty.events.length).toBe(0)
  })

  test("C3 '*' is not a wildcard: it matches no conversation", async () => {
    const s = setup(true, TENANT_ID, new Set(['*']))
    await captureStderr(() => s.run(meetingActivity()))
    expect(s.events.length).toBe(0)
    // not even a conversation whose id is literally '*'
    for (const id of ['*', '*;messageid=1']) {
      const err = await captureStderr(() => s.run(meetingActivity({ conversation: { id, tenantId: TENANT_ID } })))
      expect(s.events.length).toBe(0)
      expect(err).toContain('meeting event from a conversation not in TEAMS_MEETING_CONVERSATIONS')
    }
  })

  test("C4 the list is matched against the id without ';messageid=', and the listed conversation still gets its event", async () => {
    const s = setup(true, TENANT_ID, new Set(['19:other@thread.v2', BASE_CONVERSATION_ID]))
    await s.run(meetingActivity())
    expect(s.events.length).toBe(1)
    expect(s.events[0].conversationId).toBe(CONVERSATION_ID)
  })

  test('C5 the list is checked after the tenant: a listed conversation from another tenant is still dropped', async () => {
    const s = setup(true, TENANT_ID, new Set([BASE_CONVERSATION_ID]))
    const err = await captureStderr(() => s.run(meetingActivity({
      conversation: { id: CONVERSATION_ID, tenantId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' },
    })))
    expect(s.events.length).toBe(0)
    expect(err).toContain('meeting event from another tenant')
    // an unlisted conversation from another tenant is dropped for the tenant, which is checked first
    const both = await captureStderr(() => s.run(meetingActivity({
      conversation: { id: '19:other@thread.v2', tenantId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' },
    })))
    expect(both).toContain('meeting event from another tenant')
    expect(both).not.toContain('TEAMS_MEETING_CONVERSATIONS')
  })

  test('C6 zero-width characters, a word joiner and a BOM in the title are removed', async () => {
    const s = setup(true)
    await s.run(meetingActivity({
      value: { Title: 'Ten​der kick⁠-off﻿', MeetingType: 'Scheduled', StartTime: '2026-10-23T09:00:00Z', EndTime: null },
    }))
    expect(s.events[0].text).toBe('[Meeting started] Tender kick-off')
    expect(s.events[0].meetingEvent.title).toBe('Tender kick-off')
  })
})
