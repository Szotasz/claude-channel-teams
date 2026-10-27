/**
 * get_meeting_transcript is bound to the conversation the meeting was seen in, and the text it returns has no
 * invisible characters. The tool case in src/server.ts runs on import, so the request check is a function
 * (resolveTranscriptRequest) tested directly, and the server side is read from the source, scoped to the case.
 */
import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import {
  TRANSCRIPT_TOOL,
  cleanOutputText,
  cleanTitle,
  createMeetingStore,
  resolveTranscriptRequest,
  vttToText,
} from '../src/teams/meetingTranscripts'

const BASE = '19:meeting@thread.v2'
const CONVERSATION_ID = `${BASE};messageid=42`
const OTHER = '19:other@thread.v2'
const ORGANIZER = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee'
const LISTED: ReadonlySet<string> = new Set([BASE])

function storeWithMeeting() {
  const store = createMeetingStore(join(mkdtempSync(join(tmpdir(), 'cct-binding-')), 'meetings.json'))
  store.put('MEETING-ID', { graphMeetingId: 'graph-id-1', organizerId: ORGANIZER, conversationId: CONVERSATION_ID })
  return store
}

const errorOf = (fn: () => unknown): string => {
  try { fn() } catch (e) { return e instanceof Error ? e.message : String(e) }
  return '<no error>'
}

describe('resolveTranscriptRequest', () => {
  test('B1 a missing, empty or over-long meeting_id', () => {
    const store = storeWithMeeting()
    for (const meeting_id of [undefined, '', 7, 'x'.repeat(301)]) {
      expect(errorOf(() => resolveTranscriptRequest({ meeting_id, conversation_id: CONVERSATION_ID }, store, LISTED)))
        .toBe('get_meeting_transcript requires a meeting_id string')
    }
  })

  test('B2 a missing, empty or non-string conversation_id', () => {
    const store = storeWithMeeting()
    for (const conversation_id of [undefined, '', 7]) {
      expect(errorOf(() => resolveTranscriptRequest({ meeting_id: 'MEETING-ID', conversation_id }, store, LISTED)))
        .toBe('get_meeting_transcript requires a conversation_id string')
    }
  })

  test('B3 an unknown meeting_id: the fixed error text', () => {
    expect(errorOf(() => resolveTranscriptRequest({ meeting_id: 'NOPE', conversation_id: CONVERSATION_ID }, storeWithMeeting(), LISTED)))
      .toBe('unknown meeting_id: only meetings this channel saw a start or end event for')
  })

  test('B4 the meeting of another conversation is refused', () => {
    expect(errorOf(() => resolveTranscriptRequest({ meeting_id: 'MEETING-ID', conversation_id: OTHER }, storeWithMeeting(), new Set([BASE, OTHER]))))
      .toBe('the meeting_id does not belong to this conversation_id: pass the conversation_id of the same meeting_event tag')
  })

  test("B5 the same conversation, with or without ';messageid=', gives the stored meeting", () => {
    const store = storeWithMeeting()
    for (const conversation_id of [CONVERSATION_ID, BASE, `${BASE};messageid=99`]) {
      expect(resolveTranscriptRequest({ meeting_id: 'MEETING-ID', conversation_id }, store, LISTED).graphMeetingId).toBe('graph-id-1')
    }
  })

  test('B6 a stored meeting whose conversation is no longer on TEAMS_MEETING_CONVERSATIONS is refused', () => {
    const store = storeWithMeeting()
    for (const list of [undefined, new Set<string>(), new Set([OTHER]), new Set(['*'])]) {
      expect(errorOf(() => resolveTranscriptRequest({ meeting_id: 'MEETING-ID', conversation_id: CONVERSATION_ID }, store, list)))
        .toBe('the meeting is from a conversation not on TEAMS_MEETING_CONVERSATIONS')
    }
  })

  test("B6b '*' on the list matches no conversation, not even one whose id is literally '*'", () => {
    const store = createMeetingStore(join(mkdtempSync(join(tmpdir(), 'cct-binding-')), 'meetings.json'))
    store.put('STAR-MEETING', { graphMeetingId: 'graph-id-2', organizerId: ORGANIZER, conversationId: '*' })
    for (const list of [new Set(['*']), new Set(['*', BASE])]) {
      expect(errorOf(() => resolveTranscriptRequest({ meeting_id: 'STAR-MEETING', conversation_id: '*' }, store, list)))
        .toBe('the meeting is from a conversation not on TEAMS_MEETING_CONVERSATIONS')
    }
  })

  test('B7 the tool schema requires conversation_id next to meeting_id', () => {
    const schema = TRANSCRIPT_TOOL.inputSchema as { properties: Record<string, { type?: string }>; required: string[] }
    expect(schema.properties.conversation_id?.type).toBe('string')
    expect([...schema.required].sort()).toEqual(['conversation_id', 'meeting_id'])
  })
})

describe('invisible characters in what the tool returns', () => {
  test('I1 cleanOutputText removes a run of tag characters (hidden ASCII)', () => {
    expect(cleanOutputText('a\u{E0041}\u{E0042}\u{E007F}b')).toBe('ab')
  })

  test('I2 cleanTitle removes tag characters too', () => {
    expect(cleanTitle('Kick\u{E0068}\u{E0069}-off')).toBe('Kick-off')
  })

  test('I3 a transcript cue and its speaker lose their tag characters', () => {
    const vtt = 'WEBVTT\n\n00:00:00.000 --> 00:00:01.000\n<v Speaker\u{E0041} One>Hel\u{E0042}lo</v>\n'
    expect(vttToText(vtt)).toBe('Speaker One: Hello')
  })
})

const server = readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8')

/** The block that starts at `anchor`, from its first '{' up to the matching closing brace. */
function block(src: string, anchor: string): string {
  const start = src.indexOf(anchor)
  expect(start).toBeGreaterThan(-1)
  expect(src.indexOf(anchor, start + 1)).toBe(-1)
  let depth = 0
  for (let i = src.indexOf('{', start); i >= 0 && i < src.length; i++) {
    if (src[i] === '{') depth++
    else if (src[i] === '}' && --depth === 0) return src.slice(start, i + 1)
  }
  throw new Error(`unbalanced block: ${anchor}`)
}

describe('server.ts get_meeting_transcript case', () => {
  test('the request goes through resolveTranscriptRequest with the opt-in list', () => {
    const c = block(server, "case 'get_meeting_transcript': {")
    expect(c).toContain('const m = resolveTranscriptRequest(args, meetingStore, config.meetingConversations)')
    expect(c).not.toContain('meetingStore.get(')
  })

  test('the result text goes through withInvisibleStripped', () => {
    const c = block(server, "case 'get_meeting_transcript': {")
    expect(c).toContain('text: withInvisibleStripped({ text: transcriptResultText(m.title, t) }).text')
  })
})
