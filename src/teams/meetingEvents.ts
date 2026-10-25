/**
 * Parsing of the Teams meeting start / end activities: with
 * TEAMS_MEETING_EVENTS the adapter delivers "[Meeting started] <title>" /
 * "[Meeting ended] <title>" to Claude. The title and the meeting type are set
 * by people in the tenant and the event carries no allowlisted sender, so both
 * are user content and are cleaned with metaText in the adapter — this module
 * only parses and takes no dependency on it (importing the adapter here would
 * be a circular import). The join URL is not passed on: a join link is not
 * needed to act on the event.
 */

export type MeetingEventKind = 'started' | 'ended'

export interface MeetingEventInfo {
  kind: MeetingEventKind
  title?: string
  meetingType?: string
  meetingId?: string
  time?: string
}

export const MEETING_START = 'application/vnd.microsoft.meetingStart'
export const MEETING_END = 'application/vnd.microsoft.meetingEnd'

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

function str(v: unknown): string | undefined {
  return typeof v === 'string' && v.length > 0 ? v : undefined
}

/**
 * Recognise a Teams meeting start / end activity. `undefined` means "not a
 * meeting event" (some other activity), `null` means "a malformed one" (one of
 * the two names but no plain-object `value`). `value.Id` and `value.JoinUrl`
 * are never read: Microsoft says to use channelData.meeting.id, not the id from
 * `value`, and the join URL is not passed on.
 */
export function meetingEventFrom(activity: {
  type?: string
  channelId?: string
  name?: string
  value?: unknown
  channelData?: unknown
}): MeetingEventInfo | null | undefined {
  if (activity.type !== 'event' || activity.channelId !== 'msteams') return undefined
  const kind: MeetingEventKind | undefined =
    activity.name === MEETING_START ? 'started' : activity.name === MEETING_END ? 'ended' : undefined
  if (kind === undefined) return undefined
  const value = activity.value
  if (!isObject(value)) return null
  const channelData = isObject(activity.channelData) ? activity.channelData : undefined
  const meeting = channelData !== undefined && isObject(channelData.meeting) ? channelData.meeting : undefined
  const title = str(value.Title)
  const meetingType = str(value.MeetingType)
  const meetingId = meeting !== undefined ? str(meeting.id) : undefined
  // Teams sends the not-applicable one as null: StartTime on a meetingEnd and
  // EndTime on a meetingStart.
  const rawTime = kind === 'started' ? value.StartTime : value.EndTime
  let time: string | undefined
  if (typeof rawTime === 'string' && rawTime.length > 0) {
    const d = new Date(rawTime)
    if (!Number.isNaN(d.getTime())) time = d.toISOString()
  }
  return {
    kind,
    ...(title !== undefined ? { title } : {}),
    ...(meetingType !== undefined ? { meetingType } : {}),
    ...(meetingId !== undefined ? { meetingId } : {}),
    ...(time !== undefined ? { time } : {}),
  }
}

export function meetingEventText(e: { kind: MeetingEventKind; title?: string }): string {
  return (e.kind === 'started' ? '[Meeting started]' : '[Meeting ended]') + (e.title ? ` ${e.title}` : '')
}
