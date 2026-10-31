/**
 * Channel events: a channel was created, renamed, deleted or restored in a team where the bot is installed. Teams
 * sends these as conversationUpdate activities; they are delivered to Claude only with TEAMS_CHANNEL_EVENTS.
 * The names are user content (anyone who can create a channel in the team sets them, and the event carries no
 * allowlisted sender), so they are cleaned of control, line-separator and bidi characters and capped.
 */

export const MAX_EVENT_NAME = 100

export type ChannelEventKind = 'created' | 'renamed' | 'deleted' | 'restored'

export interface ChannelEventInfo {
  kind: ChannelEventKind
  channelId: string
  channelName?: string
  teamId?: string
  teamName?: string
}

const KINDS: Record<string, ChannelEventKind> = {
  channelCreated: 'created',
  channelRenamed: 'renamed',
  channelDeleted: 'deleted',
  channelRestored: 'restored',
}

const UNSAFE = /[\u0000-\u001f\u007f-\u009f\u2028\u2029\u202a-\u202e\u2066-\u2069]/g

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

export function cleanEventName(v: unknown): string | undefined {
  if (typeof v !== 'string') return undefined
  // Trimmed again after the cut: a replaced character just before the cap would otherwise end the name in a space.
  const cleaned = Array.from(v.replace(UNSAFE, ' ').trim()).slice(0, MAX_EVENT_NAME).join('').trim()
  return cleaned || undefined
}

export function channelEventFrom(activity: { type?: string; channelData?: unknown }): ChannelEventInfo | undefined {
  if (activity.type !== 'conversationUpdate') return undefined
  const data = activity.channelData
  if (!isObject(data)) return undefined
  const eventType = data.eventType
  const kind = typeof eventType === 'string' && Object.hasOwn(KINDS, eventType) ? KINDS[eventType] : undefined
  if (!kind) return undefined
  const channel = data.channel
  if (!isObject(channel) || typeof channel.id !== 'string' || !channel.id) return undefined
  const channelName = cleanEventName(channel.name)
  const team = data.team
  const teamId = isObject(team) && typeof team.id === 'string' && team.id ? team.id : undefined
  const teamName = isObject(team) ? cleanEventName(team.name) : undefined
  return {
    kind,
    channelId: channel.id,
    ...(channelName ? { channelName } : {}),
    ...(teamId ? { teamId } : {}),
    ...(teamName ? { teamName } : {}),
  }
}

export function channelEventText(e: ChannelEventInfo): string {
  return `[Channel ${e.kind}]` + (e.channelName ? ` ${e.channelName}` : '') + (e.teamName ? ` (team ${e.teamName})` : '')
}
