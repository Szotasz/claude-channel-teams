/**
 * Teams activity feed notification (Microsoft Graph, app-only token): pings one
 * person about a team channel or channel thread. Opt-in, see docs/activity-feed.md.
 */
import type { FetchFn } from './threadHistory.js'

const GUID = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/
const MAX_TEXT = 150

export function validateActivityArgs(args: Record<string, unknown>): {
  conversationId: string
  channelId: string
  messageId: string | null
  userId: string
  text: string
} {
  const conversationId = args.conversation_id
  if (typeof conversationId !== 'string' || conversationId.length === 0) {
    throw new Error('conversation_id is required')
  }
  const semi = conversationId.indexOf(';')
  const channelId = semi === -1 ? conversationId : conversationId.slice(0, semi)
  if (!channelId.endsWith('@thread.tacv2')) {
    throw new Error('activity notifications are only available in team channels')
  }
  const marker = ';messageid='
  const at = conversationId.indexOf(marker)
  let messageId: string | null = null
  if (at !== -1) {
    messageId = conversationId.slice(at + marker.length)
    if (!messageId) throw new Error('invalid conversation_id')
  }
  const aad = args.aad_object_id
  if (typeof aad !== 'string' || !GUID.test(aad)) {
    throw new Error('aad_object_id must be an AAD object id')
  }
  const raw = args.text
  const text = typeof raw === 'string' ? raw.trim() : ''
  if (!text) throw new Error('text is required')
  if (text.length > MAX_TEXT) throw new Error(`text must be at most ${MAX_TEXT} characters`)
  return { conversationId, channelId, messageId, userId: aad.toLowerCase(), text }
}

function networkError(err: unknown): Error {
  const code = (err as { code?: unknown } | null)?.code
  const safe = typeof code === 'string' && /^[A-Za-z_]{1,40}$/.test(code) ? code : 'network error'
  return new Error(`Graph request failed (${safe})`)
}

export async function sendActivityNotification(
  input: { teamId: string; channelId: string; messageId: string | null; userId: string; text: string },
  deps: { getToken: () => Promise<string>; fetch?: FetchFn },
): Promise<void> {
  const doFetch: FetchFn = deps.fetch ?? ((u, i) => fetch(u, i))
  const team = `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(input.teamId)}`
  const channelUrl = `${team}/channels/${encodeURIComponent(input.channelId)}`
  const topic = input.messageId ? `${channelUrl}/messages/${encodeURIComponent(input.messageId)}` : channelUrl
  const token = await deps.getToken()
  let res: Response
  try {
    res = await doFetch(`${team}/sendActivityNotification`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'content-type': 'application/json' },
      redirect: 'error',
      body: JSON.stringify({
        topic: { source: 'entityUrl', value: topic },
        activityType: 'systemDefault',
        previewText: { content: input.text },
        recipient: { '@odata.type': 'microsoft.graph.aadUserNotificationRecipient', userId: input.userId },
        templateParameters: [{ name: 'systemDefaultText', value: input.text }],
      }),
    })
  } catch (err) {
    throw networkError(err)
  }
  if (res.status >= 200 && res.status < 300) return
  let code = ''
  try {
    const c = ((await res.json()) as { error?: { code?: unknown } } | null)?.error?.code
    if (typeof c === 'string' && /^[A-Za-z0-9_.]{1,64}$/.test(c)) code = c
  } catch {
    // no JSON body
  }
  throw new Error(`Graph HTTP ${res.status}${code ? ' ' + code : ''}`)
}

export function createActivityRateLimiter(
  opts: { max?: number; windowMs?: number; maxRecipients?: number; now?: () => number } = {},
): { reserve(userId: string): (() => void) | null } {
  const max = opts.max ?? 3
  const windowMs = opts.windowMs ?? 10 * 60 * 1000
  const maxRecipients = opts.maxRecipients ?? 1000
  const now = opts.now ?? Date.now
  const users = new Map<string, number[]>()
  return {
    reserve(userId) {
      const t = now()
      const kept = (users.get(userId) ?? []).filter(x => t - x < windowMs)
      if (kept.length >= max) {
        users.set(userId, kept)
        return null
      }
      if (!users.has(userId) && users.size + 1 > maxRecipients) {
        for (const [id, stamps] of users) {
          if (stamps.every(x => t - x >= windowMs)) users.delete(id)
        }
        // Still full: refuse the new recipient. Dropping a history inside its window would hand that person new slots.
        if (users.size + 1 > maxRecipients) return null
      }
      kept.push(t)
      users.set(userId, kept)
      let released = false
      return () => {
        if (released) return
        released = true
        const list = users.get(userId)
        const i = list ? list.indexOf(t) : -1
        if (list && i !== -1) list.splice(i, 1)
      }
    },
  }
}

export async function handleSendActivity(
  args: Record<string, unknown>,
  deps: {
    enabled: boolean
    assertAllowedConversation: (conversationId: string) => unknown
    teamIdFor: (channelId: string) => string | undefined
    isAllowed: (aadObjectId: string) => boolean
    limiter: { reserve(userId: string): (() => void) | null }
    getToken: () => Promise<string>
    fetch?: FetchFn
  },
): Promise<{ sent: true }> {
  if (!deps.enabled) throw new Error('send_activity_notification is off (TEAMS_ACTIVITY_FEED is not 1)')
  const input = validateActivityArgs(args)
  deps.assertAllowedConversation(input.conversationId)
  const teamId = deps.teamIdFor(input.channelId)
  if (!teamId) throw new Error('team id unknown for this channel: it is learned from an inbound message in the channel')
  if (!deps.isAllowed(input.userId)) throw new Error('the recipient must be on the allowlist')
  const release = deps.limiter.reserve(input.userId)
  if (!release) throw new Error('too many activity notifications to this person (at most 3 in 10 minutes)')
  try {
    await sendActivityNotification(
      { teamId, channelId: input.channelId, messageId: input.messageId, userId: input.userId, text: input.text },
      { getToken: deps.getToken, fetch: deps.fetch },
    )
  } catch (err) {
    release()
    throw err
  }
  return { sent: true }
}

export const ACTIVITY_TOOL = {
  name: 'send_activity_notification',
  description:
    'Send a Teams activity feed notification to one person about a team channel or channel thread, from an inbound message. ' +
    "Needs the app's TeamsActivity.Send.Group grant in that team." +
    ' The recipient must be on the allowlist; at most 3 notifications per person in 10 minutes.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      conversation_id: {
        type: 'string',
        description: 'The conversation_id from an inbound channel <channel> tag.',
      },
      aad_object_id: { type: 'string', description: "The recipient's AAD object id." },
      text: { type: 'string', maxLength: MAX_TEXT, description: 'Notification text (at most 150 characters).' },
    },
    required: ['conversation_id', 'aad_object_id', 'text'],
  },
}

export function activityTools(enabled: boolean): (typeof ACTIVITY_TOOL)[] {
  return enabled ? [ACTIVITY_TOOL] : []
}
