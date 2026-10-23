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

export const ACTIVITY_TOOL = {
  name: 'send_activity_notification',
  description:
    'Send a Teams activity feed notification to one person about a team channel or channel thread, from an inbound message. ' +
    "Needs the app's TeamsActivity.Send.Group grant in that team.",
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
