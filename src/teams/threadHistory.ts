/**
 * Channel thread history through Microsoft Graph (app-only token).
 *
 * Needs the ChannelMessage.Read.Group application permission, granted per team
 * by resource-specific consent. Usable without the server (live tests). The
 * token comes from createGraphTokenProvider in planner.ts.
 */

export type FetchFn = (url: string, init?: RequestInit) => Promise<Response>

const GRAPH_PREFIX = 'https://graph.microsoft.com/v1.0/'
const MAX_TEXT = 4000
const CUT_MARK = ' [...]'

export interface ThreadMessage {
  id: string
  replyToId: string | null
  createdDateTime: string
  from: string
  fromAadObjectId: string | null
  allowlisted: boolean
  text: string
}

function networkError(err: unknown): Error {
  const code = (err as { code?: unknown } | null)?.code
  const safe = typeof code === 'string' && /^[A-Za-z_]{1,40}$/.test(code) ? code : 'network error'
  return new Error(`Graph request failed (${safe})`)
}

const ENTITIES: Record<string, string> = {
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  '&#39;': "'",
  '&nbsp;': ' ',
}

function htmlToText(html: string): string {
  return html
    .replace(/<br\s*\/?>|<\/p\s*>|<\/div\s*>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&(?:amp|lt|gt|quot|#39|nbsp);/g, m => ENTITIES[m]!)
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

/** At most MAX_TEXT code points, the cut mark included. */
function truncate(s: string): string {
  const cps = Array.from(s)
  return cps.length > MAX_TEXT ? cps.slice(0, MAX_TEXT - CUT_MARK.length).join('') + CUT_MARK : s
}

/** A next link may only point at Graph v1.0 (checked after URL normalisation, so '/v1.0/../beta' fails). */
function isGraphV1Link(link: string): boolean {
  try {
    const u = new URL(link)
    return link.startsWith(GRAPH_PREFIX) && u.origin === 'https://graph.microsoft.com' && u.pathname.startsWith('/v1.0/')
  } catch {
    return false
  }
}

interface RawMessage {
  id?: string
  replyToId?: string | null
  createdDateTime?: string
  messageType?: string
  deletedDateTime?: string | null
  from?: { user?: { id?: string; displayName?: string }; application?: { displayName?: string } } | null
  body?: { contentType?: string; content?: string }
}

export async function fetchThreadHistory(
  params: { teamId: string; channelId: string; rootMessageId: string; limit?: number },
  deps: {
    getToken: () => Promise<string>
    fetch?: FetchFn
    isAllowed?: (aadObjectId: string) => boolean
  },
): Promise<ThreadMessage[]> {
  const limit = params.limit ?? 30
  if (!Number.isInteger(limit) || limit < 1 || limit > 100) {
    throw new Error('limit must be an integer between 1 and 100')
  }
  const doFetch: FetchFn = deps.fetch ?? ((u, i) => fetch(u, i))

  async function get(url: string, isRoot: boolean): Promise<unknown> {
    const token = await deps.getToken()
    let res: Response
    try {
      res = await doFetch(url, {
        method: 'GET',
        headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
        redirect: 'error',
      })
    } catch (err) {
      throw networkError(err)
    }
    if (res.status === 403) {
      throw new Error(
        'Graph refused (403): the app needs the ChannelMessage.Read.Group grant for this team (manifest --rsc-all-messages)',
      )
    }
    if (res.status === 404 && isRoot) throw new Error('thread not found')
    if (res.status < 200 || res.status >= 300) {
      throw new Error(`Graph request failed (HTTP ${res.status})`)
    }
    try {
      return await res.json()
    } catch {
      throw new Error('Graph request failed (invalid response)')
    }
  }

  const convert = (m: RawMessage): ThreadMessage => {
    const user = m.from?.user
    const aad = typeof user?.id === 'string' && user.id ? user.id.toLowerCase() : null
    const content = m.body?.content ?? ''
    const text = m.body?.contentType === 'html' ? htmlToText(content) : content
    return {
      id: m.id ?? '',
      replyToId: m.replyToId ?? null,
      createdDateTime: m.createdDateTime ?? '',
      from: user?.displayName ?? m.from?.application?.displayName ?? '',
      fromAadObjectId: aad,
      allowlisted: aad !== null && deps.isAllowed?.(aad) === true,
      text: truncate(text),
    }
  }
  const keep = (m: RawMessage) => m.messageType === 'message' && (m.deletedDateTime ?? null) === null

  const base =
    `${GRAPH_PREFIX}teams/${encodeURIComponent(params.teamId)}` +
    `/channels/${encodeURIComponent(params.channelId)}` +
    `/messages/${encodeURIComponent(params.rootMessageId)}`

  const root = (await get(base, true)) as RawMessage
  const out: ThreadMessage[] = []
  if (root && keep(root)) out.push(convert(root))

  const replies: ThreadMessage[] = []
  let next: string | undefined = `${base}/replies?$top=${Math.min(50, limit)}`
  while (next && replies.length < limit) {
    const page = (await get(next, false)) as { value?: RawMessage[]; '@odata.nextLink'?: string }
    for (const m of page?.value ?? []) {
      if (replies.length >= limit) break
      if (keep(m)) replies.push(convert(m))
    }
    const link = page?.['@odata.nextLink']
    if (link !== undefined && link !== null) {
      if (typeof link !== 'string' || !isGraphV1Link(link)) {
        throw new Error('unexpected Graph next link')
      }
    }
    next = replies.length < limit ? link : undefined
  }
  replies.sort((a, b) =>
    a.createdDateTime < b.createdDateTime ? -1 : a.createdDateTime > b.createdDateTime ? 1 : a.id < b.id ? -1 : a.id > b.id ? 1 : 0,
  )
  return out.concat(replies)
}

/** The channel id and the thread root id of a channel-thread conversation id. */
export function parseChannelThread(conversationId: string): { channelId: string; rootMessageId: string } {
  const semi = conversationId.indexOf(';')
  const channelId = semi === -1 ? conversationId : conversationId.slice(0, semi)
  if (!channelId.endsWith('@thread.tacv2')) {
    throw new Error('thread history is only available for channel threads')
  }
  const marker = ';messageid='
  const at = conversationId.indexOf(marker)
  const rootMessageId = at === -1 ? '' : conversationId.slice(at + marker.length)
  if (!rootMessageId) {
    throw new Error('this conversation_id is not a channel thread (no ;messageid=)')
  }
  return { channelId, rootMessageId }
}

/** Bounded channel id -> team AAD group id map (oldest dropped). */
export function createTeamGroupStore(max = 1000): {
  set(channelId: string, groupId: string): void
  get(channelId: string): string | undefined
} {
  const m = new Map<string, string>()
  return {
    set(channelId, groupId) {
      m.delete(channelId)
      m.set(channelId, groupId)
      if (m.size > max) m.delete(m.keys().next().value as string)
    },
    get: channelId => m.get(channelId),
  }
}
