/**
 * Files shared in a team CHANNEL or a GROUP CHAT.
 *
 * In a 1:1 chat Teams hands the bot the file directly: the activity carries an
 * attachment with contentType `application/vnd.microsoft.teams.file.download.info`
 * whose content holds a short-lived downloadUrl (see attachments.ts). In a
 * channel or a group chat it does NOT — the activity only carries a `text/html`
 * part containing an empty `<attachment id="…"></attachment>` tag per shared
 * file. The file itself has to be fetched the long way round:
 *
 *   1. the Azure AD group id of the team (channels only — Graph addresses
 *      channels by group, Bot Framework by thread id),
 *   2. the message through Microsoft Graph, whose `attachments` list has the
 *      real entries (contentType `reference`, with id / contentUrl / name),
 *   3. the file from SharePoint/OneDrive through the Graph `shares` endpoint.
 *
 * Every outcome becomes an annotation line, so the agent sees the same
 * "[Attached: …]" / "[Attachment skipped: …]" notes as on the 1:1 path and can
 * tell the user what happened to a file it could not read.
 *
 * Security notes, in the same spirit as attachments.ts:
 *
 *   - The bot's s2s token (getBotToken) only ever goes to the Bot Framework
 *     service URL the activity itself named, and only when that host is on the
 *     inline-image allowlist — otherwise it would leak to any host an attacker
 *     puts in `serviceUrl`.
 *   - The Graph token (getGraphToken) only ever goes to graph.microsoft.com.
 *   - A contentUrl that is not a *.sharepoint.com https link is never fetched.
 *   - The share download is followed manually: the redirect target must be
 *     https without userinfo, and it is fetched WITHOUT any bearer token.
 */

import { mkdirSync } from 'fs'
import { join } from 'path'
import type { Attachment } from 'botframework-schema'
import {
  DOWNLOAD_TIMEOUT_MS,
  type AttachmentResult,
  downloadToTemp,
  formatAnnotation,
  isAllowedInlineImageHost,
  londonDate,
  londonTime,
  logAttachmentResult,
  moveFile,
  pickPath,
  saveResponseToTemp,
  senderSlug,
} from './attachments.js'

export const GRAPH_V1 = 'https://graph.microsoft.com/v1.0'

/** At most this many files are pulled out of one message. */
export const MAX_SHARED_FILES = 10

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
const REF_ID_RE = /^[A-Za-z0-9_-]{1,128}$/

// The placeholder Teams puts in the HTML part for every shared file.
const ATTACHMENT_TAG_RE = /<attachment\b[^>]*\bid\s*=\s*["']([^"']*)["']/gi

const REDIRECT_STATUSES = [301, 302, 303, 307, 308]

/**
 * The ids of the `<attachment id="…">` placeholders in the activity's HTML
 * parts, first-seen order, deduped, at most MAX_SHARED_FILES. These are the ids
 * the Graph message's own attachment entries carry.
 */
export function attachmentRefIds(attachments: Attachment[]): string[] {
  const out: string[] = []
  for (const att of attachments) {
    if (!att || att.contentType !== 'text/html') continue
    if (typeof att.content !== 'string') continue
    for (const m of att.content.matchAll(ATTACHMENT_TAG_RE)) {
      const id = m[1] ?? ''
      if (!REF_ID_RE.test(id)) continue
      if (!out.includes(id)) out.push(id)
      if (out.length >= MAX_SHARED_FILES) return out
    }
  }
  return out
}

/**
 * Whether this message needs the Graph round-trip at all: only in a channel or
 * group chat, only when Teams did not hand us the file directly, and only when
 * there is at least one shared-file placeholder to look up.
 */
export function needsSharedFileLookup(attachments: Attachment[], isGroup: boolean): boolean {
  if (!isGroup) return false
  const direct = attachments.some(
    att => att && att.contentType === 'application/vnd.microsoft.teams.file.download.info',
  )
  if (direct) return false
  return attachmentRefIds(attachments).length > 0
}

/**
 * Encode a sharing URL the way the Graph `shares` endpoint wants it: "u!" plus
 * the base64url (no "=" padding) of the URL bytes.
 */
export function encodeSharingUrl(url: string): string {
  return 'u!' + Buffer.from(url, 'utf8').toString('base64url')
}

/**
 * Host gate for shared-file content URLs. Only https, no userinfo, and a
 * *.sharepoint.com host — the Graph share link points there and nothing else is
 * worth fetching.
 */
export function isAllowedSharePointUrl(url: string): boolean {
  let u: URL
  try {
    u = new URL(url)
  } catch {
    return false
  }
  if (u.protocol !== 'https:') return false
  if (u.username || u.password) return false
  return u.hostname.toLowerCase().endsWith('.sharepoint.com')
}

/** Where the shared file lives, as far as the inbound activity can tell. */
export type SharedFileLookup =
  | {
      kind: 'channel'
      teamThreadId: string
      channelId: string
      messageId: string
      threadRootId?: string
      serviceUrl?: string
      aadGroupId?: string
      refIds: string[]
    }
  | { kind: 'chat'; chatId: string; messageId: string; refIds: string[] }

export interface SharedFilesDeps {
  /** App-only Microsoft Graph token. */
  getGraphToken: () => Promise<string>
  /** Bot Framework connector token. */
  getBotToken: () => Promise<string>
  /** Base storage directory, same one the 1:1 path uses. */
  baseDir: string
  /** Optional fetch override for tests. */
  fetchImpl?: typeof fetch
  /** Optional clock override for deterministic tests. Returns a Date. */
  now?: () => Date
  /** teamThreadId -> Azure AD group id, so the Bot Framework lookup happens once per team. */
  teamGroupIdCache?: Map<string, string>
}

/** Last path segment of a URL, decoded; "file" when there is nothing usable. */
function nameFromUrl(url: string): string {
  let u: URL
  try {
    u = new URL(url)
  } catch {
    return 'file'
  }
  const segments = u.pathname.split('/')
  const last = segments[segments.length - 1] ?? ''
  if (last === '') return 'file'
  try {
    return decodeURIComponent(last)
  } catch {
    return 'file'
  }
}

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

/**
 * The only redirect target we are willing to fetch: https, no userinfo, and it
 * must parse. Returns the Location value unchanged, or undefined when the
 * target fails the check.
 */
function httpsRedirectTarget(location: string | null): string | undefined {
  if (location === null) return undefined
  let u: URL
  try {
    u = new URL(location)
  } catch {
    return undefined
  }
  if (u.protocol !== 'https:' || u.username !== '' || u.password !== '') return undefined
  return location
}

async function run(
  lookup: SharedFileLookup,
  fromName: string,
  deps: SharedFilesDeps,
): Promise<{ annotations: string[]; results: AttachmentResult[] }> {
  const f = deps.fetchImpl ?? fetch
  const now = deps.now ?? (() => new Date())
  const n = lookup.refIds.length
  const countName = `${n} shared file(s)`

  const results: AttachmentResult[] = []
  const record = (r: AttachmentResult) => {
    logAttachmentResult(r)
    results.push(r)
  }
  const done = () => ({
    annotations: results.map(formatAnnotation).filter(s => s.length > 0),
    results,
  })

  // Every request gets its own timeout; none of them is allowed to hang a turn.
  const request = async (url: string, init: RequestInit): Promise<Response> => {
    const ctrl = new AbortController()
    const t = setTimeout(() => ctrl.abort(), DOWNLOAD_TIMEOUT_MS)
    try {
      return await f(url, { ...init, signal: ctrl.signal })
    } finally {
      clearTimeout(t)
    }
  }

  // Step A — the team's Azure AD group id. Graph addresses a channel through
  // its group, Bot Framework through the thread id, so this is a translation
  // rather than a permission check.
  let gid: string | undefined
  if (lookup.kind === 'channel') {
    gid =
      lookup.aadGroupId !== undefined && GUID_RE.test(lookup.aadGroupId)
        ? lookup.aadGroupId
        : deps.teamGroupIdCache?.get(lookup.teamThreadId)
    if (gid === undefined) {
      const serviceUrl = lookup.serviceUrl
      // The s2s token below must never leave the Bot Framework service, so a
      // missing or unknown service URL is a hard stop — before any token call.
      if (!serviceUrl || !isAllowedInlineImageHost(serviceUrl)) {
        record({
          kind: 'failed',
          originalName: countName,
          reason: 'could not resolve the team id (untrusted or missing service URL)',
        })
        return done()
      }
      const res = await request(
        `${serviceUrl.replace(/\/+$/, '')}/v3/teams/${encodeURIComponent(lookup.teamThreadId)}`,
        { headers: { Authorization: `Bearer ${await deps.getBotToken()}` } },
      )
      if (!res.ok) {
        record({
          kind: 'failed',
          originalName: countName,
          reason: `could not resolve the team id (HTTP ${res.status})`,
        })
        return done()
      }
      let body: unknown
      try {
        body = await res.json()
      } catch {
        body = undefined
      }
      const aadGroupId = isObject(body) ? body.aadGroupId : undefined
      if (typeof aadGroupId !== 'string' || !GUID_RE.test(aadGroupId)) {
        record({
          kind: 'failed',
          originalName: countName,
          reason: 'could not resolve the team id (no aadGroupId)',
        })
        return done()
      }
      gid = aadGroupId
      deps.teamGroupIdCache?.set(lookup.teamThreadId, gid)
    }
  }

  // Step B — the message itself, for the real attachment entries.
  let messageUrl: string
  if (lookup.kind === 'channel') {
    const base =
      `${GRAPH_V1}/teams/${encodeURIComponent(gid!)}` +
      `/channels/${encodeURIComponent(lookup.channelId)}/messages`
    const root = lookup.threadRootId
    messageUrl =
      typeof root === 'string' && root !== '' && root !== lookup.messageId
        ? `${base}/${encodeURIComponent(root)}/replies/${encodeURIComponent(lookup.messageId)}`
        : `${base}/${encodeURIComponent(lookup.messageId)}`
  } else {
    messageUrl =
      `${GRAPH_V1}/chats/${encodeURIComponent(lookup.chatId)}` +
      `/messages/${encodeURIComponent(lookup.messageId)}`
  }
  const msgRes = await request(messageUrl, {
    headers: { Authorization: `Bearer ${await deps.getGraphToken()}` },
  })
  if (!msgRes.ok) {
    // 401/403 here is almost always a missing resource-specific consent, so
    // say which one instead of leaving the operator to guess.
    const consent =
      lookup.kind === 'channel' ? 'ChannelMessage.Read.Group' : 'ChatMessage.Read.Chat'
    const suffix =
      msgRes.status === 401 || msgRes.status === 403
        ? `; the Teams app needs the ${consent} consent for this ${lookup.kind === 'channel' ? 'team' : 'chat'}`
        : ''
    record({
      kind: 'failed',
      originalName: countName,
      reason: `could not read the message from Microsoft Graph (HTTP ${msgRes.status})${suffix}`,
    })
    return done()
  }
  let msgBody: unknown
  try {
    msgBody = await msgRes.json()
  } catch {
    msgBody = undefined
  }
  const listed = isObject(msgBody) && Array.isArray(msgBody.attachments) ? msgBody.attachments : []
  const refs: { contentUrl: string; name: unknown }[] = []
  for (const el of listed) {
    if (!isObject(el)) continue
    if (el.contentType !== 'reference') continue
    if (typeof el.contentUrl !== 'string') continue
    if (typeof el.id !== 'string' || !lookup.refIds.includes(el.id)) continue
    refs.push({ contentUrl: el.contentUrl, name: el.name })
    if (refs.length >= MAX_SHARED_FILES) break
  }
  // A placeholder whose reference never turns up (a card, a loop component) is
  // not an error: there is simply nothing to download.
  if (refs.length === 0) return { annotations: [], results: [] }

  // Step C — download each reference, in the order the message lists them.
  for (const ref of refs) {
    const name =
      typeof ref.name === 'string' && ref.name.trim() !== '' ? ref.name : nameFromUrl(ref.contentUrl)
    try {
      if (!isAllowedSharePointUrl(ref.contentUrl)) {
        record({ kind: 'skipped', originalName: name, reason: 'not a SharePoint file link' })
        continue
      }
      const res = await request(
        `${GRAPH_V1}/shares/${encodeSharingUrl(ref.contentUrl)}/driveItem/content`,
        { headers: { Authorization: `Bearer ${await deps.getGraphToken()}` }, redirect: 'manual' },
      )
      let tempPath: string
      let sizeBytes: number
      if (REDIRECT_STATUSES.includes(res.status)) {
        const location = httpsRedirectTarget(res.headers.get('location'))
        if (location === undefined) {
          record({ kind: 'failed', originalName: name, reason: 'download redirect is not https' })
          continue
        }
        // The pre-authenticated SharePoint link carries its own token in the
        // query string — sending ours along would only leak it.
        ;({ tempPath, sizeBytes } = await downloadToTemp(location, {}, f))
      } else if (res.status === 200) {
        ;({ tempPath, sizeBytes } = await saveResponseToTemp(res))
      } else if (res.status === 401 || res.status === 403) {
        record({
          kind: 'skipped',
          originalName: name,
          reason:
            `the Teams app has no Microsoft Graph permission to read this file (HTTP ${res.status}); ` +
            'it needs the Files.Read.All application permission',
        })
        continue
      } else if (res.status === 404) {
        record({ kind: 'failed', originalName: name, reason: 'file not found in SharePoint (HTTP 404)' })
        continue
      } else {
        record({ kind: 'failed', originalName: name, reason: `HTTP ${res.status}` })
        continue
      }
      const when = now()
      const dateDir = join(deps.baseDir, londonDate(when))
      const prefix = `${londonTime(when)}-${senderSlug(fromName)}`
      mkdirSync(dateDir, { recursive: true, mode: 0o700 })
      const finalPath = pickPath(dateDir, prefix, name)
      moveFile(tempPath, finalPath)
      const dot = name.lastIndexOf('.')
      const rawExt = dot < 0 ? '' : name.slice(dot + 1).toLowerCase()
      const ext = /^[a-z0-9]{1,10}$/.test(rawExt) ? rawExt : ''
      record({
        kind: 'saved',
        path: finalPath,
        contentType: ext !== '' ? `application/${ext}` : 'application/octet-stream',
        sizeBytes,
        originalName: name,
      })
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err)
      if (msg.startsWith('exceeds ')) {
        record({ kind: 'skipped', originalName: name, reason: msg })
      } else {
        record({ kind: 'failed', originalName: name, reason: msg })
      }
    }
  }
  return done()
}

/**
 * Fetch the files shared with a channel / group-chat message. Never throws:
 * an unexpected error becomes one failed result, so the message still reaches
 * the agent with a note about what went wrong.
 */
export async function fetchSharedFiles(
  lookup: SharedFileLookup,
  fromName: string,
  deps: SharedFilesDeps,
): Promise<{ annotations: string[]; results: AttachmentResult[] }> {
  try {
    return await run(lookup, fromName, deps)
  } catch (err) {
    const reason = err instanceof Error ? err.message : String(err)
    const failed: AttachmentResult = { kind: 'failed', originalName: 'shared file', reason }
    logAttachmentResult(failed)
    return { annotations: [formatAnnotation(failed)], results: [failed] }
  }
}
