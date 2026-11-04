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
 *   3. the file's Graph drive item, then the bytes themselves from
 *      SharePoint/OneDrive.
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
 *   - The Graph token (getGraphToken) only ever goes to graph.microsoft.com:
 *     every request that carries it is sent with `redirect: 'error'`, so no
 *     redirect can hand it to a second host.
 *   - A contentUrl that is not a *.sharepoint.com https link is never fetched.
 *   - Ownership: Files.Read.All reads the WHOLE tenant, so without a check a
 *     sender could make the bot fetch any tenant file by pasting a link. A file
 *     is only downloaded when the conversation already shares it — in a channel
 *     the file must sit inside the channel's files folder (the one behind the
 *     Files tab, at any depth up to MAX_FOLDER_DEPTH levels), in a group chat the
 *     drive's owner must be the sender. Everything else gets a skip line.
 *   - The download is followed manually: a hop target must be a
 *     *.sharepoint.com https link without userinfo, there are at most three
 *     hops, and the hop is fetched WITHOUT any bearer token (the
 *     pre-authenticated SharePoint link carries its own in the query string).
 *   - Every request runs under one timer that only stops once the body has been
 *     read, so a slow body cannot outlive the timeout.
 */

import { mkdirSync, unlinkSync } from 'fs'
import { join } from 'path'
import type { Attachment } from 'botframework-schema'
import {
  DOWNLOAD_TIMEOUT_MS,
  MAX_ATTACHMENT_BYTES,
  type AttachmentResult,
  formatAnnotation,
  formatBytes,
  isAllowedInlineImageHost,
  localDate,
  localTime,
  logAttachmentResult,
  moveFile,
  pickPath,
  saveResponseToTemp,
  senderSlug,
} from './attachments.js'

export const GRAPH_V1 = 'https://graph.microsoft.com/v1.0'

/** At most this many files are pulled out of one message. */
export const MAX_SHARED_FILES = 10

/**
 * How many parent folders are walked up from a file to find the channel's files
 * folder: at most this many ancestry requests per file.
 */
export const MAX_FOLDER_DEPTH = 8

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
// The id is only ever COMPARED against the Graph message's own attachment ids
// — it never goes into a URL — so anything without whitespace, quotes, angle
// brackets or control characters is acceptable, bounded in length.
const REF_ID_RE = /^[^\s"'<>\u0000-\u001f\u007f]{1,256}$/

// The placeholder Teams puts in the HTML part for every shared file.
const ATTACHMENT_TAG_RE = /<attachment\b[^>]*\bid\s*=\s*["']([^"']*)["']/gi

const REDIRECT_STATUSES = [301, 302, 303, 307, 308]

// At most this many redirects are followed by hand on the download.
const MAX_DOWNLOAD_HOPS = 3

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
  | {
      kind: 'chat'
      chatId: string
      messageId: string
      /** The sender's AAD object id, lowercased — the OneDrive owner check. */
      senderAadObjectId?: string
      refIds: string[]
    }

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
  /** Per-request timeout, one timer per request. Defaults to DOWNLOAD_TIMEOUT_MS. */
  timeoutMs?: number
  /** teamThreadId -> Azure AD group id, so the Bot Framework lookup happens once per team. */
  teamGroupIdCache?: Map<string, string>
  /** `${aadGroupId}/${channelId}` -> the drive id and item id of the channel's files folder. */
  channelDriveCache?: Map<string, { driveId: string; folderId: string }>
  /** drive id -> the AAD object id of the OneDrive that owns the drive. */
  driveOwnerCache?: Map<string, string>
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
 * A file name as the agent is allowed to see it: control characters, DEL and
 * the Unicode line / paragraph separators become spaces (a line break would
 * let a name forge a line of its own after the annotation), and the name is
 * capped. The raw name still goes to pickPath, which sanitises the file name
 * on disk.
 */
const shown = (s: string) => s.replace(/[\u0000-\u001f\u007f\u2028\u2029]/g, ' ').slice(0, 200)

/** A response body as JSON, or undefined when it does not parse. */
async function readJson(res: Response): Promise<unknown> {
  try {
    return await res.json()
  } catch {
    return undefined
  }
}

/**
 * The mapping the Graph metadata and content requests share: 401/403 is a
 * missing Files.Read.All (worth saying, it is an operator fix), 404 is a file
 * that is gone, anything else is just the status.
 */
function graphFileError(status: number): { kind: 'skipped' | 'failed'; reason: string } {
  if (status === 401 || status === 403) {
    return {
      kind: 'skipped',
      reason:
        `the Teams app has no Microsoft Graph permission to read this file (HTTP ${status}); ` +
        'it needs the Files.Read.All application permission',
    }
  }
  if (status === 404) return { kind: 'failed', reason: 'file not found in SharePoint (HTTP 404)' }
  return { kind: 'failed', reason: `HTTP ${status}` }
}

/** What one hop of the download left us with: the bytes, the next hop, or a stop. */
type DownloadOutcome =
  | { tempPath: string; sizeBytes: number }
  | { next: string }
  | { fail: string; kind: 'skipped' | 'failed' }

/** What the Graph drive item gave us: where the file lives, or why it did not. */
type DriveItem =
  | { id: string; driveId: string; parentId?: string; size?: number }
  | { error: { kind: 'skipped' | 'failed'; reason: string } }

async function run(
  lookup: SharedFileLookup,
  fromName: string,
  deps: SharedFilesDeps,
): Promise<{ annotations: string[]; results: AttachmentResult[] }> {
  const f = deps.fetchImpl ?? fetch
  const now = deps.now ?? (() => new Date())
  const n = lookup.refIds.length
  const countName = shown(`${n} shared file(s)`)

  const results: AttachmentResult[] = []
  const record = (r: AttachmentResult) => {
    logAttachmentResult(r)
    results.push(r)
  }
  const done = () => ({
    annotations: results.map(formatAnnotation).filter(s => s.length > 0),
    results,
  })

  // One request, one timer — and the timer only stops once the body has been
  // read, because `use` is awaited inside the try. Reading the body outside it
  // would let a slow body outlive the timeout.
  const call = async <T>(
    url: string,
    init: RequestInit,
    use: (res: Response) => Promise<T>,
  ): Promise<T> => {
    const ctrl = new AbortController()
    const t = setTimeout(() => ctrl.abort(), deps.timeoutMs ?? DOWNLOAD_TIMEOUT_MS)
    try {
      const res = await f(url, { ...init, signal: ctrl.signal })
      return await use(res)
    } finally {
      clearTimeout(t)
    }
  }
  const graphHeaders = async (): Promise<Record<string, string>> => ({
    Authorization: `Bearer ${await deps.getGraphToken()}`,
  })

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
      const team = await call(
        `${serviceUrl.replace(/\/+$/, '')}/v3/teams/${encodeURIComponent(lookup.teamThreadId)}`,
        { headers: { Authorization: `Bearer ${await deps.getBotToken()}` }, redirect: 'error' },
        async res => ({ ok: res.ok, status: res.status, body: await readJson(res) }),
      )
      if (!team.ok) {
        record({
          kind: 'failed',
          originalName: countName,
          reason: `could not resolve the team id (HTTP ${team.status})`,
        })
        return done()
      }
      const aadGroupId = isObject(team.body) ? team.body.aadGroupId : undefined
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
  const msg = await call(
    messageUrl,
    { headers: await graphHeaders(), redirect: 'error' },
    async res => ({ ok: res.ok, status: res.status, body: await readJson(res) }),
  )
  if (!msg.ok) {
    // 401/403 here is almost always a missing resource-specific consent, so
    // say which one instead of leaving the operator to guess.
    const consent =
      lookup.kind === 'channel' ? 'ChannelMessage.Read.Group' : 'ChatMessage.Read.Chat'
    const denied = msg.status === 401 || msg.status === 403
    const suffix = denied
      ? `; the Teams app needs the ${consent} consent for this ${lookup.kind === 'channel' ? 'team' : 'chat'}`
      : ''
    record({
      // A missing permission is an explicit skip, anything else a failure.
      kind: denied ? 'skipped' : 'failed',
      originalName: countName,
      reason: `could not read the message from Microsoft Graph (HTTP ${msg.status})${suffix}`,
    })
    return done()
  }
  const listed = isObject(msg.body) && Array.isArray(msg.body.attachments) ? msg.body.attachments : []
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
  // not an error: there is simply nothing to download. One stderr line, so an
  // operator can tell "nothing to fetch" from "Graph came up empty".
  if (refs.length === 0) {
    if (n > 0) {
      process.stderr.write(
        `teams channel: shared file lookup found no file reference for ${n} placeholder(s) in message ${shown(lookup.messageId)}\n`,
      )
    }
    return { annotations: [], results: [] }
  }

  // The channel's files folder (its drive and its item id), read once per
  // message (and once per process through the cache): every file a channel
  // shares lives inside it.
  let channelFolder =
    lookup.kind === 'channel' ? deps.channelDriveCache?.get(`${gid}/${lookup.channelId}`) : undefined

  // Step C5 — the bytes. Graph answers the content request with a redirect to a
  // pre-authenticated SharePoint link, which is followed by hand: the hop is
  // fetched WITHOUT our bearer token (the link carries its own in the query
  // string), only to an allowed SharePoint address, and at most three times.
  const download = async (
    driveId: string,
    itemId: string,
  ): Promise<{ tempPath: string; sizeBytes: number } | { outcome: { kind: 'skipped' | 'failed'; reason: string } }> => {
    let current =
      `${GRAPH_V1}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`
    for (let hop = 0; ; hop++) {
      const out = await call(
        current,
        hop === 0
          ? { headers: await graphHeaders(), redirect: 'manual' }
          : { redirect: 'manual' },
        async (res): Promise<DownloadOutcome> => {
          if (res.status === 200) return await saveResponseToTemp(res)
          if (REDIRECT_STATUSES.includes(res.status)) {
            const location = res.headers.get('location')
            if (location === null) return { fail: 'download redirect without a location', kind: 'failed' as const }
            if (hop >= MAX_DOWNLOAD_HOPS) return { fail: 'too many download redirects', kind: 'failed' as const }
            const next = new URL(location, current).href
            if (!isAllowedSharePointUrl(next)) {
              return { fail: 'download redirect target is not allowed', kind: 'failed' as const }
            }
            return { next }
          }
          // The Graph content request gets the Graph permission wording; a hop
          // is fetched without any token of ours, so its answer is just a status.
          if (hop > 0) return { fail: `HTTP ${res.status}`, kind: 'failed' as const }
          const err = graphFileError(res.status)
          return { fail: err.reason, kind: err.kind }
        },
      )
      if ('next' in out) {
        current = out.next
        continue
      }
      if ('fail' in out) return { outcome: { kind: out.kind, reason: out.fail } }
      return out
    }
  }

  // Is the file inside the channel's files folder? Walks up from the file's
  // parent folder, one ancestry request per level (so at most MAX_FOLDER_DEPTH
  // per file, nothing cached), and stops as soon as it meets the folder, runs
  // out of parents, leaves the file's drive, or gets deeper than the limit.
  const insideFolder = async (
    driveId: string,
    parentId: string | undefined,
    folderId: string,
  ): Promise<{ inside: boolean } | { error: { kind: 'skipped' | 'failed'; reason: string } }> => {
    let cur = parentId
    for (let depth = 0; ; depth++) {
      if (cur === undefined) return { inside: false }
      if (cur === folderId) return { inside: true }
      if (depth >= MAX_FOLDER_DEPTH) return { inside: false }
      const step = await call(
        `${GRAPH_V1}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(cur)}?$select=id,parentReference`,
        { headers: await graphHeaders(), redirect: 'error' },
        async (res): Promise<{ next: string | undefined } | { error: { kind: 'skipped' | 'failed'; reason: string } }> => {
          if (!res.ok) return { error: graphFileError(res.status) }
          const body = await readJson(res)
          const parent = isObject(body) && isObject(body.parentReference) ? body.parentReference : undefined
          // A parent on another drive is not a step up inside this one.
          const next =
            parent !== undefined &&
            parent.driveId === driveId &&
            typeof parent.id === 'string' &&
            parent.id !== ''
              ? parent.id
              : undefined
          return { next }
        },
      )
      if ('error' in step) return step
      cur = step.next
    }
  }

  // Step C — one reference at a time, in the order the message lists them.
  for (const [index, ref] of refs.entries()) {
    const name =
      typeof ref.name === 'string' && ref.name.trim() !== '' ? ref.name : nameFromUrl(ref.contentUrl)
    const label = shown(name)
    let tempPath: string | undefined
    try {
      // C1 — a link that is not SharePoint is never fetched at all.
      if (!isAllowedSharePointUrl(ref.contentUrl)) {
        record({ kind: 'skipped', originalName: label, reason: 'not a SharePoint file link' })
        continue
      }

      // C2 — the drive item: which drive, which item, how big.
      const meta = await call(
        `${GRAPH_V1}/shares/${encodeSharingUrl(ref.contentUrl)}/driveItem?$select=id,name,size,parentReference`,
        { headers: await graphHeaders(), redirect: 'error' },
        async (res): Promise<DriveItem> => {
          if (!res.ok) return { error: graphFileError(res.status) }
          const body = await readJson(res)
          const b = isObject(body) ? body : undefined
          const parent = b !== undefined && isObject(b.parentReference) ? b.parentReference : undefined
          const id = typeof b?.id === 'string' && b.id !== '' ? b.id : undefined
          const driveId =
            typeof parent?.driveId === 'string' && parent.driveId !== '' ? parent.driveId : undefined
          const parentId = typeof parent?.id === 'string' && parent.id !== '' ? parent.id : undefined
          if (id === undefined || driveId === undefined) {
            return {
              error: {
                kind: 'failed' as const,
                reason: 'Microsoft Graph returned no drive item for this file',
              },
            }
          }
          const size = typeof b?.size === 'number' && Number.isFinite(b.size) ? b.size : undefined
          return { id, driveId, parentId, size }
        },
      )
      if ('error' in meta) {
        record({ kind: meta.error.kind, originalName: label, reason: meta.error.reason })
        continue
      }

      // C3 — ownership. Files.Read.All would let a sender pull any tenant file
      // through the bot, so a file is only taken when the conversation already
      // shares it: inside the channel's files folder (at any depth up to
      // MAX_FOLDER_DEPTH) in a channel, the sender's OneDrive in a group chat.
      if (lookup.kind === 'channel') {
        if (channelFolder === undefined) {
          const folder = await call(
            `${GRAPH_V1}/teams/${encodeURIComponent(gid!)}` +
              `/channels/${encodeURIComponent(lookup.channelId)}/filesFolder?$select=id,parentReference`,
            { headers: await graphHeaders(), redirect: 'error' },
            async res => {
              const body = await readJson(res)
              const parent =
                isObject(body) && isObject(body.parentReference) ? body.parentReference : undefined
              const driveId =
                typeof parent?.driveId === 'string' && parent.driveId !== '' ? parent.driveId : undefined
              const folderId =
                isObject(body) && typeof body.id === 'string' && body.id !== '' ? body.id : undefined
              return { ok: res.ok, status: res.status, driveId, folderId }
            },
          )
          if (!folder.ok || folder.driveId === undefined || folder.folderId === undefined) {
            // One line for the whole message, and nothing else is fetched: the
            // channel's folder is the same for every reference in it, so the
            // line names every file that is left (this one included).
            const remaining = refs.length - index
            const denied = !folder.ok && (folder.status === 401 || folder.status === 403)
            record({
              kind: denied ? 'skipped' : 'failed',
              originalName: remaining === 1 ? label : shown(`${remaining} shared file(s)`),
              reason: folder.ok
                ? "could not read the channel's files folder (no drive id or folder id)"
                : denied
                  ? `the Teams app has no Microsoft Graph permission to read this channel's files folder (HTTP ${folder.status})`
                  : `could not read the channel's files folder (HTTP ${folder.status})`,
            })
            return done()
          }
          channelFolder = { driveId: folder.driveId, folderId: folder.folderId }
          deps.channelDriveCache?.set(`${gid}/${lookup.channelId}`, channelFolder)
        }
        const notInside = {
          kind: 'skipped' as const,
          originalName: label,
          reason: "the file is not stored in this channel's files folder, so it is not downloaded",
        }
        // First the drive, which costs nothing; then the walk up the parent
        // folders to the channel's files folder, one request per level.
        if (meta.driveId !== channelFolder.driveId) {
          record(notInside)
          continue
        }
        const inside = await insideFolder(meta.driveId, meta.parentId, channelFolder.folderId)
        if ('error' in inside) {
          record({ kind: inside.error.kind, originalName: label, reason: inside.error.reason })
          continue
        }
        if (!inside.inside) {
          record(notInside)
          continue
        }
      } else {
        const sender = lookup.senderAadObjectId
        if (sender === undefined) {
          record({
            kind: 'skipped',
            originalName: label,
            reason: "the file is not in the sender's OneDrive, so it is not downloaded",
          })
          continue
        }
        let owner = deps.driveOwnerCache?.get(meta.driveId)
        if (owner === undefined) {
          const drive = await call(
            `${GRAPH_V1}/drives/${encodeURIComponent(meta.driveId)}?$select=owner`,
            { headers: await graphHeaders(), redirect: 'error' },
            async res => {
              const body = await readJson(res)
              const o = isObject(body) && isObject(body.owner) ? body.owner : undefined
              const user = o !== undefined && isObject(o.user) ? o.user : undefined
              const id = typeof user?.id === 'string' && user.id !== '' ? user.id : undefined
              return { ok: res.ok, status: res.status, id }
            },
          )
          if (!drive.ok) {
            const denied = drive.status === 401 || drive.status === 403
            record({
              kind: denied ? 'skipped' : 'failed',
              originalName: label,
              reason: denied
                ? `the Teams app has no Microsoft Graph permission to read the file's drive owner (HTTP ${drive.status})`
                : `could not read the file's drive owner (HTTP ${drive.status})`,
            })
            continue
          }
          if (drive.id === undefined) {
            // A drive with no user owner (a team library is owned by a group) is
            // not the sender's OneDrive. Nothing is cached for it.
            record({
              kind: 'skipped',
              originalName: label,
              reason: "the file is not in the sender's OneDrive, so it is not downloaded",
            })
            continue
          }
          owner = drive.id
          deps.driveOwnerCache?.set(meta.driveId, owner)
        }
        if (owner.toLowerCase() !== sender.toLowerCase()) {
          record({
            kind: 'skipped',
            originalName: label,
            reason: "the file is not in the sender's OneDrive, so it is not downloaded",
          })
          continue
        }
      }

      // C4 — a size we already know is over the cap is not worth a download.
      if (meta.size !== undefined && meta.size > MAX_ATTACHMENT_BYTES) {
        record({
          kind: 'skipped',
          originalName: label,
          reason: `exceeds ${formatBytes(MAX_ATTACHMENT_BYTES)} limit`,
        })
        continue
      }

      // C5 — the bytes themselves (see `download` above for the redirect rules).
      const bytes = await download(meta.driveId, meta.id)
      if ('outcome' in bytes) {
        record({ kind: bytes.outcome.kind, originalName: label, reason: bytes.outcome.reason })
        continue
      }
      tempPath = bytes.tempPath

      // C6 — into the dated directory, under the sender's prefix.
      const when = now()
      const dateDir = join(deps.baseDir, localDate(when))
      const prefix = `${localTime(when)}-${senderSlug(fromName)}`
      mkdirSync(dateDir, { recursive: true, mode: 0o700 })
      const finalPath = pickPath(dateDir, prefix, name)
      moveFile(tempPath, finalPath)
      const dot = name.lastIndexOf('.')
      const rawExt = dot < 0 ? '' : name.slice(dot + 1).toLowerCase()
      const ext = /^[a-z0-9]{1,10}$/.test(rawExt) ? rawExt : ''
      tempPath = undefined
      record({
        kind: 'saved',
        path: finalPath,
        contentType: ext !== '' ? `application/${ext}` : 'application/octet-stream',
        sizeBytes: bytes.sizeBytes,
        originalName: label,
      })
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err)
      // A temp file that never made it to its final name is not left behind.
      if (tempPath !== undefined) {
        try {
          unlinkSync(tempPath)
        } catch {}
      }
      if (msg.startsWith('exceeds ')) {
        record({ kind: 'skipped', originalName: label, reason: msg })
      } else {
        record({ kind: 'failed', originalName: label, reason: msg })
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
