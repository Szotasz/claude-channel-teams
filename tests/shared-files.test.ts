/**
 * Channel / group-chat shared files (src/teams/sharedFiles.ts).
 *
 * In a shared conversation Teams does not attach the file to the activity, only
 * an <attachment id="…"> placeholder in an HTML part, so the plugin has to read
 * the message through Microsoft Graph and pull the file out of SharePoint. These
 * tests script every HTTP hop with a recording fake fetch and assert the exact
 * diagnostics the agent ends up seeing.
 *
 * Fixtures are synthetic: a fictional contoso.sharepoint.com tenant, a
 * non-existent team thread id and a sender name that is nobody's.
 */

import { describe, test, expect, beforeEach, afterEach, spyOn } from 'bun:test'
import { mkdtempSync, rmSync, readFileSync, readdirSync, statSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join, dirname } from 'path'
import type { Attachment } from 'botframework-schema'

import {
  GRAPH_V1,
  MAX_SHARED_FILES,
  attachmentRefIds,
  encodeSharingUrl,
  fetchSharedFiles,
  isAllowedSharePointUrl,
  needsSharedFileLookup,
  type SharedFileLookup,
  type SharedFilesDeps,
} from '../src/teams/sharedFiles'
import {
  MAX_ATTACHMENT_BYTES,
  formatBytes,
  logAttachmentResult,
  londonDate,
  londonTime,
  saveResponseToTemp,
  senderSlug,
} from '../src/teams/attachments'

const TEAM = '19:team0001@thread.tacv2'
const CHAN = '19:chan0001@thread.tacv2'
const CHAT = '19:chat0001@thread.v2'
const GID = '00000000-0000-4000-8000-0000000000aa'
const SERVICE = 'https://smba.trafficmanager.net/emea/00000000-0000-4000-8000-000000000002/'
const SENDER = 'Test User'
const SENDER_AAD = '00000000-0000-4000-8000-0000000000BB'
const NOW = new Date('2026-07-01T10:30:00Z')
const FILE_URL = 'https://contoso.sharepoint.com/sites/x/Shared Documents/General/report.xlsx'
const BODY = 'xlsx-bytes'
const DOWNLOAD_URL = 'https://contoso.sharepoint.com/sites/x/_layouts/15/download.aspx?file=report.xlsx'
const DRIVE = 'b!drive0001'
const OTHER_DRIVE = 'b!drive0002'
const ITEM = '01FLitem!0001'

function html(content: string): Attachment {
  return { contentType: 'text/html', content } as Attachment
}

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  })
}

/** A reference attachment as Microsoft Graph describes it. */
function graphRef(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return { id: 'ref-1', contentType: 'reference', contentUrl: FILE_URL, name: 'report.xlsx', ...overrides }
}

/** The Graph message carrying those references. */
function graphMessage(attachments: unknown[]): Response {
  return jsonResponse({ attachments })
}

/** The drive item behind the sharing link. */
function graphMetadata(overrides: Record<string, unknown> = {}): Response {
  return jsonResponse({
    id: ITEM,
    name: 'report.xlsx',
    size: BODY.length,
    parentReference: { driveId: DRIVE },
    ...overrides,
  })
}

const metaUrl = () => `${GRAPH_V1}/shares/${encodeSharingUrl(FILE_URL)}/driveItem?$select=id,name,size,parentReference`
const contentUrl = (drive: string = DRIVE, item: string = ITEM) =>
  `${GRAPH_V1}/drives/${encodeURIComponent(drive)}/items/${encodeURIComponent(item)}/content`
const driveUrl = (drive: string = DRIVE) => `${GRAPH_V1}/drives/${encodeURIComponent(drive)}?$select=owner`
const filesFolderUrl = () =>
  `${GRAPH_V1}/teams/${GID}/channels/${encodeURIComponent(CHAN)}/filesFolder?$select=parentReference`

type Call = { url: string; method: string; headers: Record<string, string>; redirect?: string; signal?: AbortSignal }

function plainHeaders(h: unknown): Record<string, string> {
  if (h === undefined || h === null) return {}
  if (h instanceof Headers) return Object.fromEntries(h.entries())
  if (typeof h === 'object') {
    const out: Record<string, string> = {}
    for (const [k, v] of Object.entries(h as Record<string, unknown>)) out[k] = String(v)
    return out
  }
  return {}
}

/** fetch that records every call and answers from the first matching route. */
function fakeFetch(routes: { test: (url: string) => boolean; reply: (call: Call) => Response }[]) {
  const calls: Call[] = []
  const fn = (async (input: any, init: any = {}) => {
    const call: Call = {
      url: typeof input === 'string' ? input : input.toString(),
      method: init.method ?? 'GET',
      headers: plainHeaders(init.headers),
      redirect: init.redirect,
      signal: init.signal,
    }
    calls.push(call)
    for (const r of routes) if (r.test(call.url)) return r.reply(call)
    throw new Error(`unexpected request: ${call.url}`)
  }) as typeof fetch
  return { fn, calls }
}

/** The usual happy-path routing: bot team lookup, Graph message, drive item,
 *  the channel's files folder (or the drive's owner), the content, the hop. */
function standardRoutes(overrides: {
  message?: () => Response
  metadata?: () => Response
  filesFolder?: () => Response
  driveOwner?: () => Response
  content?: () => Response
  download?: () => Response
  team?: () => Response
} = {}) {
  return [
    {
      test: (u: string) => u.includes('/v3/teams/'),
      reply: () => overrides.team?.() ?? jsonResponse({ aadGroupId: GID }),
    },
    {
      test: (u: string) => u.includes('/messages/') || u.includes('/replies/'),
      reply: () => overrides.message?.() ?? graphMessage([graphRef()]),
    },
    {
      test: (u: string) => u.startsWith(`${GRAPH_V1}/shares/`),
      reply: () => overrides.metadata?.() ?? graphMetadata(),
    },
    {
      test: (u: string) => u.includes('/filesFolder'),
      reply: () => overrides.filesFolder?.() ?? jsonResponse({ parentReference: { driveId: DRIVE } }),
    },
    {
      test: (u: string) => u === driveUrl(),
      reply: () => overrides.driveOwner?.() ?? jsonResponse({ owner: { user: { id: SENDER_AAD } } }),
    },
    {
      test: (u: string) => u === contentUrl(),
      reply: () =>
        overrides.content?.() ??
        new Response(null, { status: 302, headers: { location: DOWNLOAD_URL } }),
    },
    {
      test: (u: string) => u === DOWNLOAD_URL,
      reply: () => overrides.download?.() ?? new Response(BODY),
    },
  ]
}

function channelLookup(overrides: Record<string, unknown> = {}): SharedFileLookup {
  return {
    kind: 'channel',
    teamThreadId: TEAM,
    channelId: CHAN,
    messageId: 'MSG',
    threadRootId: 'ROOT',
    serviceUrl: SERVICE,
    refIds: ['ref-1'],
    ...overrides,
  } as SharedFileLookup
}

function chatLookup(overrides: Record<string, unknown> = {}): SharedFileLookup {
  return {
    kind: 'chat',
    chatId: CHAT,
    messageId: 'MSG',
    senderAadObjectId: SENDER_AAD,
    refIds: ['ref-1'],
    ...overrides,
  } as SharedFileLookup
}

function deps(
  fetchImpl: typeof fetch,
  baseDir: string,
  overrides: Partial<SharedFilesDeps> & { getBotToken?: () => Promise<string> } = {},
): SharedFilesDeps {
  return {
    getGraphToken: async () => 'GRAPH',
    getBotToken: async () => 'BOT',
    baseDir,
    fetchImpl,
    now: () => NOW,
    ...overrides,
  }
}

/** Everything the plugin leaves in the OS temp directory under its own prefix. */
const attachTemps = () => readdirSync(tmpdir()).filter(n => n.startsWith('cct-attach-'))

/** Run fn with process.stderr.write captured, so the exact lines can be checked. */
async function captureStderr<T>(fn: () => Promise<T>): Promise<{ value: T; lines: string[] }> {
  const lines: string[] = []
  const orig = process.stderr.write
  process.stderr.write = ((chunk: unknown) => {
    lines.push(String(chunk))
    return true
  }) as typeof process.stderr.write
  try {
    return { value: await fn(), lines }
  } finally {
    process.stderr.write = orig
  }
}

describe('sharedFiles: attachmentRefIds', () => {
  test('S1 collects ids from double- and single-quoted tags, drops dupes and unsafe ids', () => {
    const part = html(
      '<p>see this</p>' +
        '<attachment id="ref-1"></attachment>' +
        "<attachment id='ref-2'></attachment>" +
        '<attachment id="ref-1"></attachment>' +
        '<attachment id="has/slash"></attachment>' +
        '<attachment id="has space"></attachment>' +
        '<attachment id=\'has"quote\'></attachment>',
    )
    const ids = attachmentRefIds([part])
    // A slash is fine now: the id is only ever COMPARED against the ids of the
    // Graph message's own attachments, never put in a URL. A quote ends the
    // attribute value, so no id carrying one can survive either.
    expect(ids).toEqual(['ref-1', 'ref-2', 'has/slash', 'has']) // 'has"quote' ends at the quote
    expect(ids.every(id => !/[\s"'<>]/.test(id))).toBe(true)
    expect(ids).not.toContain('has space')
    // A non-html attachment is not scanned at all.
    expect(attachmentRefIds([{ contentType: 'application/vnd.microsoft.card.adaptive', content: '{}' } as Attachment])).toEqual([])
  })

  test('S1b caps the list at MAX_SHARED_FILES', () => {
    const content = Array.from({ length: MAX_SHARED_FILES + 4 }, (_, i) => `<attachment id="ref-${i}"></attachment>`).join('')
    expect(attachmentRefIds([html(content)])).toHaveLength(MAX_SHARED_FILES)
  })
})

describe('sharedFiles: needsSharedFileLookup', () => {
  test('S2 only a group message with a placeholder and no direct file needs the lookup', () => {
    const part = html('<p>see this</p><attachment id="ref-1"></attachment>')
    expect(needsSharedFileLookup([part], false)).toBe(false)
    expect(needsSharedFileLookup([part, { contentType: 'application/vnd.microsoft.teams.file.download.info', content: {} } as Attachment], true)).toBe(false)
    expect(needsSharedFileLookup([html('<p>no tag here</p>')], true)).toBe(false)
    expect(needsSharedFileLookup([part], true)).toBe(true)
  })
})

describe('sharedFiles: encodeSharingUrl / isAllowedSharePointUrl', () => {
  test('S3 encodes as u! + base64url with no padding', () => {
    const url = 'https://contoso.sharepoint.com/sites/x/Shared Documents/General/a b.xlsx'
    const encoded = encodeSharingUrl(url)
    expect(encoded).toBe('u!' + Buffer.from(url, 'utf8').toString('base64url'))
    expect(encoded.includes('=')).toBe(false)
  })

  test('S3b only https *.sharepoint.com without userinfo is allowed', () => {
    expect(isAllowedSharePointUrl(FILE_URL)).toBe(true)
    expect(isAllowedSharePointUrl('http://contoso.sharepoint.com/x.xlsx')).toBe(false)
    expect(isAllowedSharePointUrl('https://contoso.sharepoint.com.evil.example.com/x.xlsx')).toBe(false)
    expect(isAllowedSharePointUrl('https://user:pw@contoso.sharepoint.com/x.xlsx')).toBe(false)
    expect(isAllowedSharePointUrl('https://evil.example.com/x.xlsx')).toBe(false)
    expect(isAllowedSharePointUrl('not a url')).toBe(false)
  })
})

describe('sharedFiles: constants', () => {
  test('X1 GRAPH_V1 and MAX_SHARED_FILES keep their literal values', () => {
    expect(GRAPH_V1).toBe('https://graph.microsoft.com/v1.0')
    expect(MAX_SHARED_FILES).toBe(10)
  })
})

describe('sharedFiles: fetchSharedFiles', () => {
  let dir: string

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cct-shared-files-'))
  })
  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('S4 channel reply, happy path with redirect', async () => {
    const f = fakeFetch(standardRoutes())
    const { annotations, results } = await fetchSharedFiles(channelLookup(), SENDER, deps(f.fn, dir))
    expect(f.calls.map(c => c.url)).toEqual([
      `${SERVICE.replace(/\/+$/, '')}/v3/teams/${encodeURIComponent(TEAM)}`,
      `${GRAPH_V1}/teams/${GID}/channels/${encodeURIComponent(CHAN)}/messages/ROOT/replies/MSG`,
      metaUrl(),
      filesFolderUrl(),
      contentUrl(),
      DOWNLOAD_URL,
    ])

    expect(f.calls[0].headers.Authorization).toBe('Bearer BOT')

    expect(f.calls[1].headers.Authorization).toBe('Bearer GRAPH')

    expect(f.calls[2].headers.Authorization).toBe('Bearer GRAPH')

    expect(f.calls[3].headers.Authorization).toBe('Bearer GRAPH')

    expect(f.calls[4].redirect).toBe('manual')
    expect(f.calls[4].headers.Authorization).toBe('Bearer GRAPH')

    expect(f.calls[5].url).toBe(DOWNLOAD_URL)
    expect(Object.keys(f.calls[5].headers).some(k => k.toLowerCase() === 'authorization')).toBe(false)

    expect(results.length).toBe(1)
    const saved = results[0]
    expect(saved.kind).toBe('saved')
    if (saved.kind !== 'saved') return
    expect(saved.path.startsWith(dir)).toBe(true)
    expect(saved.contentType).toBe('application/xlsx')
    expect(saved.originalName).toBe('report.xlsx')
    expect(readFileSync(saved.path, 'utf8')).toBe(BODY)
    expect(annotations).toEqual([`[Attached: ${saved.path} (application/xlsx, ${formatBytes(BODY.length)})]`])
  })

  test('X2 every token-bearing request is redirect error, the two downloads are manual', async () => {
    const f = fakeFetch(standardRoutes())
    await fetchSharedFiles(channelLookup(), SENDER, deps(f.fn, dir))
    // connector, message, drive item, files folder
    for (const i of [0, 1, 2, 3]) {
      expect(f.calls[i].redirect).toBe('error')
      expect(f.calls[i].headers.Authorization).toBe(i === 0 ? 'Bearer BOT' : 'Bearer GRAPH')
    }
    // content, hop
    for (const i of [4, 5]) {
      expect(f.calls[i].redirect).toBe('manual')
    }
    // the hop carries no token at all
    expect(Object.keys(f.calls[5].headers).some(k => k.toLowerCase() === 'authorization')).toBe(false)
  })

  test('S5 a shared teamGroupIdCache means the bot endpoint is called once for two lookups', async () => {
    const f = fakeFetch(standardRoutes())
    const cache = new Map<string, string>()
    const lookup = channelLookup()
    await fetchSharedFiles(lookup, SENDER, deps(f.fn, dir, { teamGroupIdCache: cache }))
    await fetchSharedFiles(lookup, SENDER, deps(f.fn, dir, { teamGroupIdCache: cache }))
    expect(f.calls.filter(c => c.url.includes('/v3/teams/')).length).toBe(1)
    expect(cache.get(TEAM)).toBe(GID)
  })

  test('S6 a lookup with aadGroupId never calls the bot endpoint or getBotToken', async () => {
    const f = fakeFetch(standardRoutes())
    let botTokens = 0
    const { results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir, { getBotToken: async () => { botTokens += 1; return 'BOT' } }),
    )
    expect(botTokens).toBe(0)
    expect(f.calls.some(c => c.url.includes('/v3/teams/'))).toBe(false)
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('saved')
  })

  test('S7 a root message and threadRootId === messageId both read the message, not a reply', async () => {
    for (const override of [{ threadRootId: undefined }, { threadRootId: 'MSG' }]) {
      const f = fakeFetch(standardRoutes())
      await fetchSharedFiles(channelLookup(override), SENDER, deps(f.fn, dir))
      const message = f.calls.find(c => c.url.startsWith(`${GRAPH_V1}/teams/`) && c.url.includes('/messages/'))
      expect(message).toBeDefined()
      expect(message!.url).toBe(`${GRAPH_V1}/teams/${GID}/channels/${encodeURIComponent(CHAN)}/messages/MSG`)
    }
  })

  test('S8 a chat lookup reads /chats/{id}/messages/{id}', async () => {
    const f = fakeFetch(standardRoutes({ message: () => graphMessage([]) }))
    const { annotations, results } = await fetchSharedFiles(chatLookup(), SENDER, deps(f.fn, dir))
    expect(f.calls.map(c => c.url)).toEqual([
      `${GRAPH_V1}/chats/${encodeURIComponent(CHAT)}/messages/MSG`,
    ])
    expect(results).toEqual([])
    expect(annotations).toEqual([])
  })

  test('S9 Graph message 403 in a channel names the missing RSC consent', async () => {
    const f = fakeFetch(standardRoutes({ message: () => jsonResponse({ error: 'Forbidden' }, 403) }))
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(annotations.length).toBe(1)
    expect(annotations[0]).toBe(
      '[Attachment failed: 1 shared file(s) — could not read the message from Microsoft Graph (HTTP 403); ' +
        'the Teams app needs the ChannelMessage.Read.Group consent for this team]',
    )
    expect(f.calls.some(c => c.url.startsWith(`${GRAPH_V1}/shares/`))).toBe(false)
  })

  test('S9b Graph message 403 in a chat names the chat consent', async () => {
    const f = fakeFetch(standardRoutes({ message: () => jsonResponse({ error: 'Forbidden' }, 403) }))
    const { annotations } = await fetchSharedFiles(chatLookup(), SENDER, deps(f.fn, dir))
    expect(annotations.length).toBe(1)
    expect(annotations[0].includes('ChatMessage.Read.Chat')).toBe(true)
  })

  test('S10 a 403 on the drive item is a skip that names Files.Read.All', async () => {
    const f = fakeFetch(standardRoutes({ metadata: () => jsonResponse({ error: 'Forbidden' }, 403) }))
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('skipped')
    expect(annotations.length).toBe(1)
    expect(annotations[0].startsWith('[Attachment skipped: report.xlsx')).toBe(true)
    expect(annotations[0].includes('Files.Read.All')).toBe(true)
  })

  test('S11 a contentUrl off SharePoint is skipped without a single request for it', async () => {
    const f = fakeFetch(standardRoutes({ message: () => graphMessage([graphRef({ contentUrl: 'https://evil.example.com/a.xlsx' })]) }))
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('skipped')
    expect(annotations[0].includes('not a SharePoint file link')).toBe(true)
    expect(f.calls.some(c => c.url.includes('evil.example.com'))).toBe(false)
    expect(f.calls.some(c => c.url.startsWith(`${GRAPH_V1}/shares/`))).toBe(false)
    expect(f.calls.some(c => c.url.includes('/filesFolder'))).toBe(false)
  })

  test('S12 a redirect target that is not allowed is refused and never fetched', async () => {
    const f = fakeFetch(
      standardRoutes({ content: () => new Response(null, { status: 302, headers: { location: 'http://contoso.sharepoint.com/x' } }) }),
    )
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('failed')
    expect(annotations[0]).toBe('[Attachment failed: report.xlsx — download redirect target is not allowed]')
    expect(f.calls.some(c => c.url === 'http://contoso.sharepoint.com/x')).toBe(false)
    expect(f.calls.length).toBe(4)
  })

  test('S13 an untrusted service URL stops before any token or request', async () => {
    const f = fakeFetch(standardRoutes())
    let botTokens = 0
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ serviceUrl: 'https://attacker.example.com/' }),
      SENDER,
      deps(f.fn, dir, { getBotToken: async () => { botTokens += 1; return 'BOT' } }),
    )
    expect(botTokens).toBe(0)
    expect(f.calls.length).toBe(0)
    expect(results.length).toBe(1)
    expect(annotations[0].includes('could not resolve the team id (untrusted or missing service URL)')).toBe(true)
  })

  test('S14 a 200 content body is saved directly, with no hop', async () => {
    const f = fakeFetch(standardRoutes({ content: () => new Response(BODY) }))
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('saved')
    expect(f.calls.map(c => c.url)).toEqual([
      `${GRAPH_V1}/teams/${GID}/channels/${encodeURIComponent(CHAN)}/messages/ROOT/replies/MSG`,
      metaUrl(),
      filesFolderUrl(),
      contentUrl(),
    ])
    expect(annotations[0].startsWith('[Attached: ')).toBe(true)
  })

  test('S15 a body over MAX_ATTACHMENT_BYTES is skipped for being too big', async () => {
    const f = fakeFetch(standardRoutes({ content: () => new Response(oversizedBody()) }))
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('skipped')
    expect(annotations[0].startsWith('[Attachment skipped: report.xlsx')).toBe(true)
    expect(annotations[0].includes('exceeds ')).toBe(true)
  })

  test('S16 references whose id is not in refIds, and non-reference attachments, are ignored', async () => {
    const f = fakeFetch(
      standardRoutes({
        message: () =>
          graphMessage([
            graphRef({ id: 'someone-elses-file' }),
            graphRef({ contentType: 'forwardedMessageReference' }),
          ]),
      }),
    )
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results).toEqual([])
    expect(annotations).toEqual([])
    expect(f.calls.some(c => c.url.startsWith(`${GRAPH_V1}/shares/`))).toBe(false)
  })

  test('S17 a fetch that throws on the message request becomes one failed result', async () => {
    const f = fakeFetch([
      {
        test: (u: string) => u.startsWith(`${GRAPH_V1}/teams/`),
        reply: () => {
          throw new TypeError('fetch failed')
        },
      },
    ])
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('failed')
    expect(annotations).toEqual(['[Attachment failed: shared file — fetch failed]'])
  })

  test('X3 a drive other than the channel files folder is skipped, nothing is downloaded', async () => {
    const f = fakeFetch(standardRoutes({ metadata: () => graphMetadata({ parentReference: { driveId: OTHER_DRIVE } }) }))
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('skipped')
    expect(annotations).toEqual([
      `[Attachment skipped: report.xlsx — the file is not stored in this channel's files, so it is not downloaded]`,
    ])
    expect(f.calls.some(c => c.url.includes('/items/'))).toBe(false)
    expect(f.calls.some(c => c.url === DOWNLOAD_URL)).toBe(false)
  })

  test('X4 two references in one channel share one files-folder read and both are saved', async () => {
    const f = fakeFetch(
      standardRoutes({
        message: () => graphMessage([graphRef(), graphRef({ id: 'ref-2', name: 'notes.txt' })]),
      }),
    )
    const cache = new Map<string, string>()
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID, refIds: ['ref-1', 'ref-2'] }),
      SENDER,
      deps(f.fn, dir, { channelDriveCache: cache }),
    )
    expect(f.calls.filter(c => c.url.includes('/filesFolder')).length).toBe(1)
    expect(cache.get(`${GID}/${CHAN}`)).toBe(DRIVE)
    expect(results.length).toBe(2)
    expect(results.every(r => r.kind === 'saved')).toBe(true)
    expect(annotations.length).toBe(2)
    // and a second message in the same channel needs no new files-folder read
    const before = f.calls.length
    await fetchSharedFiles(
      channelLookup({ aadGroupId: GID, refIds: ['ref-1'] }),
      SENDER,
      deps(f.fn, dir, { channelDriveCache: cache }),
    )
    expect(f.calls.filter(c => c.url.includes('/filesFolder')).length).toBe(1)
    // message, drive item, content, hop — and no files-folder read
    expect(f.calls.length).toBe(before + 4)
  })

  test('X5 a files-folder failure is one failed result and stops the whole message', async () => {
    const f = fakeFetch(
      standardRoutes({
        filesFolder: () => jsonResponse({ error: 'Forbidden' }, 403),
        message: () => graphMessage([graphRef(), graphRef({ id: 'ref-2', name: 'notes.txt' })]),
      }),
    )
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID, refIds: ['ref-1', 'ref-2'] }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('failed')
    expect(annotations).toEqual([
      `[Attachment failed: report.xlsx — could not read the channel's files folder (HTTP 403)]`,
    ])
    // one drive-item read (for the first reference), and no content at all
    expect(f.calls.filter(c => c.url.startsWith(`${GRAPH_V1}/shares/`)).length).toBe(1)
    expect(f.calls.some(c => c.url.includes('/items/'))).toBe(false)
  })

  test('X5b a files-folder answer without a drive id fails the same way', async () => {
    const f = fakeFetch(standardRoutes({ filesFolder: () => jsonResponse({ parentReference: {} }) }))
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('failed')
    expect(annotations).toEqual([
      `[Attachment failed: report.xlsx — could not read the channel's files folder (no drive id)]`,
    ])
  })

  test('X6 a chat only takes a file from the sender own OneDrive', async () => {
    // owner id in a different letter case than the lookup's sender id
    const ok = fakeFetch(standardRoutes({ driveOwner: () => jsonResponse({ owner: { user: { id: SENDER_AAD.toLowerCase() } } }) }))
    const saved = await fetchSharedFiles(chatLookup(), SENDER, deps(ok.fn, dir))
    expect(saved.results.length).toBe(1)
    expect(saved.results[0].kind).toBe('saved')
    expect(ok.calls.map(c => c.url)).toEqual([
      `${GRAPH_V1}/chats/${encodeURIComponent(CHAT)}/messages/MSG`,
      metaUrl(),
      driveUrl(),
      contentUrl(),
      DOWNLOAD_URL,
    ])

    const other = fakeFetch(standardRoutes({ driveOwner: () => jsonResponse({ owner: { user: { id: '00000000-0000-4000-8000-0000000000cc' } } }) }))
    const skipped = await fetchSharedFiles(chatLookup(), SENDER, deps(other.fn, dir))
    expect(skipped.results.length).toBe(1)
    expect(skipped.results[0].kind).toBe('skipped')
    expect(skipped.annotations).toEqual([
      `[Attachment skipped: report.xlsx — the file is not in the sender's OneDrive, so it is not downloaded]`,
    ])
    expect(other.calls.some(c => c.url.includes('/items/'))).toBe(false)

    // no sender id on the lookup: nothing is fetched past the drive item
    const none = fakeFetch(standardRoutes())
    const noSender = await fetchSharedFiles(chatLookup({ senderAadObjectId: undefined }), SENDER, deps(none.fn, dir))
    expect(noSender.results.length).toBe(1)
    expect(noSender.results[0].kind).toBe('skipped')
    expect(noSender.annotations).toEqual([
      `[Attachment skipped: report.xlsx — the file is not in the sender's OneDrive, so it is not downloaded]`,
    ])
    expect(none.calls.some(c => c.url.startsWith(`${GRAPH_V1}/drives/`))).toBe(false)
  })

  test('X6b a drive-owner read that fails or is empty is a failed result', async () => {
    const f = fakeFetch(standardRoutes({ driveOwner: () => jsonResponse({ owner: {} }, 403) }))
    const { annotations, results } = await fetchSharedFiles(chatLookup(), SENDER, deps(f.fn, dir))
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('failed')
    expect(annotations).toEqual([
      `[Attachment failed: report.xlsx — could not read the file's drive owner (HTTP 403)]`,
    ])

    const g = fakeFetch(standardRoutes({ driveOwner: () => jsonResponse({ owner: { user: {} } }) }))
    const noOwner = await fetchSharedFiles(chatLookup(), SENDER, deps(g.fn, dir))
    expect(noOwner.annotations).toEqual([
      `[Attachment failed: report.xlsx — could not read the file's drive owner (no owner)]`,
    ])
  })

  test('X7 the drive-item status mapping', async () => {
    const cases: [number, { kind: string; reason: string }][] = [
      [401, {
        kind: 'skipped',
        reason:
          'the Teams app has no Microsoft Graph permission to read this file (HTTP 401); ' +
          'it needs the Files.Read.All application permission',
      }],
      [403, {
        kind: 'skipped',
        reason:
          'the Teams app has no Microsoft Graph permission to read this file (HTTP 403); ' +
          'it needs the Files.Read.All application permission',
      }],
      [404, { kind: 'failed', reason: 'file not found in SharePoint (HTTP 404)' }],
      [500, { kind: 'failed', reason: 'HTTP 500' }],
    ]
    for (const [status, expected] of cases) {
      const f = fakeFetch(standardRoutes({ metadata: () => jsonResponse({ error: 'x' }, status) }))
      const { annotations, results } = await fetchSharedFiles(
        channelLookup({ aadGroupId: GID }),
        SENDER,
        deps(f.fn, dir),
      )
      expect(results.length).toBe(1)
      expect(results[0].kind).toBe(expected.kind)
      expect(annotations).toEqual([`[Attachment ${expected.kind === 'skipped' ? 'skipped' : 'failed'}: report.xlsx — ${expected.reason}]`])
      expect(f.calls.some(c => c.url.includes('/filesFolder'))).toBe(false)
    }

    const f = fakeFetch(standardRoutes({ metadata: () => jsonResponse({ id: ITEM, name: 'report.xlsx', size: 1 }) }))
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('failed')
    expect(annotations).toEqual([
      '[Attachment failed: report.xlsx — Microsoft Graph returned no drive item for this file]',
    ])
  })

  test('X8 a size over the cap from the drive item is skipped before any download', async () => {
    const f = fakeFetch(standardRoutes({ metadata: () => graphMetadata({ size: MAX_ATTACHMENT_BYTES + 1 }) }))
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('skipped')
    expect(annotations).toEqual([
      `[Attachment skipped: report.xlsx — exceeds ${formatBytes(MAX_ATTACHMENT_BYTES)} limit]`,
    ])
    expect(formatBytes(MAX_ATTACHMENT_BYTES)).toBe('50.0 MB')
    expect(f.calls.some(c => c.url.includes('/items/'))).toBe(false)
  })

  test('X9 every redirect status is followed by hand, and only to SharePoint', async () => {
    for (const status of [301, 303, 307, 308]) {
      const f = fakeFetch(
        standardRoutes({
          content: () => new Response(null, { status, headers: { location: DOWNLOAD_URL } }),
        }),
      )
      const { results } = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(f.fn, dir))
      expect(results.length).toBe(1)
      expect(results[0].kind).toBe('saved')
      expect(f.calls[f.calls.length - 1].url).toBe(DOWNLOAD_URL)
      expect(f.calls[f.calls.length - 1].redirect).toBe('manual')
    }

    // a relative Location is resolved against the content URL's origin, which
    // is Graph, not SharePoint — so it is refused and never fetched
    const relative = fakeFetch(
      standardRoutes({ content: () => new Response(null, { status: 302, headers: { location: '/sites/x/a.xlsx' } }) }),
    )
    const rel = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(relative.fn, dir))
    expect(rel.results[0].kind).toBe('failed')
    expect(rel.annotations).toEqual([
      '[Attachment failed: report.xlsx — download redirect target is not allowed]',
    ])
    expect(relative.calls.some(c => c.url === 'https://graph.microsoft.com/sites/x/a.xlsx')).toBe(false)

    const userinfo = fakeFetch(
      standardRoutes({ content: () => new Response(null, { status: 302, headers: { location: 'https://u:p@contoso.sharepoint.com/x' } }) }),
    )
    const cred = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(userinfo.fn, dir))
    expect(cred.annotations).toEqual([
      '[Attachment failed: report.xlsx — download redirect target is not allowed]',
    ])
    expect(userinfo.calls.some(c => c.url.includes('u:p@'))).toBe(false)

    const evil = fakeFetch(standardRoutes({ download: () => new Response(null, { status: 302, headers: { location: 'https://evil.example.com/x' } }) }))
    const ev = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(evil.fn, dir))
    expect(ev.annotations).toEqual([
      '[Attachment failed: report.xlsx — download redirect target is not allowed]',
    ])
    expect(evil.calls.some(c => c.url.includes('evil.example.com'))).toBe(false)

    // four redirects in a row is one too many
    const hops = ['https://contoso.sharepoint.com/h1', 'https://contoso.sharepoint.com/h2', 'https://contoso.sharepoint.com/h3', 'https://contoso.sharepoint.com/h4']
    const many = fakeFetch([
      { test: u => u.includes('/messages/'), reply: () => graphMessage([graphRef()]) },
      { test: u => u.startsWith(`${GRAPH_V1}/shares/`), reply: () => graphMetadata() },
      { test: u => u.includes('/filesFolder'), reply: () => jsonResponse({ parentReference: { driveId: DRIVE } }) },
      { test: u => u === contentUrl(), reply: () => new Response(null, { status: 302, headers: { location: hops[0]! } }) },
      { test: u => u === hops[0], reply: () => new Response(null, { status: 302, headers: { location: hops[1]! } }) },
      { test: u => u === hops[1], reply: () => new Response(null, { status: 302, headers: { location: hops[2]! } }) },
      { test: u => u === hops[2], reply: () => new Response(null, { status: 302, headers: { location: hops[3]! } }) },
    ])
    const tooMany = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(many.fn, dir))
    expect(tooMany.results[0].kind).toBe('failed')
    expect(tooMany.annotations).toEqual([
      '[Attachment failed: report.xlsx — too many download redirects]',
    ])
    expect(many.calls.some(c => c.url === hops[3])).toBe(false)

    const none = fakeFetch(
      standardRoutes({ content: () => new Response(null, { status: 302 }) }),
    )
    const noLocation = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(none.fn, dir))
    expect(noLocation.results[0].kind).toBe('failed')
    expect(noLocation.annotations).toEqual([
      '[Attachment failed: report.xlsx — download redirect without a location]',
    ])
  })

  test('X9b a non-200, non-redirect content answer and hop answer', async () => {
    const f = fakeFetch(standardRoutes({ content: () => new Response('nope', { status: 404 }) }))
    const c = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(f.fn, dir))
    expect(c.results[0].kind).toBe('failed')
    expect(c.annotations).toEqual([
      '[Attachment failed: report.xlsx — file not found in SharePoint (HTTP 404)]',
    ])

    // a hop carries no token of ours, so its answer is just the status
    const g = fakeFetch(standardRoutes({ download: () => new Response('nope', { status: 404 }) }))
    const h = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(g.fn, dir))
    expect(h.results[0].kind).toBe('failed')
    expect(h.annotations).toEqual(['[Attachment failed: report.xlsx — HTTP 404]'])
  })

  test('X11 a download whose body never ends is cut off by the timeout', async () => {
    const f = fakeFetch([
      { test: u => u.includes('/messages/'), reply: () => graphMessage([graphRef()]) },
      { test: u => u.startsWith(`${GRAPH_V1}/shares/`), reply: () => graphMetadata() },
      { test: u => u.includes('/filesFolder'), reply: () => jsonResponse({ parentReference: { driveId: DRIVE } }) },
      { test: u => u === contentUrl(), reply: call => new Response(neverEndingBody(call.signal)) },
    ])
    const before = attachTemps()
    const started = Date.now()
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir, { timeoutMs: 50 }),
    )
    expect(Date.now() - started).toBeLessThan(2000)
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('failed')
    expect(annotations.length).toBe(1)
    expect(attachTemps().length).toBe(before.length)
  })

  test('X12 the name is shown, not trusted: the path stays in the dated directory', async () => {
    const newline = fakeFetch(standardRoutes({ message: () => graphMessage([graphRef({ name: 'a\nb.xlsx' })]) }))
    const nl = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(newline.fn, dir))
    expect(nl.results.length).toBe(1)
    expect(nl.results[0].originalName).toBe('a b.xlsx')
    // the agent sees the newline as a space; the file on disk gets pickPath's
    // own sanitisation, which turns it into an underscore
    expect(nl.annotations).toEqual([
      `[Attached: ${join(dir, londonDate(NOW), `${londonTime(NOW)}-${senderSlug(SENDER)}-a_b.xlsx`)} (application/xlsx, ${formatBytes(BODY.length)})]`,
    ])

    // the Unicode line / paragraph separators break a line too
    for (const sep of ['\u2028', '\u2029']) {
      const f = fakeFetch(standardRoutes({ message: () => graphMessage([graphRef({ name: `a${sep}b.xlsx` })]) }))
      const r = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(f.fn, dir))
      expect(r.results[0].originalName).toBe('a b.xlsx')
    }

    const escape = fakeFetch(standardRoutes({ message: () => graphMessage([graphRef({ name: '../../escape.xlsx' })]) }))
    const esc = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(escape.fn, dir))
    const saved = esc.results[0]
    expect(saved.kind).toBe('saved')
    if (saved.kind !== 'saved') return
    expect(dirname(saved.path)).toBe(join(dir, londonDate(NOW)))
    expect(saved.path.startsWith(join(dir, londonDate(NOW), `${londonTime(NOW)}-${senderSlug(SENDER)}-`))).toBe(true)
    expect(statSync(join(dir, londonDate(NOW))).mode & 0o777).toBe(0o700)
  })

  test('X13 a placeholder with no matching reference logs one stderr line', async () => {
    const f = fakeFetch(standardRoutes({ message: () => graphMessage([graphRef({ id: 'someone-elses' })]) }))
    const { value, lines } = await captureStderr(() =>
      fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(f.fn, dir)),
    )
    expect(value.results).toEqual([])
    expect(value.annotations).toEqual([])
    expect(lines).toContain(
      'teams channel: shared file lookup found no file reference for 1 placeholder(s) in message MSG\n',
    )
  })

  test('X15 the team-id translation and its failures', async () => {
    // an aadGroupId that is not a guid falls through to the connector
    const called = fakeFetch(standardRoutes())
    await fetchSharedFiles(channelLookup({ aadGroupId: 'not-a-guid' }), SENDER, deps(called.fn, dir))
    expect(called.calls.some(c => c.url.includes('/v3/teams/'))).toBe(true)

    const bad = fakeFetch(standardRoutes({ team: () => jsonResponse({ aadGroupId: 'not-a-guid' }) }))
    const noGuid = await fetchSharedFiles(channelLookup(), SENDER, deps(bad.fn, dir))
    expect(noGuid.results[0].kind).toBe('failed')
    expect(noGuid.annotations).toEqual([
      '[Attachment failed: 1 shared file(s) — could not resolve the team id (no aadGroupId)]',
    ])

    const err = fakeFetch(standardRoutes({ team: () => jsonResponse({ error: 'boom' }, 500) }))
    const http500 = await fetchSharedFiles(channelLookup(), SENDER, deps(err.fn, dir))
    expect(http500.results[0].kind).toBe('failed')
    expect(http500.annotations).toEqual([
      '[Attachment failed: 1 shared file(s) — could not resolve the team id (HTTP 500)]',
    ])
  })

  test('X16 message 401 in a channel and a chat, and unusable reference entries', async () => {
    const chan = fakeFetch(standardRoutes({ message: () => jsonResponse({ error: 'x' }, 401) }))
    const c = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(chan.fn, dir))
    expect(c.annotations).toEqual([
      '[Attachment failed: 1 shared file(s) — could not read the message from Microsoft Graph (HTTP 401); ' +
        'the Teams app needs the ChannelMessage.Read.Group consent for this team]',
    ])

    const chat = fakeFetch(standardRoutes({ message: () => jsonResponse({ error: 'x' }, 401) }))
    const g = await fetchSharedFiles(chatLookup(), SENDER, deps(chat.fn, dir))
    expect(g.annotations).toEqual([
      '[Attachment failed: 1 shared file(s) — could not read the message from Microsoft Graph (HTTP 401); ' +
        'the Teams app needs the ChatMessage.Read.Chat consent for this chat]',
    ])

    // a reference whose contentUrl is not a string is not a file reference
    const odd = fakeFetch(
      standardRoutes({ message: () => graphMessage([{ id: 'ref-1', contentType: 'reference', contentUrl: 42, name: 'x' }]) }),
    )
    const o = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(odd.fn, dir))
    expect(o.results).toEqual([])
    expect(odd.calls.some(c => c.url.startsWith(`${GRAPH_V1}/shares/`))).toBe(false)

    // an empty Graph name falls back to the last URL segment
    const unnamed = fakeFetch(standardRoutes({ message: () => graphMessage([graphRef({ name: '' })]) }))
    const u = await fetchSharedFiles(channelLookup({ aadGroupId: GID }), SENDER, deps(unnamed.fn, dir))
    expect(u.results.length).toBe(1)
    expect(u.results[0].originalName).toBe('report.xlsx')
  })
})

describe('sharedFiles: saveResponseToTemp', () => {
  test('X10 a declared size over the cap throws before the body is read', async () => {
    let pulled = false
    const res = new Response(
      new ReadableStream({
        pull() {
          pulled = true
        },
      }),
      { status: 200, headers: { 'content-length': String(MAX_ATTACHMENT_BYTES + 1) } },
    )
    let err: unknown
    try {
      await saveResponseToTemp(res)
    } catch (e) {
      err = e
    }
    expect(err instanceof Error ? err.message : String(err)).toBe(
      `exceeds ${formatBytes(MAX_ATTACHMENT_BYTES)} limit`,
    )
    expect(pulled).toBe(false)
  })

  test('X10b a streamed body over the cap throws and leaves no temp file', async () => {
    const before = attachTemps()
    let err: unknown
    try {
      await saveResponseToTemp(new Response(oversizedBody()))
    } catch (e) {
      err = e
    }
    expect(err instanceof Error ? err.message : String(err)).toBe(
      `exceeds ${formatBytes(MAX_ATTACHMENT_BYTES)} limit`,
    )
    expect(attachTemps().length).toBe(before.length)
  })

  test('X10c the temp file is private, and an empty body makes an empty file', async () => {
    const { tempPath, sizeBytes } = await saveResponseToTemp(new Response('hello'))
    expect(sizeBytes).toBe(5)
    expect(statSync(tempPath).mode & 0o777).toBe(0o600)
    expect(readFileSync(tempPath, 'utf8')).toBe('hello')
    rmSync(tempPath, { force: true })

    const empty = await saveResponseToTemp(new Response(null, { status: 200 }))
    expect(empty.sizeBytes).toBe(0)
    expect(statSync(empty.tempPath).mode & 0o777).toBe(0o600)
    rmSync(empty.tempPath, { force: true })
  })

  test('X10d a temp name that is already taken is not created over, nor removed', async () => {
    const now = spyOn(Date, 'now').mockReturnValue(1700000000000)
    const rnd = spyOn(Math, 'random').mockReturnValue(0.5)
    const taken = join(tmpdir(), `cct-attach-1700000000000-${(0.5).toString(36).slice(2, 8)}`)
    writeFileSync(taken, 'not ours')
    try {
      let err: unknown
      try {
        await saveResponseToTemp(new Response('hello'))
      } catch (e) {
        err = e
      }
      expect((err as NodeJS.ErrnoException | undefined)?.code).toBe('EEXIST')
      expect(readFileSync(taken, 'utf8')).toBe('not ours')
    } finally {
      now.mockRestore()
      rnd.mockRestore()
      rmSync(taken, { force: true })
    }
  })
})

describe('attachments: logAttachmentResult', () => {
  test('X14 one exact stderr line per outcome', async () => {
    const { value, lines } = await captureStderr(async () => {
      logAttachmentResult({ kind: 'saved', path: '/x/a.xlsx', contentType: 'application/xlsx', sizeBytes: 10, originalName: 'a.xlsx' })
      logAttachmentResult({ kind: 'skipped', reason: 'not a SharePoint file link', originalName: 'a.xlsx' })
      logAttachmentResult({ kind: 'failed', reason: 'HTTP 500', originalName: 'a.xlsx' })
      logAttachmentResult({ kind: 'ignored', reason: 'html repeats the message text', originalName: 'a.xlsx' })
      logAttachmentResult({ kind: 'inline', label: 'HTML content', text: 'hi', originalName: 'a.xlsx' })
      return 'done'
    })
    expect(value).toBe('done')
    expect(lines).toEqual([
      'teams channel: attachment saved — /x/a.xlsx (10 B)\n',
      'teams channel: attachment skipped — a.xlsx: not a SharePoint file link\n',
      'teams channel: attachment failed — a.xlsx: HTTP 500\n',
      'teams channel: attachment ignored — a.xlsx: html repeats the message text\n',
      'teams channel: attachment inlined — a.xlsx (HTML content, 2 chars)\n',
    ])
  })
})

/** A body of more than MAX_ATTACHMENT_BYTES, one 1 MB chunk per pull so the
 *  allocation stays small while the byte count crosses the cap. */
function oversizedBody(): ReadableStream<Uint8Array> {
  const chunk = new Uint8Array(1024 * 1024)
  let left = Math.ceil((MAX_ATTACHMENT_BYTES + 1) / chunk.byteLength)
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (left > 0) {
        left -= 1
        controller.enqueue(chunk)
      } else {
        controller.close()
      }
    },
  })
}

/** A body that never yields a chunk and errors as soon as the request aborts. */
function neverEndingBody(signal?: AbortSignal): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      const fail = () => controller.error(new Error('body aborted'))
      if (signal?.aborted) fail()
      else signal?.addEventListener('abort', fail, { once: true })
    },
  })
}
