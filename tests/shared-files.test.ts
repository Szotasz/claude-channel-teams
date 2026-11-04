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

import { describe, test, expect, beforeEach, afterEach } from 'bun:test'
import { mkdtempSync, rmSync, readFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
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
import { MAX_ATTACHMENT_BYTES } from '../src/teams/attachments'

const TEAM = '19:team0001@thread.tacv2'
const CHAN = '19:chan0001@thread.tacv2'
const CHAT = '19:chat0001@thread.v2'
const GID = '00000000-0000-4000-8000-0000000000aa'
const SERVICE = 'https://smba.trafficmanager.net/emea/00000000-0000-4000-8000-000000000002/'
const SENDER = 'Test User'
const NOW = new Date('2026-07-01T10:30:00Z')
const FILE_URL = 'https://contoso.sharepoint.com/sites/x/Shared Documents/General/report.xlsx'
const BODY = 'xlsx-bytes'
const DOWNLOAD_URL = 'https://contoso.sharepoint.com/sites/x/_layouts/15/download.aspx?file=report.xlsx'

function html(content: string): Attachment {
  return { contentType: 'text/html', content } as Attachment
}

function graphMessage(attachments: unknown[]): Response {
  return jsonResponse({ attachments })
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

type Call = { url: string; method: string; headers: Record<string, string>; redirect?: string }

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
    }
    calls.push(call)
    for (const r of routes) if (r.test(call.url)) return r.reply(call)
    throw new Error(`unexpected request: ${call.url}`)
  }) as typeof fetch
  return { fn, calls }
}

/** The usual happy-path routing: bot team lookup, Graph message, share, download. */
function standardRoutes(overrides: {
  message?: () => Response
  share?: () => Response
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
      reply: () =>
        overrides.share?.() ??
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

describe('sharedFiles: attachmentRefIds', () => {
  test('S1 collects ids from double- and single-quoted tags, drops dupes and unsafe ids', () => {
    const part = html(
      '<p>see this</p>' +
        '<attachment id="ref-1"></attachment>' +
        "<attachment id='ref-2'></attachment>" +
        '<attachment id="ref-1"></attachment>' +
        '<attachment id="has/slash"></attachment>' +
        '<attachment id="has&quot;quote"></attachment>',
    )
    expect(attachmentRefIds([part])).toEqual(['ref-1', 'ref-2'])
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
    expect(f.calls.length).toBe(4)

    expect(f.calls[0].url).toBe(`${SERVICE.replace(/\/+$/, '')}/v3/teams/${encodeURIComponent(TEAM)}`)
    expect(f.calls[0].headers.Authorization).toBe('Bearer BOT')

    expect(f.calls[1].url).toBe(
      `${GRAPH_V1}/teams/${GID}/channels/${encodeURIComponent(CHAN)}/messages/ROOT/replies/MSG`,
    )
    expect(f.calls[1].headers.Authorization).toBe('Bearer GRAPH')

    expect(f.calls[2].url).toBe(`${GRAPH_V1}/shares/${encodeSharingUrl(FILE_URL)}/driveItem/content`)
    expect(f.calls[2].redirect).toBe('manual')
    expect(f.calls[2].headers.Authorization).toBe('Bearer GRAPH')

    expect(f.calls[3].url).toBe(DOWNLOAD_URL)
    expect(Object.keys(f.calls[3].headers).some(k => k.toLowerCase() === 'authorization')).toBe(false)

    expect(results.length).toBe(1)
    const saved = results[0]
    expect(saved.kind).toBe('saved')
    if (saved.kind !== 'saved') return
    expect(saved.path.startsWith(dir)).toBe(true)
    expect(saved.contentType).toBe('application/xlsx')
    expect(saved.originalName).toBe('report.xlsx')
    expect(readFileSync(saved.path, 'utf8')).toBe(BODY)
    expect(annotations).toEqual([`[Attached: ${saved.path} (application/xlsx, ${BODY.length} B)]`])
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
  })

  test('S7 a root message and threadRootId === messageId both read the message, not a reply', async () => {
    for (const override of [{ threadRootId: undefined }, { threadRootId: 'MSG' }]) {
      const f = fakeFetch(standardRoutes())
      await fetchSharedFiles(channelLookup(override), SENDER, deps(f.fn, dir))
      const message = f.calls.find(c => c.url.startsWith(`${GRAPH_V1}/teams/`))
      expect(message).toBeDefined()
      expect(message!.url).toBe(`${GRAPH_V1}/teams/${GID}/channels/${encodeURIComponent(CHAN)}/messages/MSG`)
    }
  })

  test('S8 a chat lookup reads /chats/{id}/messages/{id}', async () => {
    const f = fakeFetch(standardRoutes({ message: () => graphMessage([]) }))
    const lookup: SharedFileLookup = { kind: 'chat', chatId: CHAT, messageId: 'MSG', refIds: ['ref-1'] }
    const { annotations, results } = await fetchSharedFiles(lookup, SENDER, deps(f.fn, dir))
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
    expect(annotations[0].includes('could not read the message from Microsoft Graph (HTTP 403)')).toBe(true)
    expect(annotations[0].includes('ChannelMessage.Read.Group')).toBe(true)
    expect(f.calls.some(c => c.url.startsWith(`${GRAPH_V1}/shares/`))).toBe(false)
  })

  test('S9b Graph message 403 in a chat names the chat consent', async () => {
    const f = fakeFetch(standardRoutes({ message: () => jsonResponse({ error: 'Forbidden' }, 403) }))
    const { annotations } = await fetchSharedFiles(
      { kind: 'chat', chatId: CHAT, messageId: 'MSG', refIds: ['ref-1'] },
      SENDER,
      deps(f.fn, dir),
    )
    expect(annotations.length).toBe(1)
    expect(annotations[0].includes('ChatMessage.Read.Chat')).toBe(true)
  })

  test('S10 a 403 on the share content is a skip that names Files.Read.All', async () => {
    const f = fakeFetch(standardRoutes({ share: () => jsonResponse({ error: 'Forbidden' }, 403) }))
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
  })

  test('S12 a non-https redirect target is refused and never fetched', async () => {
    const f = fakeFetch(
      standardRoutes({ share: () => new Response(null, { status: 302, headers: { location: 'http://contoso.sharepoint.com/x' } }) }),
    )
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('failed')
    expect(annotations[0].includes('download redirect is not https')).toBe(true)
    expect(f.calls.some(c => c.url === 'http://contoso.sharepoint.com/x')).toBe(false)
    expect(f.calls.length).toBe(2)
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

  test('S14 a 200 share content body is saved directly', async () => {
    const f = fakeFetch(standardRoutes({ share: () => new Response(BODY) }))
    const { annotations, results } = await fetchSharedFiles(
      channelLookup({ aadGroupId: GID }),
      SENDER,
      deps(f.fn, dir),
    )
    expect(results.length).toBe(1)
    expect(results[0].kind).toBe('saved')
    expect(f.calls.length).toBe(2)
    expect(annotations[0].startsWith('[Attached: ')).toBe(true)
  })

  test('S15 a body over MAX_ATTACHMENT_BYTES is skipped for being too big', async () => {
    const tooBig = new Uint8Array(MAX_ATTACHMENT_BYTES + 1)
    const f = fakeFetch(standardRoutes({ share: () => new Response(tooBig) }))
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
})
