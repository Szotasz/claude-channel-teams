import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { resolveSharePoint } from '../src/config'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createFileSender, isSharedConversationRef } from '../src/teams/sendFile'
import { createSharePointUploader } from '../src/teams/sharepoint'

const json = (status: number, body: unknown) =>
  new Response(JSON.stringify(body), { status })
const TOKEN = () => json(200, { access_token: 'fake-tok', expires_in: 3600 })
const ITEM = (status = 200) => json(status, { webUrl: 'https://contoso.sharepoint.com/i', name: 'n.pdf' })

function setup(queue: Response[]) {
  const calls: { url: string; method: string; headers: Record<string, string>; body: any; redirect?: string }[] = []
  let t = 1_000_000
  const fetchFn = (async (url: any, init: any = {}) => {
    calls.push({ url: String(url), method: init.method, headers: { ...(init.headers ?? {}) }, body: init.body, redirect: init.redirect })
    const r = queue.shift()
    if (!r) throw new Error('queue empty')
    return r
  }) as any
  const up = createSharePointUploader({
    tenantId: 'ten', appId: 'aid', appPassword: 'the-secret',
    siteId: 'contoso.sharepoint.com,1,2', folder: 'Teams bot files',
    fetch: fetchFn, now: () => t,
  })
  return { calls, up, advance: (ms: number) => { t += ms } }
}

describe('token', () => {
  test('requests, caches, refreshes', async () => {
    const s = setup([TOKEN(), ITEM(), ITEM(), TOKEN(), ITEM()])
    await s.up.upload('a.pdf', Buffer.from('x'))
    expect(s.calls[0].url).toBe('https://login.microsoftonline.com/ten/oauth2/v2.0/token')
    expect(s.calls[0].method).toBe('POST')
    expect(s.calls[0].headers['Content-Type']).toBe('application/x-www-form-urlencoded')
    const form = new URLSearchParams(String(s.calls[0].body))
    expect(form.get('client_id')).toBe('aid')
    expect(form.get('client_secret')).toBe('the-secret')
    expect(form.get('scope')).toBe('https://graph.microsoft.com/.default')
    expect(form.get('grant_type')).toBe('client_credentials')
    await s.up.upload('a.pdf', Buffer.from('x'))
    expect(s.calls.length).toBe(3)
    s.advance((3600 - 60) * 1000 + 1)
    await s.up.upload('a.pdf', Buffer.from('x'))
    expect(s.calls.length).toBe(5)
    expect(s.calls[3].url).toContain('/oauth2/v2.0/token')
  })

  test('failure message is scrubbed', async () => {
    const s = setup([json(401, { error: 'invalid_client', error_description: 'bad the-secret' })])
    let msg = ''
    try { await s.up.upload('a.pdf', Buffer.from('x')) } catch (e: any) { msg = e.message }
    expect(msg).toBe('SharePoint token request failed (HTTP 401, invalid_client)')
    expect(msg).not.toContain('the-secret')
  })
})

describe('upload', () => {
  test('small upload', async () => {
    // The second upload reuses the cached token: no second token request.
    const s = setup([TOKEN(), ITEM(201), ITEM()])
    const r = await s.up.upload('a bé.pdf', Buffer.from('x'))
    expect(r).toEqual({ webUrl: 'https://contoso.sharepoint.com/i', name: 'n.pdf' })
    expect(s.calls[1].method).toBe('PUT')
    expect(s.calls[1].url).toBe(
      'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,1,2/drive/root:/Teams%20bot%20files/a%20b%C3%A9.pdf:/content?@microsoft.graph.conflictBehavior=rename',
    )
    expect(s.calls[1].headers.Authorization).toBe('Bearer fake-tok')
    await s.up.upload('x/y.pdf', Buffer.from('x'))
    expect(s.calls[2].url).toContain('/Teams%20bot%20files/x_y.pdf:/content')
  })

  test('chunked upload', async () => {
    const s = setup([
      TOKEN(),
      json(200, { uploadUrl: 'https://contoso.sharepoint.com/up' }),
      json(202, {}),
      ITEM(201),
    ])
    const r = await s.up.upload('big.bin', Buffer.alloc(5 * 1024 * 1024))
    expect(r.name).toBe('n.pdf')
    expect(s.calls[1].url).toEndWith(':/createUploadSession')
    expect(s.calls[1].method).toBe('POST')
    expect(s.calls[2].url).toBe('https://contoso.sharepoint.com/up')
    expect(s.calls[2].headers['Content-Range']).toBe('bytes 0-3932159/5242880')
    expect(s.calls[3].headers['Content-Range']).toBe('bytes 3932160-5242879/5242880')
    expect(s.calls[2].headers.Authorization).toBeUndefined()
    expect(s.calls[3].headers.Authorization).toBeUndefined()
    // No request follows a redirect: not the token request, the session, or a chunk.
    expect(s.calls.map(c => c.redirect)).toEqual(['error', 'error', 'error', 'error'])
  })

  test('the single PUT does not follow redirects, and a dot filename stays in the folder', async () => {
    const s = setup([TOKEN(), ITEM(201), ITEM(201)])
    await s.up.upload('..', Buffer.from('x'))
    expect(s.calls[1].redirect).toBe('error')
    expect(s.calls[1].url).toContain('/Teams%20bot%20files/_:/content')
    await s.up.upload('.', Buffer.from('x'))
    expect(s.calls[2].url).toContain('/Teams%20bot%20files/_:/content')
  })

  for (const bad of ['https://evil.example/up', 'http://x.sharepoint.com/up']) {
    test(`rejects upload url ${bad}`, async () => {
      const s = setup([TOKEN(), json(200, { uploadUrl: bad })])
      await expect(s.up.upload('big.bin', Buffer.alloc(5 * 1024 * 1024))).rejects.toThrow(
        'SharePoint returned an unexpected upload URL',
      )
      expect(s.calls.length).toBe(2)
    })
  }
})

describe('resolveSharePoint', () => {
  test('cases', () => {
    expect(resolveSharePoint({})).toBeNull()
    const id = 'contoso.sharepoint.com,11111111-2222-4333-8444-555555555555,66666666-7777-4888-9999-aaaaaaaaaaaa'
    expect(resolveSharePoint({ TEAMS_SHAREPOINT_SITE_ID: id })).toEqual({ siteId: id, folder: 'Teams bot files' })
    const long253 = Array.from({ length: 4 }, () => 'a'.repeat(62)).join('.') + '.c'
    expect(long253.length).toBe(253)
    expect(resolveSharePoint({ TEAMS_SHAREPOINT_SITE_ID: long253 + ',' + id.split(',').slice(1).join(',') })).not.toBeNull()
    expect(resolveSharePoint({ TEAMS_SHAREPOINT_SITE_ID: 'a/b' })).toBeNull()
    const g = '11111111-2222-4333-8444-555555555555,66666666-7777-4888-9999-aaaaaaaaaaaa'
    for (const bad of ['not-a-site-id', '..', 'contoso.sharepoint.com,1,2', 'contoso.sharepoint.com,' + 'x'.repeat(36) + ',' + 'y'.repeat(36), '-bad.sharepoint.com,' + g, 'bad-.sharepoint.com,' + g, 'contoso,' + g, Array.from({ length: 5 }, () => 'a'.repeat(60)).join('.') + '.com,' + g]) {
      expect(resolveSharePoint({ TEAMS_SHAREPOINT_SITE_ID: bad })).toBeNull()
    }
    expect(resolveSharePoint({ TEAMS_SHAREPOINT_SITE_ID: id, TEAMS_SHAREPOINT_FOLDER: 'x/../y' })).toBeNull()
    expect(resolveSharePoint({ TEAMS_SHAREPOINT_SITE_ID: id, TEAMS_SHAREPOINT_FOLDER: 'Bids/2026/' })).toEqual({
      siteId: id, folder: 'Bids/2026',
    })
  })
})

describe('fetch errors never carry the request URL', () => {
  test('a redirect error with the pre-authenticated URL in its message is rethrown without it', async () => {
    const calls: string[] = []
    const responses = [
      new Response(JSON.stringify({ access_token: 'fake-tok', expires_in: 3600 }), { status: 200 }),
      new Response(JSON.stringify({ uploadUrl: 'https://contoso.sharepoint.com/up?tempauth=CAPABILITY123' }), { status: 200 }),
    ]
    const fetchFn = (async (url: any) => {
      calls.push(String(url))
      const r = responses.shift()
      if (r) return r
      const e: any = new Error(`UnexpectedRedirect fetching "${String(url)}".`)
      e.code = 'UnexpectedRedirect'
      throw e
    }) as any
    const up = createSharePointUploader({
      tenantId: 'ten', appId: 'aid', appPassword: 'the-secret',
      siteId: 'contoso.sharepoint.com,1,2', folder: 'F', fetch: fetchFn,
    })
    let msg = ''
    try { await up.upload('big.bin', Buffer.alloc(5 * 1024 * 1024)) } catch (e: any) { msg = e.message }
    expect(msg).toBe('SharePoint request failed (UnexpectedRedirect)')
    expect(msg).not.toContain('CAPABILITY123')
  })
})

describe('sendFile with SharePoint', () => {
  const ID = '00000000-0000-4000-8000-000000000001'
  function harness(convId: string, conversation: any, sp: any, downloadBaseUrl: string | null) {
    const dir = mkdtempSync(join(tmpdir(), 'cct-sp-'))
    const allowlist = createAllowlist(join(dir, 'a.json'))
    allowlist.addEntry(ID)
    const refs = createConversationRefStore()
    refs.put(convId, { conversation } as any, ID)
    let sent: any
    let minted = 0
    const adapter = {
      async continueConversationAsync(_a: string, _r: any, logic: any) {
        await logic({ async sendActivity(a: any) { sent = a } })
      },
    } as any
    const outbox = {
      isPathAllowed: () => true,
      async mint(m: any) { minted++; return { token: 'tok12345678', filename: m.filename, sizeBytes: m.content.length, mime: m.mime } },
    } as any
    const sendFile = createFileSender({
      config: { appId: 'x' }, adapter, allowlist, refs, outbox, downloadBaseUrl, sharePoint: sp,
    })
    return { sendFile, sent: () => sent, minted: () => minted, convId }
  }
  const ok = () => ({ calls: 0, async upload() { this.calls++; return { webUrl: 'https://contoso.sharepoint.com/x', name: 'r.pdf' } } })

  test('shared conversation uses SharePoint, even without downloadBaseUrl', async () => {
    const sp = ok()
    const h = harness('c1', { id: '19:c@thread.tacv2', conversationType: 'channel' }, sp, null)
    const r = await h.sendFile({ conversationId: 'c1', content: Buffer.from('x'), filename: 'r.pdf' })
    expect(h.sent().attachments[0].content.actions[0].url).toBe('https://contoso.sharepoint.com/x')
    expect(h.minted()).toBe(0)
    expect(r.webUrl).toBe('https://contoso.sharepoint.com/x')
    expect(r.token).toBeUndefined()
  })

  test('1:1 keeps the download link', async () => {
    const sp = ok()
    const h = harness('c2', { id: 'a:1' }, sp, 'https://bot.test/files')
    const r = await h.sendFile({ conversationId: 'c2', content: Buffer.from('x'), filename: 'r.pdf' })
    expect(sp.calls).toBe(0)
    expect(h.minted()).toBe(1)
    expect(r.token).toBeString()
  })

  test('failure falls back to the download link', async () => {
    const sp = { async upload() { throw new Error('boom') } }
    const h = harness('c3', { id: '19:c@thread.tacv2', conversationType: 'channel' }, sp, 'https://bot.test/files')
    const r = await h.sendFile({ conversationId: 'c3', content: Buffer.from('x'), filename: 'r.pdf' })
    expect(h.minted()).toBe(1)
    expect(r.token).toBeString()
  })

  test('a card that fails to post after a successful upload is an error, not a second card', async () => {
    const sp = ok()
    const h = harness('c4', { id: '19:c@thread.tacv2', conversationType: 'channel' }, sp, 'https://bot.test/files')
    ;(h as any).sendFile = createFileSender({
      config: { appId: 'x' },
      adapter: { async continueConversationAsync() { throw new Error('connector down') } } as any,
      allowlist: (() => { const d = mkdtempSync(join(tmpdir(), 'cct-sp-')); const a = createAllowlist(join(d, 'a.json')); a.addEntry(ID); return a })(),
      refs: (() => { const r = createConversationRefStore(); r.put('c4', { conversation: { id: '19:c@thread.tacv2', conversationType: 'channel' } } as any, ID); return r })(),
      outbox: { isPathAllowed: () => true, async mint() { throw new Error('mint must not run') } } as any,
      downloadBaseUrl: 'https://bot.test/files',
      sharePoint: sp,
    })
    await expect((h as any).sendFile({ conversationId: 'c4', content: Buffer.from('x'), filename: 'r.pdf' })).rejects.toThrow('connector down')
    expect(sp.calls).toBe(1)
  })
})

describe('isSharedConversationRef', () => {
  test('cases', () => {
    expect(isSharedConversationRef({ conversation: { id: 'a', conversationType: 'channel' } } as any)).toBe(true)
    expect(isSharedConversationRef({ conversation: { id: 'a', conversationType: 'groupChat' } } as any)).toBe(true)
    expect(isSharedConversationRef({ conversation: { id: 'a', isGroup: true } } as any)).toBe(true)
    expect(isSharedConversationRef({ conversation: { id: '19:x@thread.v2' } } as any)).toBe(true)
    expect(isSharedConversationRef({ conversation: { id: 'a:1' } } as any)).toBe(false)
  })
})
