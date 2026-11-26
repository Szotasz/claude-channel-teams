/**
 * TEAMSIMGEMPTY1120: an image sent to the bot in a 1:1 chat arrived as a 0-character channel event.
 *
 * Measured paths to a 0-character event (real inbound path: makeTurnHandler + processAttachments wired as in
 * server.ts):
 *   - a text/html part whose only content is an <img> (no image/* part next to it): the text/html branch
 *     returned 'ignored', no annotation, empty text;
 *   - no attachment at all and empty text.
 * These tests drive the same path: the adapter turn handler with the production attachment handler.
 */

import { describe, expect, test, beforeEach, afterEach } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from './helpers/sdk'

import { makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { htmlImageSources, MAX_HTML_IMAGES, processAttachments } from '../src/teams/attachments'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const TOKEN = 'SECRET-BOT-TOKEN-xyz'
const SMBA = 'https://smba.trafficmanager.net/emea/v3/attachments/0-weu-d1-abc/views/original'
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0])
const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0, 0x10, 0x4a, 0x46, 0x49, 0x46, 0, 1])

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

interface FetchCall { url: string; authorization?: string }

function personal(overrides: Record<string, unknown>): Activity {
  return {
    type: 'message',
    id: 'msg-1',
    timestamp: new Date('2026-11-20T06:00:00Z'),
    from: { id: 'user-1', name: 'Owner', aadObjectId: FIXTURE_ID },
    conversation: { id: 'a:1personal', tenantId: TENANT_ID, conversationType: 'personal' },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/emea/',
    ...overrides,
  } as unknown as Activity
}

let dir: string
beforeEach(() => { dir = mkdtempSync(join(tmpdir(), 'cct-imgempty-')) })
afterEach(() => { rmSync(dir, { recursive: true, force: true }) })

/** Runs one activity through the turn handler with the production attachment handler; captures stderr. */
async function deliver(activity: Activity, body: Uint8Array = PNG) {
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(FIXTURE_ID)
  const events: any[] = []
  const calls: FetchCall[] = []
  let tokenRequests = 0
  const fetchImpl = (async (u: any, init?: RequestInit) => {
    const h = (init?.headers ?? {}) as Record<string, string>
    calls.push({ url: String(u), authorization: h['Authorization'] ?? h['authorization'] })
    return new Response(body, { status: 200 })
  }) as unknown as typeof fetch
  const log: string[] = []
  const write = process.stderr.write
  ;(process.stderr as any).write = (chunk: unknown) => { log.push(String(chunk)); return true }
  try {
    const turn = makeTurnHandler({
      config: {
        stateDir: dir, allowlistFile: join(dir, 'allowlist.json'), appId: 'app-id', appPassword: 'pw',
        appType: 'SingleTenant', tenantId: TENANT_ID, bindHost: '127.0.0.1', port: 0, receivedFilesDir: dir,
      } as any,
      adapter: {} as any,
      allowlist,
      refs: createConversationRefStore(),
      onEvent: e => events.push(e),
      attachments: async (atts, fromName, messageText) =>
        processAttachments(atts, fromName, {
          baseDir: dir,
          getBearerToken: async () => { tokenRequests += 1; return TOKEN },
          messageText,
          fetchImpl,
          now: () => new Date('2026-11-20T06:00:00Z'),
          timeZone: 'UTC',
        }),
    })
    await turn(new TurnContext(new NoopAdapter() as any, activity))
  } finally {
    ;(process.stderr as any).write = write
  }
  return { events, calls, tokenRequests, log }
}

describe('measured failing shape: text/html holding only an <img>, no image/* part (1:1 chat)', () => {
  test('the image is downloaded through the inline-image path and the event names the saved file', async () => {
    const html = `<div><div><span><img height="250" src="${SMBA}" width="300" itemtype="http://schema.skype.com/AMSImage"></span></div></div>`
    const { events, calls } = await deliver(personal({ text: '', attachments: [{ contentType: 'text/html', content: html }] }))
    expect(events.length).toBe(1)
    expect(events[0].text.length).toBeGreaterThan(0)
    expect(events[0].text).toMatch(/^\[Attached: .*0600-owner-image\.png \(image\/png, 12 B\)\]$/)
    expect(calls).toEqual([{ url: SMBA, authorization: `Bearer ${TOKEN}` }])
  })

  test('the image type comes from the downloaded bytes (an <img> declares none)', async () => {
    const html = `<p><img alt="image" src="${SMBA}"></p>`
    const { events } = await deliver(personal({ text: '', attachments: [{ contentType: 'text/html', content: html }] }), JPEG)
    expect(events[0].text).toMatch(/image\.jpeg \(image\/jpeg, 12 B\)\]$/)
  })

  test('unknown bytes are saved as .bin, not labelled as an image type', async () => {
    const html = `<p><img src="${SMBA}"></p>`
    const { events } = await deliver(personal({ text: '', attachments: [{ contentType: 'text/html', content: html }] }), new Uint8Array([1, 2, 3]))
    expect(events[0].text).toMatch(/image\.bin \(application\/octet-stream, 3 B\)\]$/)
  })

  test('an image/* part next to the HTML copy is downloaded once, the <img> is not fetched again', async () => {
    const { events, calls } = await deliver(personal({
      text: '',
      attachments: [
        { contentType: 'image/png', contentUrl: SMBA },
        { contentType: 'text/html', content: `<div><img src="${SMBA}"></div>` },
      ],
    }))
    expect(calls.length).toBe(1)
    expect(events[0].text).toMatch(/^\[Attached: .*image\.png \(image\/png, 12 B\)\]$/)
  })
})

describe('the bot token never goes to a host outside the inline-image allowlist', () => {
  for (const src of [
    'https://attacker.example.com/x.png',
    'http://smba.trafficmanager.net/emea/v3/attachments/x/views/original',
    'https://evil.trafficmanager.net/v3/attachments/x',
    'https://smba.trafficmanager.net.attacker.example.com/x',
    'https://user:pw@smba.trafficmanager.net/x',
    'https://graph.microsoft.com/v1.0/chats/19:x/messages/1/hostedContents/aWQ=/$value',
  ]) {
    test(`<img src="${src}">: no request, no token, the empty-message note`, async () => {
      const { events, calls, tokenRequests } = await deliver(personal({
        text: '', attachments: [{ contentType: 'text/html', content: `<div><img src="${src}"></div>` }],
      }))
      expect(calls).toEqual([])
      expect(tokenRequests).toBe(0)
      expect(events.length).toBe(1)
      expect(events[0].text).toBe('[Teams message with no text: 1 attachment arrived (text/html), none could be read]')
    })
  }

  test('an entity-encoded foreign host is decoded before the check', async () => {
    const { calls, tokenRequests } = await deliver(personal({
      text: '', attachments: [{ contentType: 'text/html', content: '<img src="https&#58;//attacker.example.com/x.png">' }],
    }))
    expect(calls).toEqual([])
    expect(tokenRequests).toBe(0)
  })
})

describe('never a 0-character channel event', () => {
  test('empty text and every attachment ignored: a visible note naming count and content types', async () => {
    const { events } = await deliver(personal({ text: '', attachments: [{ contentType: 'text/html', content: '<div><p>&nbsp;</p></div>' }] }))
    expect(events.length).toBe(1)
    expect(events[0].text.length).toBeGreaterThan(0)
    expect(events[0].text).toBe('[Teams message with no text: 1 attachment arrived (text/html), none could be read]')
  })

  test('empty text and no attachment at all: a visible note, not 0 characters', async () => {
    const { events } = await deliver(personal({ text: '' }))
    expect(events.length).toBe(1)
    expect(events[0].text).toBe('[Teams message with no text and no attachment]')
  })

  test('no text field and no attachment: the same note', async () => {
    const { events } = await deliver(personal({}))
    expect(events[0].text).toBe('[Teams message with no text and no attachment]')
  })

  test('text present and the HTML copy ignored: the text passes through unchanged, no note', async () => {
    const { events } = await deliver(personal({ text: 'hello', attachments: [{ contentType: 'text/html', content: '<p>hello</p>' }] }))
    expect(events[0].text).toBe('hello')
  })
})

describe('one log line per inbound message activity: attachment count and content types only', () => {
  test('count and types, no URL, token or file name', async () => {
    const { log } = await deliver(personal({
      text: '',
      attachments: [
        { contentType: 'image/png', contentUrl: SMBA, name: 'secret-plan.png' },
        { contentType: 'text/html', content: `<div><img src="${SMBA}"></div>` },
        {
          contentType: 'application/vnd.microsoft.teams.file.download.info',
          name: 'secret-plan.pdf',
          content: { downloadUrl: 'https://contoso-my.sharepoint.com/personal/x/download.aspx?tempauth=TEMPAUTHVALUE', fileType: 'pdf' },
        },
        { contentType: 'https://attacker.example.com/as-a-type', contentUrl: 'https://attacker.example.com/y' },
      ],
    }))
    const lines = log.filter(l => l.startsWith('teams channel: inbound attachments'))
    expect(lines).toEqual([
      'teams channel: inbound attachments count=4 types=image/png,text/html,application/vnd.microsoft.teams.file.download.info,other\n',
    ])
    for (const bad of ['https://', 'http://', 'smba', 'sharepoint', 'attacker', 'TEMPAUTHVALUE', 'tempauth', 'secret-plan', TOKEN]) {
      expect(lines[0]).not.toContain(bad)
    }
  })

  test('a message with no attachment logs count=0', async () => {
    const { log } = await deliver(personal({ text: 'hi' }))
    expect(log.filter(l => l.startsWith('teams channel: inbound attachments'))).toEqual(['teams channel: inbound attachments count=0 types=none\n'])
  })
})

describe('htmlImageSources', () => {
  test('reads double, single and unquoted src, decodes entities, skips non-img tags', () => {
    expect(htmlImageSources(`<img src="https://a.test/1?x=1&amp;y=2"><IMG SRC='https://a.test/2'><img src=https://a.test/3><imgx src="https://a.test/4"><image src="https://a.test/5">`))
      .toEqual(['https://a.test/1?x=1&y=2', 'https://a.test/2', 'https://a.test/3'])
  })
  test('deduplicates and stops at MAX_HTML_IMAGES', () => {
    const many = Array.from({ length: MAX_HTML_IMAGES + 5 }, (_, i) => `<img src="https://a.test/${i}"><img src="https://a.test/${i}">`).join('')
    const got = htmlImageSources(many)
    expect(got.length).toBe(MAX_HTML_IMAGES)
    expect(new Set(got).size).toBe(MAX_HTML_IMAGES)
  })
  test('an unclosed <img is not read and does not hang', () => {
    const t0 = Date.now()
    expect(htmlImageSources('<img src="https://a.test/x"' + '<img '.repeat(50_000))).toEqual([])
    expect(Date.now() - t0).toBeLessThan(2000)
  })
})
