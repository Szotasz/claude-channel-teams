import { describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { loadConfig } from '../src/config.js'
import {
  ACTIVITY_TOOL,
  activityTools,
  sendActivityNotification,
  validateActivityArgs,
} from '../src/teams/activityFeed.js'

const GUID = 'AAAAAAAA-bbbb-4ccc-8ddd-EEEEEEEEEEEE'
const ok = { conversation_id: '19:abc@thread.tacv2;messageid=1700000000000', aad_object_id: GUID, text: 'hi' }
const bad = (over: Record<string, unknown>) => () => validateActivityArgs({ ...ok, ...over })

describe('validateActivityArgs', () => {
  test('conversation_id required', () => {
    expect(bad({ conversation_id: '' })).toThrow('conversation_id is required')
    expect(bad({ conversation_id: 5 })).toThrow('conversation_id is required')
  })
  test('only team channels', () => {
    expect(bad({ conversation_id: '19:x@thread.v2' })).toThrow('activity notifications are only available in team channels')
    expect(bad({ conversation_id: 'a:1xyz' })).toThrow('activity notifications are only available in team channels')
  })
  test('empty messageid', () => {
    expect(bad({ conversation_id: '19:abc@thread.tacv2;messageid=' })).toThrow('invalid conversation_id')
  })
  test('aad_object_id must be a GUID', () => {
    expect(bad({ aad_object_id: 'nope' })).toThrow('aad_object_id must be an AAD object id')
    expect(bad({ aad_object_id: undefined })).toThrow('aad_object_id must be an AAD object id')
  })
  test('text required', () => {
    expect(bad({ text: '   ' })).toThrow('text is required')
    expect(bad({ text: 3 })).toThrow('text is required')
  })
  test('150/151 boundary after trim', () => {
    expect(validateActivityArgs({ ...ok, text: ' ' + 'a'.repeat(150) + ' ' }).text).toBe('a'.repeat(150))
    expect(bad({ text: 'a'.repeat(151) })).toThrow('text must be at most 150 characters')
  })
  test('lower-cases the GUID and trims text', () => {
    const v = validateActivityArgs({ ...ok, text: '  hi  ' })
    expect(v.userId).toBe(GUID.toLowerCase())
    expect(v.text).toBe('hi')
  })
  test('parses channel and message id', () => {
    const t = validateActivityArgs(ok)
    expect(t.channelId).toBe('19:abc@thread.tacv2')
    expect(t.messageId).toBe('1700000000000')
    const c = validateActivityArgs({ ...ok, conversation_id: '19:abc@thread.tacv2' })
    expect(c.channelId).toBe('19:abc@thread.tacv2')
    expect(c.messageId).toBeNull()
  })
})

type Call = { url: string; init: RequestInit }
function fake(res: () => Response | Promise<Response>) {
  const calls: Call[] = []
  const fetch = async (url: string, init?: RequestInit) => {
    calls.push({ url, init: init ?? {} })
    return res()
  }
  return { calls, fetch }
}
const base = { teamId: 'team-1', channelId: '19:abc@thread.tacv2', messageId: '1700000000000', userId: GUID.toLowerCase(), text: 'look' }
const send = (f: ReturnType<typeof fake>, input = base) =>
  sendActivityNotification(input, { getToken: async () => 'TOK', fetch: f.fetch })
const expectedBody = (value: string) => ({
  topic: { source: 'entityUrl', value },
  activityType: 'systemDefault',
  previewText: { content: 'look' },
  recipient: { '@odata.type': 'microsoft.graph.aadUserNotificationRecipient', userId: GUID.toLowerCase() },
  templateParameters: [{ name: 'systemDefaultText', value: 'look' }],
})

describe('sendActivityNotification', () => {
  test('thread case posts the message URL', async () => {
    const f = fake(() => new Response(null, { status: 204 }))
    await send(f)
    const c = f.calls[0]!
    expect(c.url).toBe('https://graph.microsoft.com/v1.0/teams/team-1/sendActivityNotification')
    expect(c.init.method).toBe('POST')
    expect(c.init.redirect).toBe('error')
    const h = c.init.headers as Record<string, string>
    expect(h.Authorization).toBe('Bearer TOK')
    expect(h['content-type']).toBe('application/json')
    expect(JSON.parse(c.init.body as string)).toEqual(
      expectedBody('https://graph.microsoft.com/v1.0/teams/team-1/channels/19%3Aabc%40thread.tacv2/messages/1700000000000'),
    )
  })
  test('channel case posts the channel URL', async () => {
    const f = fake(() => new Response(null, { status: 204 }))
    await send(f, { ...base, messageId: null })
    expect(JSON.parse(f.calls[0]!.init.body as string)).toEqual(
      expectedBody('https://graph.microsoft.com/v1.0/teams/team-1/channels/19%3Aabc%40thread.tacv2'),
    )
  })
  test('encodes team id in the URL', async () => {
    const f = fake(() => new Response(null, { status: 204 }))
    await send(f, { ...base, teamId: 'a/b c' })
    expect(f.calls[0]!.url).toBe('https://graph.microsoft.com/v1.0/teams/a%2Fb%20c/sendActivityNotification')
  })
  test('403 with a safe code', async () => {
    const f = fake(() => Response.json({ error: { code: 'Forbidden', message: 'leak SECRET' } }, { status: 403 }))
    const e = await send(f).catch(x => x as Error)
    expect((e as Error).message).toBe('Graph HTTP 403 Forbidden')
    expect((e as Error).message).not.toContain('SECRET')
  })
  test('unsafe code is dropped', async () => {
    const f = fake(() => Response.json({ error: { code: 'bad code!' } }, { status: 403 }))
    await expect(send(f)).rejects.toThrow(/^Graph HTTP 403$/)
  })
  test('non-JSON error body', async () => {
    const f = fake(() => new Response('oops', { status: 500 }))
    await expect(send(f)).rejects.toThrow(/^Graph HTTP 500$/)
  })
  test('fetch throwing', async () => {
    const f = fake(() => {
      throw new Error('boom https://graph.microsoft.com/secret')
    })
    await expect(send(f)).rejects.toThrow('Graph request failed (network error)')
    const g = fake(() => {
      throw Object.assign(new Error('x'), { code: 'ECONNRESET' })
    })
    await expect(send(g)).rejects.toThrow('Graph request failed (ECONNRESET)')
  })
})

describe('activityTools', () => {
  test('gated by the flag', () => {
    expect(activityTools(false)).toEqual([])
    expect(activityTools(true)).toEqual([ACTIVITY_TOOL])
    expect(ACTIVITY_TOOL.inputSchema.required).toEqual(['conversation_id', 'aad_object_id', 'text'])
  })
})

describe('loadConfig activityFeed', () => {
  // loadConfig reads process.env after loading <stateDir>/.env, so the keys are set there and restored afterwards.
  test('true only for TEAMS_ACTIVITY_FEED exactly 1', () => {
    const dir = mkdtempSync(join(tmpdir(), 'activity-cfg-'))
    const keys = ['TEAMS_BOT_APP_ID', 'TEAMS_BOT_APP_PASSWORD', 'TEAMS_BOT_TENANT_ID', 'TEAMS_ACTIVITY_FEED'] as const
    const saved = keys.map(k => [k, process.env[k]] as const)
    try {
      process.env.TEAMS_BOT_APP_ID = 'app'
      process.env.TEAMS_BOT_APP_PASSWORD = 'pw'
      process.env.TEAMS_BOT_TENANT_ID = 'tenant'
      const feed = (v?: string) => {
        if (v === undefined) delete process.env.TEAMS_ACTIVITY_FEED
        else process.env.TEAMS_ACTIVITY_FEED = v
        return loadConfig({ TEAMS_PLUGIN_STATE_DIR: dir }).activityFeed
      }
      expect(feed()).toBe(false)
      expect(feed('1')).toBe(true)
      for (const v of ['true', '0', ' 1', 'yes', '']) expect(feed(v)).toBe(false)
    } finally {
      for (const [k, v] of saved) {
        if (v === undefined) delete process.env[k]
        else process.env[k] = v
      }
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
