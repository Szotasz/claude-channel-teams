import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { loadConfig, parseTeamIds } from '../src/config'
import {
  CHANNEL_TOOL,
  channelTools,
  createChannel,
  validateCreateChannelArgs,
} from '../src/teams/channels'

const T1 = '11111111-1111-4111-8111-111111111111'
const T2 = 'AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA'
const URL1 = `https://graph.microsoft.com/v1.0/teams/${T1}/channels`

interface Call {
  url: string
  method: string
  headers: Record<string, string>
  body: string | undefined
}

function makeFetch(answer: { status?: number; json?: unknown }) {
  const calls: Call[] = []
  const fn = (async (url: any, init: any) => {
    calls.push({ url: String(url), method: init?.method ?? 'GET', headers: { ...(init?.headers ?? {}) }, body: init?.body })
    return new Response(JSON.stringify(answer.json ?? {}), { status: answer.status ?? 200 })
  }) as unknown as typeof fetch
  return { fn, calls }
}

const input = { teamId: T1, name: 'General ideas', description: 'About ideas' }
const tokenFn = () => {
  let n = 0
  return { getToken: async () => { n++; return 'TOK' }, count: () => n }
}

describe('parseTeamIds', () => {
  test('valid', () => {
    expect([...parseTeamIds(`${T1}, ${T1.replace(/1/g, '2')}`)!]).toEqual([T1, T1.replace(/1/g, '2')])
  })
  test('mixed case is lower-cased', () => {
    expect([...parseTeamIds(T2)!]).toEqual([T2.toLowerCase()])
  })
  test('empty -> empty set', () => {
    expect(parseTeamIds(undefined)!.size).toBe(0)
    expect(parseTeamIds('')!.size).toBe(0)
  })
  test('one bad entry -> null', () => {
    expect(parseTeamIds(`${T1},nope`)).toBeNull()
  })
})

describe('validateCreateChannelArgs', () => {
  const one = new Set([T1])
  const two = new Set([T1, T2.toLowerCase()])
  const err = (a: Record<string, unknown>, s: ReadonlySet<string> = one) => {
    try {
      validateCreateChannelArgs(a, s)
    } catch (e) {
      return (e as Error).message
    }
    return undefined
  }

  test('error messages', () => {
    expect(err({ name: 'x' }, new Set())).toBe('create_channel is off (TEAMS_CHANNEL_CREATE_TEAMS is not set)')
    expect(err({ name: 'x' }, two)).toBe('team_id is required when more than one team is configured')
    expect(err({ name: 'x', team_id: null }, two)).toBe('team_id is required when more than one team is configured')
    expect(err({ name: 'x', team_id: 'other' })).toBe('team not allowed')
    expect(err({ name: 'x', team_id: 5 })).toBe('team not allowed')
    expect(err({})).toBe('name is required')
    expect(err({ name: 3 })).toBe('name is required')
    expect(err({ name: '   ' })).toBe('name is required')
    expect(err({ name: 'a'.repeat(51) })).toBe('name can be at most 50 characters')
    expect(err({ name: 'a\u0007b' })).toBe('name must not contain control characters')
    expect(err({ name: 'a', description: 5 })).toBe('description must be a string')
    expect(err({ name: 'a', description: 'a'.repeat(1025) })).toBe('description can be at most 1024 characters')
    expect(err({ name: 'a', description: 'a\u0001' })).toBe('description must not contain control characters other than tab and line breaks')
  })

  test('single-team default and case-insensitive team_id', () => {
    expect(validateCreateChannelArgs({ name: 'x' }, one).teamId).toBe(T1)
    expect(validateCreateChannelArgs({ name: 'x', team_id: T2 }, two).teamId).toBe(T2.toLowerCase())
  })

  test('name limits and trimming', () => {
    expect(validateCreateChannelArgs({ name: 'a'.repeat(50) }, one).name).toBe('a'.repeat(50))
    expect(validateCreateChannelArgs({ name: `  ${'a'.repeat(50)}  ` }, one).name).toBe('a'.repeat(50))
    expect(validateCreateChannelArgs({ name: ' x ' }, one).name).toBe('x')
  })

  test('C1 characters are refused', () => {
    expect(err({ name: 'a\u0085b' })).toBe('name must not contain control characters')
    expect(err({ name: 'a', description: 'a\u009fb' })).toBe('description must not contain control characters other than tab and line breaks')
  })

  test('DEL is refused', () => {
    expect(err({ name: 'a\u007fb' })).toBe('name must not contain control characters')
    expect(err({ name: 'a', description: 'a\u007fb' })).toBe('description must not contain control characters other than tab and line breaks')
  })

  test('description rules', () => {
    expect(validateCreateChannelArgs({ name: 'x', description: ' a\nb\tc ' }, one).description).toBe('a\nb\tc')
    expect(validateCreateChannelArgs({ name: 'x', description: 'a'.repeat(1024) }, one).description).toBe('a'.repeat(1024))
    expect('description' in validateCreateChannelArgs({ name: 'x', description: '  ' }, one)).toBe(false)
    expect('description' in validateCreateChannelArgs({ name: 'x', description: '' }, one)).toBe(false)
  })
})

describe('createChannel', () => {
  test('request shape and result', async () => {
    const f = makeFetch({ status: 201, json: { id: '19:abc@thread.tacv2', webUrl: 'https://teams.microsoft.com/x' } })
    const t = tokenFn()
    const r = await createChannel(input, { getToken: t.getToken, fetch: f.fn })
    expect(r).toEqual({ channelId: '19:abc@thread.tacv2', webUrl: 'https://teams.microsoft.com/x' })
    expect(f.calls.length).toBe(1)
    expect(t.count()).toBe(1)
    expect(f.calls[0]!.url).toBe(URL1)
    expect(f.calls[0]!.method).toBe('POST')
    expect(f.calls[0]!.headers.Authorization).toBe('Bearer TOK')
    expect(JSON.parse(f.calls[0]!.body!)).toEqual({
      displayName: 'General ideas',
      description: 'About ideas',
      membershipType: 'standard',
    })
  })

  test('no description in the body when none', async () => {
    const f = makeFetch({ status: 201, json: { id: '19:abc' } })
    const r = await createChannel({ teamId: T1, name: 'N' }, { getToken: async () => 'TOK', fetch: f.fn })
    expect(r).toEqual({ channelId: '19:abc' })
    expect(JSON.parse(f.calls[0]!.body!)).toEqual({ displayName: 'N', membershipType: 'standard' })
  })

  test('non-https webUrl dropped', async () => {
    const f = makeFetch({ status: 201, json: { id: '19:abc', webUrl: 'http://x' } })
    expect(await createChannel(input, { getToken: async () => 'TOK', fetch: f.fn })).toEqual({ channelId: '19:abc' })
  })

  test('no id', async () => {
    const f = makeFetch({ status: 201, json: { displayName: 'x' } })
    await expect(createChannel(input, { getToken: async () => 'TOK', fetch: f.fn })).rejects.toThrow(
      'Graph returned no channel id',
    )
  })

  test('403', async () => {
    const f = makeFetch({ status: 403, json: { error: { code: 'Forbidden' } } })
    await expect(createChannel(input, { getToken: async () => 'TOK', fetch: f.fn })).rejects.toThrow(
      'Graph HTTP 403 Forbidden',
    )
  })

})

describe('description line breaks and length', () => {
  const one = new Set([T1])
  test('tab, LF and CR are allowed; other control characters are not', () => {
    expect(validateCreateChannelArgs({ name: 'x', description: 'a\r\nb\tc' }, one).description).toBe('a\r\nb\tc')
    expect(() => validateCreateChannelArgs({ name: 'x', description: 'a\fb' }, one)).toThrow(
      'description must not contain control characters other than tab and line breaks',
    )
    expect(() => validateCreateChannelArgs({ name: 'x', description: 'a\u000bb' }, one)).toThrow(
      'description must not contain control characters other than tab and line breaks',
    )
  })
  test('the length is counted after trimming', () => {
    expect(validateCreateChannelArgs({ name: 'x', description: `  ${'d'.repeat(1024)}  ` }, one).description).toBe(
      'd'.repeat(1024),
    )
  })
})

describe('channelTools', () => {
  test('empty and non-empty', () => {
    expect(channelTools(new Set())).toEqual([])
    expect(channelTools(new Set([T1]))).toEqual([CHANNEL_TOOL])
  })
})

describe('loadConfig channelCreateTeams', () => {
  const KEYS = ['TEAMS_CHANNEL_CREATE_TEAMS', 'TEAMS_PLUGIN_STATE_DIR', 'TEAMS_BOT_APP_ID', 'TEAMS_BOT_APP_PASSWORD', 'TEAMS_BOT_TENANT_ID']
  const saved: Record<string, string | undefined> = {}
  let dir: string
  beforeEach(() => {
    for (const k of KEYS) saved[k] = process.env[k]
    dir = mkdtempSync(join(tmpdir(), 'channels-cfg-'))
    process.env.TEAMS_PLUGIN_STATE_DIR = dir
    process.env.TEAMS_BOT_APP_ID = 'a'
    process.env.TEAMS_BOT_APP_PASSWORD = 'b'
    process.env.TEAMS_BOT_TENANT_ID = 'c'
  })
  afterEach(() => {
    for (const k of KEYS) {
      if (saved[k] === undefined) delete process.env[k]
      else process.env[k] = saved[k]
    }
    rmSync(dir, { recursive: true, force: true })
  })

  test('valid', () => {
    process.env.TEAMS_CHANNEL_CREATE_TEAMS = `${T1},${T2}`
    expect([...loadConfig().channelCreateTeams]).toEqual([T1, T2.toLowerCase()])
  })

  test('invalid -> empty set and a stderr line', () => {
    process.env.TEAMS_CHANNEL_CREATE_TEAMS = 'bad'
    const lines: string[] = []
    const orig = process.stderr.write
    process.stderr.write = ((s: any) => { lines.push(String(s)); return true }) as any
    try {
      expect(loadConfig().channelCreateTeams.size).toBe(0)
    } finally {
      process.stderr.write = orig
    }
    expect(lines).toContain('teams channel: TEAMS_CHANNEL_CREATE_TEAMS has an invalid team id; create_channel stays off\n')
  })
})
