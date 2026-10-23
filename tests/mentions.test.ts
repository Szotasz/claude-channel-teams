import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import {
  parseMentionTags,
  validateMentions,
  mentionEntity,
  isChannelReference,
  createMemberLookup,
  isUsableDisplayName,
} from '../src/teams/mentions'
import { loadConfig } from '../src/config'

const AAD = 'ABCDEF01-0000-4000-8000-0000000000AB'
const tags = new Map([['Finance', 'QUJD']])
const v = (raw: unknown, text = '<at>Ann</at> <at>Finance</at>', t = tags, ch = true) =>
  validateMentions(raw, text, t, ch)
const user = (extra = {}) => ({ type: 'user', aad_object_id: AAD, name: 'Ann', ...extra })

describe('parseMentionTags', () => {
  test('empty', () => {
    expect(parseMentionTags(undefined)).toEqual(new Map())
    expect(parseMentionTags('  ')).toEqual(new Map())
  })
  test('valid, ordered, trimmed', () => {
    expect([...parseMentionTags('Finance=QUJD,Legal=REVG==')!]).toEqual([['Finance', 'QUJD'], ['Legal', 'REVG==']])
    expect([...parseMentionTags(' Finance = QUJD ')!]).toEqual([['Finance', 'QUJD']])
  })
  test('bad entries give null', () => {
    for (const s of ['Finance', 'Fin<x>=QUJD', 'Finance=not base64!', 'Finance=QUJD,Finance=REVG']) {
      expect(parseMentionTags(s)).toBeNull()
    }
  })
})

describe('validateMentions', () => {
  test('errors', () => {
    expect(() => v('x')).toThrow('mentions must be an array')
    expect(() => v([])).toThrow('mentions must not be empty')
    expect(() => v(Array(11).fill(user()))).toThrow('mentions can list at most 10')
    expect(() => v([null])).toThrow('mentions[0] must be an object')
    expect(() => v([[]])).toThrow('mentions[0] must be an object')
    expect(() => v([{ type: 'x' }])).toThrow('mentions[0].type must be "user" or "tag"')
    expect(() => v([user({ aad_object_id: 'nope' })])).toThrow('mentions[0].aad_object_id must be an AAD object id')
    expect(() => v([user({ name: 'A<b' })])).toThrow('mentions[0].name must be 1-100 characters without < > , =')
    expect(() => v([user({ name: 'A\u0085B' })], '<at>A\u0085B</at>')).toThrow(
      'mentions[0].name must be 1-100 characters without < > , =',
    )
    expect(() => v([user({ name: ' Ann' })])).toThrow('mentions[0].name must be 1-100 characters without < > , =')
    expect(() => v([{ type: 'tag', tag: 'Finance' }], 'x', new Map())).toThrow(
      'tag mentions are not enabled (set TEAMS_MENTION_TAGS)',
    )
    expect(() => v([{ type: 'tag', tag: 'finance' }])).toThrow('mentions[0].tag is not in TEAMS_MENTION_TAGS')
    expect(() => v([{ type: 'tag' }])).toThrow('mentions[0].tag is not in TEAMS_MENTION_TAGS')
    expect(() => v([{ type: 'tag', tag: 'Finance' }], 'x', tags, false)).toThrow(
      'tag mentions work only in a team channel',
    )
    const t = { type: 'tag', tag: 'Finance' }
    expect(() => v([t, t])).toThrow('mentions[1] is listed twice')
    expect(() => v([user()], 'hi')).toThrow('mentions[0]: the text must contain <at>Ann</at>')
    expect(() => v([t], 'hi')).toThrow('mentions[0]: the text must contain <at>Finance</at>')
  })
  test('limits: 10 mentions and a 100-character name are accepted; a comma or an equals sign is not', () => {
    const ten = Array.from({ length: 10 }, (_, i) =>
      user({ aad_object_id: `abcdef01-0000-4000-8000-0000000000${String(i).padStart(2, '0')}`, name: `P${i}` }),
    )
    const text = ten.map(m => `<at>${m.name}</at>`).join(' ')
    expect(v(ten, text).length).toBe(10)
    const long = 'A'.repeat(100)
    expect(v([user({ name: long })], `<at>${long}</at>`).length).toBe(1)
    expect(() => v([user({ name: `${long}A` })], `<at>${long}A</at>`)).toThrow(
      'mentions[0].name must be 1-100 characters without < > , =',
    )
    expect(() => v([user({ name: 'Ann, Bob' })], '<at>Ann, Bob</at>')).toThrow(
      'mentions[0].name must be 1-100 characters without < > , =',
    )
    expect(() => v([user({ name: 'A=B' })], '<at>A=B</at>')).toThrow(
      'mentions[0].name must be 1-100 characters without < > , =',
    )
  })
  test('two entries may not use the same name', () => {
    const other = user({ aad_object_id: 'abcdef01-0000-4000-8000-0000000000cd' })
    expect(() => v([user(), other])).toThrow('mentions[1] uses the same name as an earlier entry')
    expect(() => v([user({ name: 'Finance' }), { type: 'tag', tag: 'Finance' }], '<at>Finance</at>')).toThrow(
      'mentions[1] uses the same name as an earlier entry',
    )
  })
  test('duplicate users in both case orders', () => {
    expect(() => v([user(), user({ aad_object_id: AAD.toLowerCase() })])).toThrow('mentions[1] is listed twice')
    expect(() => v([user({ aad_object_id: AAD.toLowerCase() }), user()])).toThrow('mentions[1] is listed twice')
  })
  test('valid mixed list', () => {
    expect(v([user(), { type: 'tag', tag: 'Finance' }])).toEqual([
      { type: 'user', aadObjectId: AAD.toLowerCase(), name: 'Ann' },
      { type: 'tag', name: 'Finance', id: 'QUJD' },
    ])
  })
})

test('mentionEntity', () => {
  expect(mentionEntity('Ann', '29:a', false)).toEqual({
    type: 'mention',
    text: '<at>Ann</at>',
    mentioned: { id: '29:a', name: 'Ann' },
  })
  expect(mentionEntity('Finance', 'QUJD', true)).toEqual({
    type: 'mention',
    text: '<at>Finance</at>',
    mentioned: { id: 'QUJD', name: 'Finance', type: 'tag' },
  })
})

test('isChannelReference', () => {
  expect(isChannelReference({ conversation: { conversationType: 'channel' } })).toBe(true)
  expect(isChannelReference({ conversation: { id: '19:abc@thread.tacv2;messageid=1' } })).toBe(true)
  expect(isChannelReference({ conversation: { id: 'a;b', conversationType: 'personal' } })).toBe(false)
  expect(isChannelReference({ conversation: { id: '19:abc@thread.v2', conversationType: 'groupChat' } })).toBe(false)
  expect(isChannelReference(undefined)).toBe(false)
})

describe('createMemberLookup', () => {
  const ctx = {} as any
  const lookup = (f: () => Promise<any>) => createMemberLookup(f)
  test('member id', async () => {
    expect(await lookup(async () => ({ id: 'x', name: 'Ann Smith' }))(ctx, AAD)).toEqual({ id: 'x', name: 'Ann Smith' })
    expect(await lookup(async () => ({ id: 'x' }))(ctx, AAD)).toEqual({ id: 'x', name: undefined })
  })
  test('not member errors', async () => {
    for (const e of [
      { statusCode: 404 },
      { code: 'MemberNotFoundInConversation' },
      { body: { error: { code: 'MemberNotFoundInConversation' } } },
    ]) {
      expect(await lookup(async () => { throw e })(ctx, AAD)).toBeUndefined()
    }
  })
  test('other errors rethrown', async () => {
    const e = { statusCode: 500 }
    await expect(lookup(async () => { throw e })(ctx, AAD)).rejects.toBe(e)
  })
  test('no id', async () => {
    expect(await lookup(async () => ({}))(ctx, AAD)).toBeUndefined()
  })
})

test('isUsableDisplayName', () => {
  for (const n of ['Ann Smith', 'Søren Åberg', 'Ann, Sales', 'A'.repeat(256)]) expect(isUsableDisplayName(n)).toBe(true)
  for (const n of ['', ' Ann', 'A<b', 'A>b', 'A\nB', 'A\u0085B', 'A\u009fB', 'A'.repeat(257), 5, undefined]) {
    expect(isUsableDisplayName(n)).toBe(false)
  }
})

describe('loadConfig mentionTags', () => {
  const KEYS = ['TEAMS_BOT_APP_ID', 'TEAMS_BOT_APP_PASSWORD', 'TEAMS_BOT_TENANT_ID', 'TEAMS_MENTION_TAGS', 'TEAMS_STATE_DIR']
  const lines: string[] = []
  const warning = 'teams channel: TEAMS_MENTION_TAGS has an invalid entry; tag mentions stay off\n'
  function withEnv(tagsVal: string, fn: (c: ReturnType<typeof loadConfig>) => void) {
    const saved = Object.fromEntries(KEYS.map(k => [k, process.env[k]]))
    const w = process.stderr.write
    try {
      process.env.TEAMS_BOT_APP_ID = 'a'
      process.env.TEAMS_BOT_APP_PASSWORD = 'p'
      process.env.TEAMS_BOT_TENANT_ID = 't'
      process.env.TEAMS_STATE_DIR = mkdtempSync(join(tmpdir(), 'cct-ment-'))
      process.env.TEAMS_MENTION_TAGS = tagsVal
      lines.length = 0
      process.stderr.write = ((s: string) => (lines.push(String(s)), true)) as any
      fn(loadConfig())
    } finally {
      process.stderr.write = w
      for (const k of KEYS) {
        if (saved[k] === undefined) delete process.env[k]
        else process.env[k] = saved[k]
      }
    }
  }
  test('valid', () =>
    withEnv('Finance=QUJD', c => {
      expect(c.mentionTags.get('Finance')).toBe('QUJD')
      expect(lines.filter(l => l === warning)).toEqual([])
    }))
  test('invalid', () =>
    withEnv('Finance', c => {
      expect(c.mentionTags.size).toBe(0)
      expect(lines.filter(l => l === warning)).toEqual([warning])
    }))
})
