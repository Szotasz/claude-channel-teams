import { describe, expect, test, beforeEach, spyOn } from 'bun:test'
import { TeamsInfo } from 'botbuilder'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createReplySender } from '../src/teams/reply'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const AAD = 'abcdef01-0000-4000-8000-0000000000ab'

const config = {
  stateDir: '/tmp/unused',
  allowlistFile: '/tmp/unused',
  appId: 'app-id-from-config',
  appPassword: 'app-pw',
  appType: 'SingleTenant' as const,
  tenantId: 't',
  bindHost: '127.0.0.1',
  port: 0,
}

describe('reply with mentions', () => {
  let dir: string
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cct-replyment-'))
  })

  function setup(conversation: any, extra: Record<string, unknown> = {}) {
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(FIXTURE_ID)
    const refs = createConversationRefStore()
    refs.put('conv', { conversation } as any, FIXTURE_ID)
    const state = { calls: 0, sends: 0, sent: undefined as any }
    const adapter = {
      async continueConversationAsync(_a: string, _r: any, logic: any) {
        state.calls++
        await logic({
          async sendActivity(act: any) {
            state.sends++
            state.sent = act
            return { id: 'x' }
          },
        })
      },
    } as any
    const sender = createReplySender({ config, adapter, allowlist, refs, ...extra } as any)
    return { ...sender, state }
  }
  const user = { type: 'user', aad_object_id: AAD, name: 'Ann' }
  const tag = { type: 'tag', tag: 'Finance' }
  const channel = { id: '19:x@thread.tacv2', conversationType: 'channel' }
  const tagsMap = new Map([['Finance', 'QUJD']])

  test('the default lookup is TeamsInfo.getMember: a 404 refuses, a member gets its id', async () => {
    const spy = spyOn(TeamsInfo, 'getMember').mockImplementation(async () => {
      throw { statusCode: 404 }
    })
    try {
      const s = setup({ id: 'c' })
      await expect(s.sendReply('conv', 'Hi <at>Ann</at>', undefined, { mentions: [user] })).rejects.toThrow(
        'mentions: not a member of this conversation: Ann',
      )
      expect(s.state.sends).toBe(0)
      spy.mockImplementation(async () => ({ id: '29:from-getmember', name: 'Ann' }) as any)
      await s.sendReply('conv', 'Hi <at>Ann</at>', undefined, { mentions: [user] })
      expect(s.state.sent.entities).toEqual([
        { type: 'mention', text: '<at>Ann</at>', mentioned: { id: '29:from-getmember', name: 'Ann' } },
      ])
    } finally {
      spy.mockRestore()
    }
  })

  test('the lookup gets the lower-cased AAD id', async () => {
    const seen: string[] = []
    const s = setup({ id: 'c' }, { lookupMember: async (_c: unknown, id: string) => (seen.push(id), { id: '29:m', name: 'Ann' }) })
    await s.sendReply('conv', 'Hi <at>Ann</at>', undefined, { mentions: [{ ...user, aad_object_id: AAD.toUpperCase() }] })
    expect(seen).toEqual([AAD])
  })

  test('a non-member next to a valid tag: nothing is sent', async () => {
    const s = setup(channel, { lookupMember: async () => undefined, mentionTags: tagsMap })
    await expect(
      s.sendReply('conv', '<at>Ann</at> <at>Finance</at>', undefined, { mentions: [user, tag] }),
    ).rejects.toThrow('mentions: not a member of this conversation: Ann')
    expect(s.state.sends).toBe(0)
  })

  test('member person', async () => {
    const s = setup({ id: 'c' }, { lookupMember: async () => ({ id: '29:mri-a', name: 'Ann' }) })
    await s.sendReply('conv', 'Hi <at>Ann</at>', undefined, { mentions: [user] })
    expect(s.state.sent.entities).toEqual([
      { type: 'mention', text: '<at>Ann</at>', mentioned: { id: '29:mri-a', name: 'Ann' } },
    ])
  })

  test('non-member person', async () => {
    const s = setup({ id: 'c' }, { lookupMember: async () => undefined })
    await expect(s.sendReply('conv', 'Hi <at>Ann</at>', undefined, { mentions: [user] })).rejects.toThrow(
      new Error('mentions: not a member of this conversation: Ann'),
    )
    expect(s.state.sends).toBe(0)
  })

  test('lookup error', async () => {
    const err = new Error('boom')
    const s = setup({ id: 'c' }, { lookupMember: async () => { throw err } })
    await expect(s.sendReply('conv', 'Hi <at>Ann</at>', undefined, { mentions: [user] })).rejects.toBe(err)
    expect(s.state.sends).toBe(0)
  })

  test('tag in channel', async () => {
    const s = setup(channel, { mentionTags: tagsMap })
    await s.sendReply('conv', 'Hi <at>Finance</at>', undefined, { mentions: [tag] })
    expect(s.state.sent.entities).toEqual([
      { type: 'mention', text: '<at>Finance</at>', mentioned: { id: 'QUJD', name: 'Finance', type: 'tag' } },
    ])
  })

  test('tag in personal conversation', async () => {
    const s = setup({ id: 'c', conversationType: 'personal' }, { mentionTags: tagsMap })
    await expect(s.sendReply('conv', 'Hi <at>Finance</at>', undefined, { mentions: [tag] })).rejects.toThrow(
      'tag mentions work only in a team channel',
    )
    expect(s.state.calls).toBe(0)
  })

  test('no mentions: no entities key', async () => {
    const s = setup({ id: 'c' })
    await s.sendReply('conv', 'hi')
    expect(s.state.sent).toEqual({ type: 'message', text: 'hi', textFormat: 'markdown' })
  })

  test('mention first, then AI entity', async () => {
    const s = setup({ id: 'c' }, { lookupMember: async () => ({ id: '29:mri-a', name: 'Ann' }) })
    await s.sendReply('conv', 'Hi <at>Ann</at>', undefined, { mentions: [user], aiGenerated: true })
    expect(s.state.sent.entities).toHaveLength(2)
    expect(s.state.sent.entities[0].type).toBe('mention')
    expect(s.state.sent.entities[1].additionalType).toEqual(['AIGeneratedContent'])
  })

  const people = [
    { type: 'user', aad_object_id: AAD, name: 'Ann' },
    { type: 'user', aad_object_id: 'abcdef01-0000-4000-8000-0000000000cd', name: 'Bob' },
  ]

  test('the shown name is the looked-up display name', async () => {
    const s = setup({ id: 'c' }, { lookupMember: async () => ({ id: '29:mri-a', name: 'Ann Smith' }) })
    await s.sendReply('conv', 'Hi <at>Security Team</at>', undefined, {
      mentions: [{ type: 'user', aad_object_id: AAD, name: 'Security Team' }],
    })
    expect(s.state.sent.text).toBe('Hi <at>Ann Smith</at>')
    expect(s.state.sent.entities).toEqual([
      { type: 'mention', text: '<at>Ann Smith</at>', mentioned: { id: '29:mri-a', name: 'Ann Smith' } },
    ])
  })

  test('no usable display name', async () => {
    for (const name of [undefined, 'A<b>']) {
      const s = setup({ id: 'c' }, { lookupMember: async () => ({ id: '29:m', name }) })
      await expect(s.sendReply('conv', 'Hi <at>Ann</at>', undefined, { mentions: [user] })).rejects.toThrow(
        new Error('mentions: Teams returned no usable display name for Ann'),
      )
      expect(s.state.sends).toBe(0)
    }
  })

  test('two people with the same looked-up name', async () => {
    const s = setup({ id: 'c' }, { lookupMember: async () => ({ id: '29:m', name: 'Ann Smith' }) })
    await expect(
      s.sendReply('conv', '<at>Ann</at> <at>Bob</at>', undefined, { mentions: people }),
    ).rejects.toThrow(new Error('mentions: two entries would show the same name: Ann Smith'))
    expect(s.state.sends).toBe(0)
  })

  test('two looked-up names that differ only in letter case', async () => {
    let n = 0
    const s = setup({ id: 'c' }, { lookupMember: async () => ({ id: `29:m${n}`, name: n++ === 0 ? 'Ann Smith' : 'ann smith' }) })
    await expect(
      s.sendReply('conv', '<at>Ann</at> <at>Bob</at>', undefined, { mentions: people }),
    ).rejects.toThrow(new Error('mentions: two entries would show the same name: ann smith'))
    expect(s.state.sends).toBe(0)
  })

  test('every place of a name is renamed, and $ patterns in a name stay literal', async () => {
    const s = setup({ id: 'c' }, { lookupMember: async () => ({ id: '29:m', name: 'A $& B $1' }) })
    await s.sendReply('conv', '<at>Ann</at>, again <at>Ann</at>', undefined, { mentions: [user] })
    expect(s.state.sent.text).toBe('<at>A $& B $1</at>, again <at>A $& B $1</at>')
  })

  test('a C1 control character in the looked-up name is not usable', async () => {
    const s = setup({ id: 'c' }, { lookupMember: async () => ({ id: '29:m', name: 'A\u0085B' }) })
    await expect(s.sendReply('conv', 'Hi <at>Ann</at>', undefined, { mentions: [user] })).rejects.toThrow(
      new Error('mentions: Teams returned no usable display name for Ann'),
    )
    expect(s.state.sends).toBe(0)
  })

  test('a looked-up name equal to a tag name', async () => {
    const s = setup(channel, { lookupMember: async () => ({ id: '29:m', name: 'Finance' }), mentionTags: tagsMap })
    await expect(
      s.sendReply('conv', '<at>Ann</at> <at>Finance</at>', undefined, { mentions: [user, tag] }),
    ).rejects.toThrow(new Error('mentions: two entries would show the same name: Finance'))
    expect(s.state.sends).toBe(0)
  })

  test('renames do not cascade', async () => {
    const names: Record<string, string> = { [AAD]: 'Ann', 'abcdef01-0000-4000-8000-0000000000cd': 'Carl' }
    const s = setup(
      { id: 'c' },
      { lookupMember: async (_c: unknown, id: string) => ({ id: `29:${id}`, name: names[id] }) },
    )
    await s.sendReply('conv', '<at>Bob</at> and <at>Ann</at>', undefined, {
      mentions: [
        { type: 'user', aad_object_id: AAD, name: 'Bob' },
        { type: 'user', aad_object_id: 'abcdef01-0000-4000-8000-0000000000cd', name: 'Ann' },
      ],
    })
    expect(s.state.sent.text).toBe('<at>Ann</at> and <at>Carl</at>')
  })
})
