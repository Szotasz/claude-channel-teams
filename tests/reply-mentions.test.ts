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
      spy.mockImplementation(async () => ({ id: '29:from-getmember' }) as any)
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
    const s = setup({ id: 'c' }, { lookupMember: async (_c: unknown, id: string) => (seen.push(id), '29:m') })
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
    const s = setup({ id: 'c' }, { lookupMember: async () => '29:mri-a' })
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
    const s = setup({ id: 'c' }, { lookupMember: async () => '29:mri-a' })
    await s.sendReply('conv', 'Hi <at>Ann</at>', undefined, { mentions: [user], aiGenerated: true })
    expect(s.state.sent.entities).toHaveLength(2)
    expect(s.state.sent.entities[0].type).toBe('mention')
    expect(s.state.sent.entities[1].additionalType).toEqual(['AIGeneratedContent'])
  })
})
