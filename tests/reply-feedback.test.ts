import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createReplySender } from '../src/teams/reply'
import { createStatusStreams } from '../src/teams/statusStream'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'

function setup(opts: { feedback?: boolean; streams?: boolean; failFinal?: unknown } = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-replyfb-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(FIXTURE_ID)
  const refs = createConversationRefStore()
  refs.put('dm', { conversation: { id: 'a:1', conversationType: 'personal' } } as any, FIXTURE_ID)
  const sent: any[] = []
  const adapter = {
    async continueConversationAsync(_a: string, _r: any, logic: any) {
      await logic({
        async sendActivity(act: any) {
          sent.push(act)
          if (act.type === 'message' && opts.failFinal && act.entities?.some((e: any) => e.type === 'streaminfo')) throw opts.failFinal
          return { id: 's-1' }
        },
      })
    },
  } as any
  const statusStreams = opts.streams ? createStatusStreams({ sleep: async () => {} }) : undefined
  const config = {
    stateDir: '/tmp/unused', allowlistFile: '/tmp/unused', appId: 'app', appPassword: 'pw',
    appType: 'SingleTenant' as const, tenantId: TENANT_ID, bindHost: '127.0.0.1', port: 0,
  }
  const sender = createReplySender({
    config, adapter, allowlist, refs, typingPump: { stop() {} } as any, statusStreams,
    ...(opts.feedback !== undefined ? { feedback: opts.feedback } : {}),
  })
  return { ...sender, sent, statusStreams }
}

const messages = (s: { sent: any[] }) => s.sent.filter(a => a.type === 'message')

describe('feedback buttons on replies', () => {
  test('on: channelData is exactly the feedback loop', async () => {
    const s = setup({ feedback: true })
    await s.sendReply('dm', 'hi')
    expect(messages(s).at(-1).channelData).toEqual({ feedbackLoop: { type: 'default' } })
  })

  test('every reply gets its own channelData object', async () => {
    const s = setup({ feedback: true })
    await s.sendReply('dm', 'one')
    await s.sendReply('dm', 'two')
    const [a, b] = messages(s).slice(-2)
    expect(a.channelData).toEqual(b.channelData)
    expect(a.channelData).not.toBe(b.channelData)
    expect(a.channelData.feedbackLoop).not.toBe(b.channelData.feedbackLoop)
  })

  test('off: no channelData key', async () => {
    const s = setup({ feedback: false })
    await s.sendReply('dm', 'hi')
    expect('channelData' in messages(s).at(-1)).toBe(false)
  })

  test('open stream: final message has channelData and the final streaminfo', async () => {
    const s = setup({ feedback: true, streams: true })
    await s.sendStatus('dm', 'working')
    await s.sendReply('dm', 'done')
    const last = messages(s).at(-1)
    expect(last.channelData).toEqual({ feedbackLoop: { type: 'default' } })
    expect(last.entities?.some((e: any) => e.type === 'streaminfo' && e.streamType === 'final')).toBe(true)
  })

  test('403 plain resend keeps the channelData', async () => {
    const s = setup({ feedback: true, streams: true, failFinal: { statusCode: 403, code: 'ContentStreamNotAllowed' } })
    await s.sendStatus('dm', 'working')
    await s.sendReply('dm', 'done')
    const last = messages(s).at(-1)
    expect(last.entities?.some((e: any) => e.type === 'streaminfo') ?? false).toBe(false)
    expect(last.channelData).toEqual({ feedbackLoop: { type: 'default' } })
  })
})
