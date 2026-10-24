import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { ConversationNotAllowedError, createReplySender } from '../src/teams/reply'
import { createStatusStreams } from '../src/teams/statusStream'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'

function makeConfig() {
  return {
    stateDir: '/tmp/unused',
    allowlistFile: '/tmp/unused',
    appId: 'app-id-from-config',
    appPassword: 'app-pw',
    appType: 'SingleTenant' as const,
    tenantId: TENANT_ID,
    bindHost: '127.0.0.1',
    port: 0,
  }
}

describe('status updates in replies', () => {
  let dir: string
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cct-replystatus-'))
  })

  function setup(opts: { streams?: boolean; failFinal?: unknown; onSleep?: () => void } = {}) {
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(FIXTURE_ID)
    const refs = createConversationRefStore()
    refs.put('dm', { conversation: { id: 'a:1', conversationType: 'personal' } } as any, FIXTURE_ID)
    refs.put('chan', { conversation: { id: '19:x@thread.tacv2', conversationType: 'channel' } } as any, FIXTURE_ID)
    const state = { sent: [] as any[], stops: [] as string[], failFinal: opts.failFinal }
    const adapter = {
      async continueConversationAsync(_a: string, _r: any, logic: any) {
        await logic({
          async sendActivity(act: any) {
            state.sent.push(act)
            if (act.type === 'message' && state.failFinal && act.entities?.some((e: any) => e.type === 'streaminfo')) {
              const err = state.failFinal
              throw err
            }
            return { id: 's-1' }
          },
        })
      },
    } as any
    const typingPump = { stop: (id: string) => void state.stops.push(id) } as any
    const statusStreams = opts.streams === false ? undefined : createStatusStreams({ sleep: async () => opts.onSleep?.() })
    const sender = createReplySender({ config: makeConfig(), adapter, allowlist, refs, typingPump, statusStreams })
    return { ...sender, state, allowlist, refs, statusStreams }
  }

  test('a person removed from the allowlist while the update waits gets nothing', async () => {
    let s: ReturnType<typeof setup>
    s = setup({ onSleep: () => void s.allowlist.removeEntry(FIXTURE_ID) })
    await s.sendStatus('dm', 'one')
    await expect(s.sendStatus('dm', 'two')).rejects.toBeInstanceOf(ConversationNotAllowedError)
    expect(s.state.sent.length).toBe(1)
  })

  test('a reply that waited in the status queue re-checks the allowlist', async () => {
    let s: ReturnType<typeof setup>
    let revoke = false
    s = setup({ onSleep: () => void (revoke && s.allowlist.removeEntry(FIXTURE_ID)) })
    await s.sendStatus('dm', 'one')
    revoke = true
    await expect(s.sendReply('dm', 'Done')).rejects.toBeInstanceOf(ConversationNotAllowedError)
    expect(s.state.sent.length).toBe(1)
  })

  test('a conversation that stops being a 1:1 chat while the update waits gets nothing', async () => {
    let s: ReturnType<typeof setup>
    s = setup({
      onSleep: () =>
        s.refs.put('dm', { conversation: { id: '19:y@thread.tacv2', conversationType: 'channel' } } as any, FIXTURE_ID),
    })
    await s.sendStatus('dm', 'one')
    await expect(s.sendStatus('dm', 'two')).rejects.toThrow('status updates work only in a 1:1 chat')
    expect(s.state.sent.length).toBe(1)
  })

  test('a reply whose conversation stopped being a 1:1 chat while it waited does not close the stream', async () => {
    let s: ReturnType<typeof setup>
    let flip = false
    s = setup({
      onSleep: () =>
        void (flip &&
          s.refs.put('dm', { conversation: { id: '19:y@thread.tacv2', conversationType: 'channel' } } as any, FIXTURE_ID)),
    })
    await s.sendStatus('dm', 'one')
    flip = true
    await s.sendReply('dm', 'Done')
    const last = s.state.sent.at(-1)
    expect(last.type).toBe('message')
    expect((last.entities ?? []).some((e: any) => e.type === 'streaminfo')).toBe(false)
  })

  test('only 1:1 replies go through the status queue', async () => {
    const s = setup()
    let queued = 0
    const original = s.statusStreams!.withStream
    s.statusStreams!.withStream = ((id: string, fn: any) => (queued++, original(id, fn))) as any
    await s.sendReply('chan', 'In the channel')
    expect(queued).toBe(0)
    await s.sendReply('dm', 'In the chat')
    expect(queued).toBe(1)
  })

  test('outside a 1:1 chat the call is refused before anything else', async () => {
    const s = setup()
    await expect(s.sendStatus('chan', '')).rejects.toThrow('status updates work only in a 1:1 chat')
  })

  test('an invalid text does not stop the typing indicator', async () => {
    const s = setup()
    await expect(s.sendStatus('dm', '   ')).rejects.toThrow('status text must be 1-1000 characters')
    expect(s.state.stops).toEqual([])
    expect(s.state.sent).toEqual([])
  })

  test('not enabled', async () => {
    const s = setup({ streams: false })
    await expect(s.sendStatus('dm', 'hi')).rejects.toThrow('status updates are not enabled (set TEAMS_STATUS_UPDATES)')
    expect(s.state.sent).toEqual([])
  })

  test('channel refused', async () => {
    const s = setup()
    await expect(s.sendStatus('chan', 'hi')).rejects.toThrow('status updates work only in a 1:1 chat')
    expect(s.state.sent).toEqual([])
  })

  test('status then replies', async () => {
    const s = setup()
    await s.sendStatus('dm', 'Working')
    expect(s.state.sent).toHaveLength(1)
    expect(s.state.sent[0].type).toBe('typing')
    expect(s.state.stops).toEqual(['dm'])
    await s.sendReply('dm', 'Done')
    expect(s.state.sent[1].entities).toEqual([{ type: 'streaminfo', streamId: 's-1', streamType: 'final' }])
    await s.sendReply('dm', 'Again')
    expect('entities' in s.state.sent[2]).toBe(false)
  })

  test('ai label comes before the final entity', async () => {
    const s = setup()
    await s.sendStatus('dm', 'Working')
    await s.sendReply('dm', 'Done', undefined, { aiGenerated: true })
    const ents = s.state.sent[1].entities
    expect(ents).toHaveLength(2)
    expect(ents[0].additionalType).toEqual(['AIGeneratedContent'])
    expect(ents[1]).toEqual({ type: 'streaminfo', streamId: 's-1', streamType: 'final' })
  })

  test('closed stream: reply resent as a new message', async () => {
    const s = setup({ failFinal: { statusCode: 403, code: 'ContentStreamNotAllowed' } })
    await s.sendStatus('dm', 'Working')
    await s.sendReply('dm', 'Done')
    expect(s.state.sent).toHaveLength(3)
    expect(s.state.sent[2].entities?.some((e: any) => e.type === 'streaminfo') ?? false).toBe(false)
  })

  test('other errors reject after one send', async () => {
    const s = setup({ failFinal: { statusCode: 500 } })
    await s.sendStatus('dm', 'Working')
    await expect(s.sendReply('dm', 'Done')).rejects.toEqual({ statusCode: 500 })
    expect(s.state.sent).toHaveLength(2)
  })

  test('no status update: plain activity', async () => {
    const s = setup()
    await s.sendReply('dm', 'hi')
    expect(s.state.sent).toEqual([{ type: 'message', text: 'hi', textFormat: 'markdown' }])
  })
})
