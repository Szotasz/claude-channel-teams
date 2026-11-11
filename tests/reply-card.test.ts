import { describe, expect, test, beforeEach, afterEach } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createReplySender } from '../src/teams/reply'
import { validateAdaptiveCard, MAX_CARD_CHARS } from '../src/teams/cards'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const CT = 'application/vnd.microsoft.card.adaptive'

const validCard = () => ({
  type: 'AdaptiveCard',
  version: '1.5',
  body: [{ type: 'TextBlock', text: 'Go?' }],
  actions: [{ type: 'Action.OpenUrl', title: 'Open', url: 'https://example.com/x' }],
})

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

describe('reply with card / message id', () => {
  let dir: string
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cct-replycard-'))
  })
  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  function setup(sendResult: unknown, allowed = true) {
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    if (allowed) allowlist.addEntry(FIXTURE_ID)
    const refs = createConversationRefStore()
    refs.put('conv-abc', { conversation: { id: 'conv-abc' } } as any, FIXTURE_ID)
    const state = { calls: 0, sent: undefined as any, stops: 0 }
    const adapter = {
      async continueConversationAsync(_a: string, _r: any, logic: any) {
        state.calls++
        await logic({
          async sendActivity(act: any) {
            state.sent = act
            return sendResult
          },
        })
      },
    } as any
    const typingPump = { stop: () => void state.stops++ } as any
    const sender = createReplySender({ config: makeConfig(), adapter, allowlist, refs, typingPump })
    return { ...sender, state }
  }

  test('R1 returns the activity id when present', async () => {
    expect(await setup({ id: 'act-1' }).sendReply('conv-abc', 'hi')).toBe('act-1')
    expect(await setup(undefined).sendReply('conv-abc', 'hi')).toBeUndefined()
    expect(await setup({ id: '' }).sendReply('conv-abc', 'hi')).toBeUndefined()
  })

  test('R2 card is attached', async () => {
    const s = setup({ id: 'a' })
    const card = validCard()
    await s.sendReply('conv-abc', 'hi', card)
    expect(s.state.sent).toEqual({
      type: 'message',
      text: 'hi',
      textFormat: 'markdown',
      attachments: [{ contentType: CT, content: card }],
    })
  })

  test('R3 no card means no attachments', async () => {
    const s = setup({ id: 'a' })
    await s.sendReply('conv-abc', 'hi')
    expect('attachments' in s.state.sent).toBe(false)
  })

  test('R4 invalid card sends nothing and keeps the typing pump', async () => {
    const s = setup({ id: 'a' })
    await expect(s.sendReply('conv-abc', 'hi', { ...validCard(), version: '1.6' })).rejects.toThrow(
      'invalid card',
    )
    expect(s.state.calls).toBe(0)
    expect(s.state.stops).toBe(0)
  })

  test('R5 card check comes before the gate', async () => {
    const s = setup({ id: 'a' }, false)
    await expect(s.sendReply('conv-abc', 'hi', { type: 'x' })).rejects.toThrow('invalid card')
  })
})

describe('validateAdaptiveCard', () => {
  const err = (c: unknown) => {
    const r = validateAdaptiveCard(c)
    return r.ok ? '' : r.error
  }

  test('valid card', () => {
    expect(validateAdaptiveCard(validCard())).toEqual({ ok: true })
  })

  test('shape errors', () => {
    for (const c of [null, [], 'x']) expect(err(c)).toBe('card must be a JSON object')
    expect(err({ type: 'Hero', version: '1.5' })).toBe('card.type must be "AdaptiveCard"')
    expect(err({ type: 'AdaptiveCard' })).toBe('card.version must be 1.0 to 1.5')
    expect(err({ type: 'AdaptiveCard', version: '1.6' })).toBe('card.version must be 1.0 to 1.5')
  })

  test('too large', () => {
    const c = { type: 'AdaptiveCard', version: '1.5', body: [{ type: 'TextBlock', text: 'x'.repeat(28001) }] }
    expect(err(c)).toContain('too large')
    expect(MAX_CARD_CHARS).toBe(28000)
  })

  test('https-only urls', () => {
    const base = { type: 'AdaptiveCard', version: '1.5' }
    expect(err({ ...base, actions: [{ url: 'http://a.b' }] })).toBe('card.actions[0].url must be an https:// URL')
    expect(err({ ...base, body: [{ url: 'javascript:alert(1)' }] })).toContain('card.body[0].url')
    expect(
      err({ ...base, body: [{ columns: [{}, { selectAction: { url: 'data:x' } }] }] }),
    ).toContain('card.body[0].columns[1].selectAction.url')
    expect(err({ ...base, body: [{}, { backgroundImage: 'http://x' }] })).toContain('card.body[1].backgroundImage')
    expect(validateAdaptiveCard({ ...base, body: [{ iconUrl: 'https://ok' }] })).toEqual({ ok: true })
    expect(err({ ...base, body: [{ iconUrl: 'http://x' }] })).toContain('card.body[0].iconUrl')
  })

  test('msteams post-back actions are refused, Submit and Execute pass', () => {
    const base = { type: 'AdaptiveCard', version: '1.5' }
    expect(validateAdaptiveCard({ ...base, actions: [{ type: 'Action.Submit', title: 'Go', data: { go: true } }] })).toEqual({ ok: true })
    expect(validateAdaptiveCard({ ...base, actions: [{ type: 'Action.Execute', verb: 'go' }] })).toEqual({ ok: true })
    expect(
      validateAdaptiveCard({ ...base, body: [{ type: 'Container', items: [], selectAction: { type: 'Action.Submit', data: {} } }] }),
    ).toEqual({ ok: true })
    expect(
      validateAdaptiveCard({ ...base, body: [{ type: 'ColumnSet', columns: [{ type: 'Column', selectAction: { type: 'Action.Execute' } }] }] }),
    ).toEqual({ ok: true })
    for (const t of ['imBack', 'messageBack', 'signin']) {
      expect(
        err({ ...base, actions: [{ type: 'Action.ShowCard', card: { type: 'AdaptiveCard', data: { msteams: { type: t, text: 'y abcde' } } } }] }),
      ).toBe(`card.actions[0].card.data.msteams is an msteams ${t} action, which would post text as the user`)
    }
    expect(
      validateAdaptiveCard({
        ...base,
        msteams: { width: 'Full' },
        actions: [{ type: 'Action.OpenUrl', url: 'https://x.example' }, { type: 'Action.ShowCard', card: { type: 'AdaptiveCard', body: [] } }],
      }),
    ).toEqual({ ok: true })
  })

  test('a refresh is refused', () => {
    expect(err({ type: 'AdaptiveCard', version: '1.5', refresh: { action: { type: 'Action.Execute', verb: 'x' } }, body: [] })).toBe(
      'card.refresh is a refresh, which fires without a click',
    )
  })

  test('nesting', () => {
    let node: any = {}
    for (let i = 0; i < 40; i++) node = { items: [node] }
    expect(err({ type: 'AdaptiveCard', version: '1.5', body: [node] })).toContain('nested too deeply')
    const deep = (levels: number) => {
      let v: any = { type: 'TextBlock' }
      for (let i = 0; i < levels; i++) v = { items: v }
      return v
    }
    expect(validateAdaptiveCard({ type: 'AdaptiveCard', version: '1.5', body: deep(31) })).toEqual({ ok: true })
    const e = err({ type: 'AdaptiveCard', version: '1.5', body: deep(32) })
    expect(e).toContain('nested too deeply')
    expect(e).toStartWith('card.body' + '.items'.repeat(32) + ' ')
  })
})
