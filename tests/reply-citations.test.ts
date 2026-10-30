import { describe, expect, test, beforeEach, afterEach } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { buildMessageEntity, createReplySender, UnknownConversationError } from '../src/teams/reply'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'

const cfg = {
  stateDir: '/tmp/unused',
  allowlistFile: '/tmp/unused',
  appId: 'app-id-from-config',
  appPassword: 'app-pw',
  appType: 'SingleTenant' as const,
  tenantId: TENANT_ID,
  bindHost: '127.0.0.1',
  port: 0,
}

const base = { type: 'https://schema.org/Message', '@type': 'Message', '@context': 'https://schema.org', '@id': '' }
const cite = (o: Record<string, unknown> = {}) => ({ name: 'N', abstract: 'A', ...o }) as any

describe('reply citations / AI label', () => {
  let dir: string
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cct-replycite-'))
  })
  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  function setup() {
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(FIXTURE_ID)
    const refs = createConversationRefStore()
    refs.put('conv-abc', { conversation: { id: 'conv-abc' } } as any, FIXTURE_ID)
    const state = { calls: 0, sent: undefined as any }
    const adapter = {
      async continueConversationAsync(_a: string, _r: any, logic: any) {
        state.calls++
        await logic({
          async sendActivity(act: any) {
            state.sent = act
            return { id: 'm1' }
          },
        })
      },
    } as any
    const sender = createReplySender({ config: cfg, adapter, allowlist, refs })
    return { ...sender, state }
  }

  test('E1 no options -> no entities key', async () => {
    const s = setup()
    await s.sendReply('conv-abc', 'hi')
    expect('entities' in s.state.sent).toBe(false)
  })

  test('E2 ai_generated only', async () => {
    const s = setup()
    await s.sendReply('conv-abc', 'hi', undefined, { aiGenerated: true })
    expect(s.state.sent.entities).toEqual([{ ...base, additionalType: ['AIGeneratedContent'] }])
    expect('citation' in s.state.sent.entities[0]).toBe(false)
  })

  test('E3 two citations', async () => {
    const s = setup()
    await s.sendReply('conv-abc', 'hi [1]', undefined, {
      citations: [cite({ url: 'https://example.com/claim-1', keywords: ['a', 'b'] }), cite({ name: 'M', abstract: 'B' })],
    })
    expect(s.state.sent.entities).toEqual([
      {
        ...base,
        citation: [
          { '@type': 'Claim', position: 1, appearance: { '@type': 'DigitalDocument', name: 'N', abstract: 'A', url: 'https://example.com/claim-1', keywords: ['a', 'b'] } },
          { '@type': 'Claim', position: 2, appearance: { '@type': 'DigitalDocument', name: 'M', abstract: 'B' } },
        ],
      },
    ])
    expect('additionalType' in s.state.sent.entities[0]).toBe(false)
  })

  test('E4 both -> one entity', async () => {
    const s = setup()
    await s.sendReply('conv-abc', 'hi', undefined, { aiGenerated: true, citations: [cite()] })
    expect(s.state.sent.entities).toHaveLength(1)
    expect(s.state.sent.entities[0].additionalType).toEqual(['AIGeneratedContent'])
    expect(s.state.sent.entities[0].citation).toHaveLength(1)
    expect(Object.keys(s.state.sent.entities[0])).toEqual(['type', '@type', '@context', '@id', 'additionalType', 'citation'])
  })

  test('E5 explicit positions kept', async () => {
    const s = setup()
    await s.sendReply('conv-abc', 'hi', undefined, { citations: [cite({ position: 3 }), cite({ position: 1 })] })
    expect(s.state.sent.entities[0].citation.map((c: any) => c.position)).toEqual([3, 1])
  })

  const invalid: [string, any[]][] = [
    ['21 citations', Array.from({ length: 21 }, () => cite())],
    ['empty name', [cite({ name: '' })]],
    ['81-char name', [cite({ name: 'x'.repeat(81) })]],
    ['empty abstract', [cite({ abstract: '' })]],
    ['161-char abstract', [cite({ abstract: 'x'.repeat(161) })]],
    ['position 0', [cite({ position: 0 })]],
    ['position 1.5', [cite({ position: 1.5 })]],
    ['duplicate positions', [cite({ position: 2 }), cite({ position: 2 })]],
    ['position equal to default', [cite(), cite({ position: 1 })]],
    ['javascript url', [cite({ url: 'javascript:alert(1)' })]],
    ['not a url', [cite({ url: 'not a url' })]],
    ['4 keywords', [cite({ keywords: ['a', 'b', 'c', 'd'] })]],
    ['empty keyword', [cite({ keywords: [''] })]],
    ['29-char keyword', [cite({ keywords: ['x'.repeat(29)] })]],
    ['unknown key', [cite({ titel: 'x' })]],
  ]
  for (const [label, citations] of invalid) {
    test(`E6 rejects ${label}`, async () => {
      const s = setup()
      await expect(s.sendReply('conv-abc', 'hi', undefined, { citations })).rejects.toThrow(/citation/)
      expect(s.state.calls).toBe(0)
    })
  }

  test('E10 limits inclusive, code points', async () => {
    const s = setup()
    await s.sendReply('conv-abc', 'hi', undefined, {
      citations: Array.from({ length: 20 }, () => cite({ name: 'x'.repeat(80), abstract: 'y'.repeat(160), keywords: ['k'.repeat(28), 'k'.repeat(28), 'k'.repeat(28)] })),
    })
    expect(s.state.sent.entities[0].citation).toHaveLength(20)
    await s.sendReply('conv-abc', 'hi', undefined, { citations: [cite({ name: '😀'.repeat(80) })] })
    expect(s.state.sent.entities[0].citation[0].appearance.name).toBe('😀'.repeat(80))
  })

  test('E7 card plus citations', async () => {
    const s = setup()
    const card = { type: 'AdaptiveCard', version: '1.5', body: [{ type: 'TextBlock', text: 'Go?' }] }
    await s.sendReply('conv-abc', 'hi', card, { citations: [cite()] })
    expect(s.state.sent.attachments).toEqual([{ contentType: 'application/vnd.microsoft.card.adaptive', content: card }])
    expect(s.state.sent.entities).toHaveLength(1)
    expect(s.state.sent.textFormat).toBe('markdown')
  })

  test('E8 validation before allowlist check', async () => {
    const s = setup()
    const p = s.sendReply('unknown-conv', 'hi', undefined, { citations: [cite({ name: '' })] })
    await expect(p).rejects.toThrow(/citation 0/)
    await expect(p).rejects.not.toBeInstanceOf(UnknownConversationError)
  })

  test('E9 buildMessageEntity undefined when nothing requested', () => {
    expect(buildMessageEntity({})).toBeUndefined()
    expect(buildMessageEntity({ aiGenerated: false, citations: [] })).toBeUndefined()
  })

  test('E11 the url that goes out is the parsed one; http is accepted', async () => {
    const s = setup()
    await s.sendReply('conv-abc', 'hi [1] [2]', undefined, {
      citations: [cite({ url: ' https://exa\tmple.com/a ' }), cite({ url: 'http://example.org/b' })],
    })
    const c = s.state.sent.entities[0].citation
    expect(c[0].appearance.url).toBe('https://example.com/a')
    expect(c[1].appearance.url).toBe('http://example.org/b')
  })

  test('E12 null in an optional citation field means absent', async () => {
    const s = setup()
    await s.sendReply('conv-abc', 'hi [1]', undefined, { citations: [cite({ url: null, keywords: null, position: null })] })
    expect(s.state.sent.entities[0].citation).toEqual([
      { '@type': 'Claim', position: 1, appearance: { '@type': 'DigitalDocument', name: 'N', abstract: 'A' } },
    ])
  })

  test('E13 a citation that is not an object, and keywords that are not an array, are refused', () => {
    for (const bad of [[null], ['x'], [[cite()]], [cite({ keywords: 'a' })]]) {
      expect(() => buildMessageEntity({ citations: bad as any })).toThrow()
    }
  })
})
