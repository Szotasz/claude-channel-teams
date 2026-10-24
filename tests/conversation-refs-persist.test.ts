import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { chmodSync, mkdirSync, mkdtempSync, readdirSync, rmSync, statSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { createConversationRefStore } from '../src/teams/conversationRefs.js'

const mkRef = (id = 'conv-1', serviceUrl = 'https://smba.trafficmanager.net/emea/') =>
  ({ conversation: { id }, serviceUrl, user: { id: 'u' }, bot: { id: 'b' } }) as any

let dir: string
let file: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'refs-'))
  file = join(dir, 'conversation-refs.json')
})
afterEach(() => {
  try { chmodSync(dir, 0o700) } catch {}
  rmSync(dir, { recursive: true, force: true })
})

function writeRefs(entries: unknown[]) {
  writeFileSync(file, JSON.stringify({ version: 1, refs: entries }))
}
const entry = (o: Record<string, unknown> = {}) => ({
  conversationId: 'conv-1', ref: mkRef(), aadObjectId: 'aad', lastSeen: 1, ...o,
})

describe('conversation refs persistence', () => {
  test('P1 round trip', () => {
    createConversationRefStore({ file }).put('conv-1', mkRef(), 'AAD-X')
    const got = createConversationRefStore({ file }).get('conv-1')
    expect(got?.ref).toEqual(mkRef())
    expect(got?.aadObjectId).toBe('aad-x')
  })

  test('P2 mode 0600, no tmp left', () => {
    createConversationRefStore({ file }).put('conv-1', mkRef(), 'a')
    expect(statSync(file).mode & 0o777).toBe(0o600)
    expect(readdirSync(dir).filter((f) => f.includes('.tmp-'))).toEqual([])
  })

  test('P3 allowlist re-checked', () => {
    createConversationRefStore({ file }).put('conv-1', mkRef(), 'a')
    const s = createConversationRefStore({ file, isAllowed: () => false })
    expect(s.get('conv-1')).toBeUndefined()
    expect(s.size()).toBe(0)
  })

  test('P4 foreign service hosts dropped', () => {
    writeRefs([
      entry({ conversationId: 'c1', ref: mkRef('c1', 'https://evil.test/') }),
      entry({ conversationId: 'c2', ref: mkRef('c2', 'http://smba.trafficmanager.net/') }),
      entry({ conversationId: 'c3', ref: mkRef('c3') }),
    ])
    const s = createConversationRefStore({ file })
    expect(s.get('c1')).toBeUndefined()
    expect(s.get('c2')).toBeUndefined()
    expect(s.get('c3')).toBeDefined()
  })

  test('P5 bad files give empty store', () => {
    writeFileSync(file, 'not json')
    expect(createConversationRefStore({ file }).size()).toBe(0)
    writeFileSync(file, JSON.stringify({ version: 2, refs: [entry()] }))
    expect(createConversationRefStore({ file }).size()).toBe(0)
  })

  test('P6 maxEntries keeps most recent', () => {
    let t = 0
    const s = createConversationRefStore({ file, maxEntries: 3, now: () => ++t })
    for (let i = 1; i <= 5; i++) s.put(`c${i}`, mkRef(`c${i}`), 'a')
    expect(s.size()).toBe(3)
    const r = createConversationRefStore({ file, maxEntries: 3 })
    expect(r.size()).toBe(3)
    expect(r.get('c1')).toBeUndefined()
    expect(r.get('c2')).toBeUndefined()
    for (const id of ['c3', 'c4', 'c5']) expect(r.get(id)).toBeDefined()
  })

  test('P7 no file, nothing written', () => {
    createConversationRefStore().put('conv-1', mkRef(), 'a')
    expect(readdirSync(dir)).toEqual([])
  })

  test('P8 mismatched conversation id dropped', () => {
    writeRefs([entry({ conversationId: 'conv-1', ref: mkRef('other') })])
    expect(createConversationRefStore({ file }).size()).toBe(0)
  })

  test.skipIf(process.getuid?.() === 0)('P9 write failure does not throw', () => {
    const ro = join(dir, 'ro')
    mkdirSync(ro)
    chmodSync(ro, 0o500)
    const s = createConversationRefStore({ file: join(ro, 'refs.json') })
    expect(() => s.put('conv-1', mkRef(), 'a')).not.toThrow()
    expect(s.get('conv-1')).toBeDefined()
  })

  test('P10 mostRecent after reload', () => {
    let t = 10
    const s = createConversationRefStore({ file, now: () => t })
    s.put('a', mkRef('a'), 'x')
    t = 30
    s.put('b', mkRef('b'), 'x')
    t = 20
    s.put('c', mkRef('c'), 'x')
    expect(createConversationRefStore({ file }).mostRecentConversationId()).toBe('b')
  })

  test('P11 a conversation listed twice keeps its newest entry', () => {
    writeRefs([entry({ aadObjectId: 'old', lastSeen: 5 }), entry({ aadObjectId: 'new', lastSeen: 9 })])
    const r = createConversationRefStore({ file })
    expect(r.size()).toBe(1)
    expect(r.get('conv-1')?.aadObjectId).toBe('new')
    expect(r.get('conv-1')?.lastSeen).toBe(9)
  })

  test('P12 a duplicate does not take a maxEntries slot', () => {
    writeRefs([
      entry({ lastSeen: 9 }),
      entry({ lastSeen: 8 }),
      entry({ conversationId: 'conv-2', ref: mkRef('conv-2'), lastSeen: 7 }),
    ])
    const r = createConversationRefStore({ file, maxEntries: 2 })
    expect(r.size()).toBe(2)
    expect(r.get('conv-2')).toBeDefined()
  })

  test('P13 a file over maxEntries loads only the most recent', () => {
    writeRefs([1, 4, 2, 3].map(n => entry({ conversationId: `c${n}`, ref: mkRef(`c${n}`), lastSeen: n })))
    const r = createConversationRefStore({ file, maxEntries: 2 })
    expect(r.size()).toBe(2)
    expect(r.get('c4')).toBeDefined()
    expect(r.get('c3')).toBeDefined()
  })

  test('P14 entry checks on load', () => {
    const id512 = 'x'.repeat(512)
    const id513 = 'y'.repeat(513)
    writeRefs([
      entry({ conversationId: id512, ref: mkRef(id512) }),
      entry({ conversationId: id513, ref: mkRef(id513) }),
      entry({ conversationId: 'c-empty-aad', ref: mkRef('c-empty-aad'), aadObjectId: '' }),
      entry({ conversationId: 'c-bad-seen', ref: mkRef('c-bad-seen'), lastSeen: 'x' }),
      entry({ conversationId: 'c-upper', ref: mkRef('c-upper'), aadObjectId: 'AAD-UPPER' }),
    ])
    const s = createConversationRefStore({ file })
    expect(s.get(id512)).toBeDefined()
    expect(s.get(id513)).toBeUndefined()
    expect(s.get('c-empty-aad')).toBeUndefined()
    expect(s.get('c-bad-seen')).toBeUndefined()
    expect(s.get('c-upper')?.aadObjectId).toBe('aad-upper')
  })

  test('P15 an allowlist check that throws drops the entry', () => {
    writeRefs([entry()])
    const s = createConversationRefStore({ file, isAllowed: () => { throw new Error('boom') } })
    expect(s.size()).toBe(0)
  })

  test('P16 a missing parent directory is created 0700', () => {
    const nested = join(dir, 'state', 'conversation-refs.json')
    createConversationRefStore({ file: nested }).put('conv-1', mkRef(), 'a')
    expect(statSync(join(dir, 'state')).mode & 0o777).toBe(0o700)
  })
})
