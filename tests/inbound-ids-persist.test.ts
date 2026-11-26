/**
 * Persistence of the inbound-id store (TEAMS_BACKFILL): the round trip, the bounds on load, the 0600 mode, the
 * one-warning behaviour for a file that is not usable, and that a failing save never throws into the turn.
 */

import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { chmodSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { createInboundIds, MAX_ID_CONVERSATIONS, MAX_IDS_PER_CONVERSATION } from '../src/teams/inboundIds'

let dir: string
let file: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'cct-inbound-ids-'))
  file = join(dir, 'inbound-ids.json')
})
afterEach(() => {
  try { chmodSync(dir, 0o700) } catch {}
  rmSync(dir, { recursive: true, force: true })
})

/** The loader refuses a file that others can write, so every fixture is written owner-only. */
const writeOwnerOnly = (path: string, body: string) => writeFileSync(path, body, { mode: 0o600 })

/** Captures stderr while `run` executes; a write that is not a string is passed through untouched. */
function capture(run: () => void): string[] {
  const lines: string[] = []
  const write = process.stderr.write
  process.stderr.write = ((chunk: unknown) => {
    lines.push(String(chunk))
    return true
  }) as typeof process.stderr.write
  try {
    run()
  } finally {
    process.stderr.write = write
  }
  return lines
}

describe('inbound ids persistence', () => {
  test('a round trip through the file', () => {
    const a = createInboundIds({ file })
    expect(a.firstSeen('19:abc@thread.tacv2', '1')).toBe(true)
    expect(a.firstSeen('19:abc@thread.tacv2', '2')).toBe(true)
    expect(a.firstSeen('19:other@thread.tacv2', '1')).toBe(true)
    a.forget('19:abc@thread.tacv2', '2')
    const b = createInboundIds({ file })
    expect(b.firstSeen('19:abc@thread.tacv2', '1')).toBe(false)
    expect(b.firstSeen('19:abc@thread.tacv2', '2')).toBe(true)
    expect(b.firstSeen('19:other@thread.tacv2', '1')).toBe(false)
  })

  test('mode 0600, no temp file left behind, nothing written without a file', () => {
    createInboundIds({ file }).firstSeen('c', '1')
    expect(statSync(file).mode & 0o777).toBe(0o600)
    expect(readdirSync(dir).filter(f => f.includes('.tmp'))).toEqual([])
    createInboundIds().firstSeen('c', '1')
    expect(readdirSync(dir)).toEqual(['inbound-ids.json'])
  })

  test('a missing file starts empty with no warning', () => {
    const lines = capture(() => {
      const s = createInboundIds({ file })
      expect(s.firstSeen('c', '1')).toBe(true)
    })
    expect(lines).toEqual([])
  })

  test('the bounds are applied on load, keeping the newest', () => {
    const ids = Array.from({ length: MAX_IDS_PER_CONVERSATION + 5 }, (_, i) => `id${i}`)
    writeOwnerOnly(file, JSON.stringify({ version: 1, conversations: [['k0', ids]] }))
    const s = createInboundIds({ file })
    // Repeats return before anything is dropped, so the oldest end can be probed without disturbing the store.
    expect(s.firstSeen('k0', `id${MAX_IDS_PER_CONVERSATION + 4}`)).toBe(false) // the newest end is kept
    expect(s.firstSeen('k0', 'id5')).toBe(false)
    expect(s.firstSeen('k0', 'id4')).toBe(true) // pushed out on load: only the newest 500 are kept
    expect(s.firstSeen('k0', 'id0')).toBe(true)
    const file2 = join(dir, 'inbound-ids-2.json')
    const conversations = Array.from({ length: MAX_ID_CONVERSATIONS }, (_, i) => [`c${i}`, ['x']] as [string, string[]])
    writeOwnerOnly(file2, JSON.stringify({ version: 1, conversations }))
    const t = createInboundIds({ file: file2 })
    t.firstSeen('one-more', 'y') // over the conversation bound: the oldest goes
    expect(t.firstSeen('c0', 'x')).toBe(true)
    writeOwnerOnly(file2, JSON.stringify({ version: 1, conversations }))
    expect(createInboundIds({ file: file2 }).firstSeen('c1', 'x')).toBe(false)
  })

  test('a corrupt or wrong-shaped file starts empty with exactly one warning', () => {
    for (const body of [
      'not json',
      JSON.stringify({ version: 2, conversations: [] }),
      JSON.stringify({ version: 1, conversations: { k: ['1'] } }),
      JSON.stringify({ version: 1, conversations: [['k']] }),
      JSON.stringify({ version: 1, conversations: [['k', ['1'], 'extra']] }),
      JSON.stringify({ version: 1, conversations: [[1, ['1']]] }),
      JSON.stringify({ version: 1, conversations: [['k', [1]]] }),
      JSON.stringify({ version: 1, conversations: [['k', ['']]] }),
    ]) {
      writeOwnerOnly(file, body)
      const lines = capture(() => {
        const s = createInboundIds({ file })
        expect(s.firstSeen('k', '1')).toBe(true)
      })
      expect(lines.length).toBe(1)
      expect(lines[0]).toContain('inbound-ids file unreadable')
      // Nothing of the file's content leaks into the warning.
      expect(lines[0]).not.toContain('not json')
      expect(lines[0]).not.toContain('conversations')
    }
    // A file that IS the expected shape loads without a word.
    writeOwnerOnly(file, JSON.stringify({ version: 1, conversations: [['SECRET-CONVERSATION-KEY', ['SECRET-MESSAGE-ID']]] }))
    const lines = capture(() => {
      expect(createInboundIds({ file }).firstSeen('SECRET-CONVERSATION-KEY', 'SECRET-MESSAGE-ID')).toBe(false)
    })
    expect(lines).toEqual([])
  })

  test('a file that is not ours or is writable by others is refused', () => {
    writeOwnerOnly(file, JSON.stringify({ version: 1, conversations: [['k', ['1']]] }))
    chmodSync(file, 0o620)
    const group = capture(() => {
      const s = createInboundIds({ file })
      expect(s.firstSeen('k', '1')).toBe(true)
    })
    expect(group.length).toBe(1)
    expect(group[0]).toContain('writable by group or others')
    chmodSync(file, 0o600)
    const foreign = capture(() => {
      const s = createInboundIds({ file, uid: () => (process.getuid?.() ?? 0) + 4242 })
      expect(s.firstSeen('k', '1')).toBe(true)
    })
    expect(foreign.length).toBe(1)
    expect(foreign[0]).toContain('belongs to another user')
  })

  test.skipIf(process.getuid?.() === 0)('a failed save warns once per process and does not throw', () => {
    const ro = mkdtempSync(join(dir, 'ro-'))
    chmodSync(ro, 0o500)
    const target = join(ro, 'inbound-ids.json')
    const lines = capture(() => {
      const s = createInboundIds({ file: target })
      expect(() => {
        s.firstSeen('c', '1')
        s.firstSeen('c', '2')
      }).not.toThrow()
      expect(s.firstSeen('c', '1')).toBe(false)
    })
    expect(lines.length).toBe(1)
    expect(lines[0]).toContain('could not save inbound ids')
  })

  test('the in-memory behaviour is unchanged without a file', () => {
    const s = createInboundIds()
    expect(s.firstSeen('c', '1')).toBe(true)
    expect(s.firstSeen('c', '1')).toBe(false)
    expect(s.firstSeen('d', '1')).toBe(true)
    s.forget('c', '1')
    expect(s.firstSeen('c', '1')).toBe(true)
    expect(() => s.forget('nope', '1')).not.toThrow()
    expect(MAX_IDS_PER_CONVERSATION).toBe(500)
    expect(MAX_ID_CONVERSATIONS).toBe(1000)
  })
})
