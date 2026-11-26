/**
 * Decision log tests: the line shape (and only the allowed keys), the promise that no message text or display
 * name ever reaches the file, rotation by size, the 0600 mode, and that a broken path never throws.
 */

import { afterEach, describe, expect, test } from 'bun:test'
import { chmodSync, existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { createDecisionLog } from '../src/teams/decisionLog'
import { makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createInboundIds } from '../src/teams/inboundIds'

let dir: string
let file: string
function fresh(): void {
  dir = mkdtempSync(join(tmpdir(), 'cct-decisions-'))
  file = join(dir, 'decisions.log')
}
afterEach(() => {
  try { chmodSync(dir, 0o700) } catch {}
  rmSync(dir, { recursive: true, force: true })
})

const KEYS = ['ts', 'decision', 'reason', 'source', 'message_id', 'channel']
const read = () => readFileSync(file, 'utf8').split('\n').filter(l => l.length > 0).map(l => JSON.parse(l) as Record<string, unknown>)
const log = (maxBytes = 1 << 20, now: () => number = () => Date.parse('2026-06-01T12:00:00Z')) =>
  createDecisionLog({ file, maxBytes, now })

describe('decision log', () => {
  test('one JSON object per line, only the allowed keys', () => {
    fresh()
    log().record({ decision: 'deliver', reason: 'ok', source: 'live', messageId: 'm1', conversationId: '19:abc@thread.tacv2;messageid=1' })
    log().record({ decision: 'backfill_error', reason: 'throttled', source: 'backfill' })
    const lines = read()
    expect(lines).toHaveLength(2)
    for (const line of lines) {
      expect(Object.keys(line).every(k => KEYS.includes(k))).toBe(true)
      expect(typeof line.ts).toBe('string')
      expect(line.ts).toBe('2026-06-01T12:00:00.000Z')
    }
    expect(lines[0]).toEqual({
      ts: '2026-06-01T12:00:00.000Z',
      decision: 'deliver',
      reason: 'ok',
      source: 'live',
      message_id: 'm1',
      channel: '19:abc@thread.tacv2',
    })
    expect(lines[1]).toEqual({ ts: '2026-06-01T12:00:00.000Z', decision: 'backfill_error', reason: 'throttled', source: 'backfill' })
  })

  test('a reason that is not a short fixed code becomes "other"; ids are bounded', () => {
    fresh()
    log().record({ decision: 'drop', reason: 'not a code, and definitely too long anyway', source: 'live' })
    log().record({ decision: 'drop', reason: 'ok', source: 'live', messageId: 'x'.repeat(80) })
    const lines = read()
    expect(lines[0]!.reason).toBe('other')
    expect((lines[1]!.message_id as string).length).toBe(64)
  })

  test('never message text, display names or aad object ids', async () => {
    fresh()
    const SECRET_TEXT = 'SECRET-MESSAGE-TEXT-4f2a'
    const SECRET_NAME = 'Secret-Sender-Name-91'
    const AAD = '00000000-0000-4000-8000-000000000001'
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(AAD)
    const events: any[] = []
    const handler = makeTurnHandler({
      config: { stateDir: dir, allowlistFile: join(dir, 'a.json'), appId: 'app-1', appPassword: 'p', appType: 'SingleTenant', tenantId: 't', bindHost: '127.0.0.1', port: 0 } as any,
      adapter: {} as any,
      allowlist,
      refs: createConversationRefStore(),
      onEvent: e => events.push(e),
      inboundIds: createInboundIds(),
      decisions: log(),
    } as any)
    await handler({
      activity: {
        type: 'message',
        id: 'm1',
        text: `<at>Bot</at> ${SECRET_TEXT}`,
        from: { id: 'u', name: SECRET_NAME, aadObjectId: AAD },
        conversation: { id: '19:abc@thread.tacv2;messageid=1', tenantId: 't', conversationType: 'channel', isGroup: true },
        recipient: { id: 'bot', name: 'Bot' },
        channelId: 'msteams',
        serviceUrl: 'https://smba.trafficmanager.net/emea/',
        entities: [{ type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot', name: 'Bot' } }],
      } as any,
    })
    expect(events).toHaveLength(1)
    const raw = readFileSync(file, 'utf8')
    expect(raw).not.toContain(SECRET_TEXT)
    expect(raw).not.toContain(SECRET_NAME)
    expect(raw).not.toContain(AAD)
    expect(read().map(l => `${l.decision}/${l.reason}`)).toEqual(['deliver/ok'])
  })

  test('rotation moves the old file aside and keeps the current one within the cap', () => {
    fresh()
    const d = log(250)
    for (let i = 0; i < 5; i++) d.record({ decision: 'drop', reason: 'duplicate', source: 'live', messageId: `m${i}` })
    expect(existsSync(`${file}.1`)).toBe(true)
    expect(statSync(file).size).toBeLessThanOrEqual(250)
    expect(read()).toHaveLength(1)
    expect(readFileSync(`${file}.1`, 'utf8').split('\n').filter(l => l.length > 0).length).toBeGreaterThanOrEqual(1)
  })

  test('the file is created mode 0600', () => {
    fresh()
    log().record({ decision: 'deliver', reason: 'ok', source: 'live' })
    expect(statSync(file).mode & 0o777).toBe(0o600)
  })

  test.skipIf(process.getuid?.() === 0)('an unwritable path warns once and never throws', () => {
    fresh()
    const ro = mkdtempSync(join(dir, 'ro-'))
    chmodSync(ro, 0o500)
    const d = createDecisionLog({ file: join(ro, 'decisions.log'), maxBytes: 1 << 20 })
    const lines: string[] = []
    const write = process.stderr.write
    process.stderr.write = ((chunk: unknown) => { lines.push(String(chunk)); return true }) as typeof process.stderr.write
    try {
      expect(() => {
        d.record({ decision: 'deliver', reason: 'ok', source: 'live' })
        d.record({ decision: 'drop', reason: 'ok', source: 'live' })
      }).not.toThrow()
    } finally {
      process.stderr.write = write
    }
    expect(lines.length).toBe(1)
    expect(lines[0]).toContain('decision log write failed')
    expect(readdirSync(ro)).toEqual([])
  })
})
