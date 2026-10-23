import { describe, expect, test, beforeEach, afterEach } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join, basename, dirname } from 'path'
import type { Attachment } from 'botframework-schema'

import { resolveTimeZone } from '../src/config'
import { localDate, localTime, processAttachments } from '../src/teams/attachments'

const NOW = new Date('2026-10-22T22:30:00Z')

describe('localDate / localTime with explicit time zone', () => {
  test('Europe/London', () => {
    expect(localDate(NOW, 'Europe/London')).toBe('2026-10-22')
    expect(localTime(NOW, 'Europe/London')).toBe('2330')
  })
  test('Europe/Budapest', () => {
    expect(localDate(NOW, 'Europe/Budapest')).toBe('2026-10-23')
    expect(localTime(NOW, 'Europe/Budapest')).toBe('0030')
  })
  test('America/New_York', () => {
    expect(localDate(NOW, 'America/New_York')).toBe('2026-10-22')
    expect(localTime(NOW, 'America/New_York')).toBe('1830')
  })
})

describe('localDate / localTime without a time zone', () => {
  test('uses the host default', () => {
    expect(localDate(NOW)).toBe(
      new Intl.DateTimeFormat('en-CA', { year: 'numeric', month: '2-digit', day: '2-digit' }).format(NOW),
    )
    const hhmm = new Intl.DateTimeFormat('en-GB', { hour: '2-digit', minute: '2-digit', hour12: false })
      .format(NOW)
      .replace(':', '')
    expect(localTime(NOW)).toBe(hhmm)
  })
})

describe('resolveTimeZone', () => {
  function capture<T>(fn: () => T): { result: T; lines: string[] } {
    const lines: string[] = []
    const orig = process.stderr.write
    process.stderr.write = ((chunk: unknown) => {
      lines.push(String(chunk))
      return true
    }) as typeof process.stderr.write
    try {
      return { result: fn(), lines }
    } finally {
      process.stderr.write = orig
    }
  }

  test('unset / blank -> undefined', () => {
    expect(resolveTimeZone({})).toBeUndefined()
    expect(resolveTimeZone({ TEAMS_PLUGIN_TZ: '   ' })).toBeUndefined()
  })
  test('trims a valid zone', () => {
    expect(resolveTimeZone({ TEAMS_PLUGIN_TZ: '  Europe/Budapest ' })).toBe('Europe/Budapest')
  })
  test('unknown zone -> undefined with one stderr line', () => {
    const { result, lines } = capture(() => resolveTimeZone({ TEAMS_PLUGIN_TZ: 'Mars/Olympus' }))
    expect(result).toBeUndefined()
    expect(lines).toEqual([
      'teams channel: TEAMS_PLUGIN_TZ "Mars/Olympus" is not a known time zone; using the host time zone\n',
    ])
  })
  test('newline in value is neutralised', () => {
    const { result, lines } = capture(() => resolveTimeZone({ TEAMS_PLUGIN_TZ: 'Europe/Budapest\nX' }))
    expect(result).toBeUndefined()
    expect(lines.length).toBe(1)
    const line = lines[0]!
    expect(line.indexOf('\n')).toBe(line.length - 1)
    expect(line).toContain('"Europe/Budapest?X"')
  })
})

describe('processAttachments honours deps.timeZone', () => {
  let baseDir: string
  beforeEach(() => { baseDir = mkdtempSync(join(tmpdir(), 'cct-tz-')) })
  afterEach(() => { rmSync(baseDir, { recursive: true, force: true }) })

  // A zone far from any plausible host zone, so a missing timeZone pass-through cannot hide behind the host default.
  test('Pacific/Kiritimati -> next-day folder and 1230- prefix', async () => {
    const imageUrl = 'https://smba.trafficmanager.net/uk/v3/attachments/abc/views/original'
    const fetchImpl = (async () =>
      new Response(new Uint8Array([0x89, 0x50, 0x4e, 0x47]), { status: 200 })) as unknown as typeof fetch
    const att = { contentType: 'image/png', contentUrl: imageUrl } as Attachment
    const { results } = await processAttachments([att], 'Adnan Khan', {
      baseDir,
      fetchImpl,
      getBearerToken: async () => 'T',
      now: () => NOW,
      timeZone: 'Pacific/Kiritimati',
    })
    expect(results[0]!.kind).toBe('saved')
    const path = (results[0] as { path: string }).path
    expect(dirname(path)).toBe(join(baseDir, '2026-10-23'))
    expect(basename(path).startsWith('1230-')).toBe(true)
  })
})
