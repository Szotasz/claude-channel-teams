/**
 * TEAMS_INBOUND_DEBOUNCE_MS — the inbound debounce window. Off (0) unless a
 * whole number of milliseconds is set; anything else turns the feature off with
 * one warning instead of guessing. Follows tests/files-base-url.test.ts: the
 * parser is pure and takes the env value directly, so no process.env fiddling.
 */

import { describe, expect, test } from 'bun:test'

import { parseInboundDebounceMs } from '../src/config'

const VALID = 'teams channel: TEAMS_INBOUND_DEBOUNCE_MS must be a whole number of milliseconds; debouncing is off'

function parse(raw: string | undefined): { value: number; warnings: string[] } {
  const warnings: string[] = []
  const value = parseInboundDebounceMs(raw, m => warnings.push(m))
  return { value, warnings }
}

describe('parseInboundDebounceMs', () => {
  test('unset or empty -> 0 (debouncing off), no warning', () => {
    expect(parse(undefined)).toEqual({ value: 0, warnings: [] })
    expect(parse('')).toEqual({ value: 0, warnings: [] })
    expect(parse('   ')).toEqual({ value: 0, warnings: [] })
  })

  test('a whole number of milliseconds is used as-is', () => {
    expect(parse('1500').value).toBe(1500)
    expect(parse('0').value).toBe(0)
    expect(parse('0').warnings).toEqual([])
    expect(parse('200').value).toBe(200)
    expect(parse('1500').warnings).toEqual([])
  })

  test('anything else -> 0 plus exactly one warning line', () => {
    for (const raw of ['-5', 'abc', '1.5']) {
      expect(parse(raw)).toEqual({ value: 0, warnings: [VALID] })
    }
  })

  test('the value is capped at 10000 without a warning', () => {
    expect(parse('999999')).toEqual({ value: 10000, warnings: [] })
    expect(parse('10000').value).toBe(10000)
    expect(parse('10001').value).toBe(10000)
  })
})
