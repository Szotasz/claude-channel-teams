/**
 * Wiring checks for the meeting-events option. src/server.ts starts the listener and the MCP
 * transport on import, so like tests/server-wiring.test.ts these read the source, scoped to the
 * block that does the work. The config side (parseFlag) is exercised directly: there is no
 * existing config test to follow.
 */
import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

import { parseFlag } from '../src/config'

const read = (...p: string[]) => readFileSync(join(import.meta.dir, '..', 'src', ...p), 'utf8')
const server = read('server.ts')

/** The block that starts at `anchor`, from its first '{' up to the matching closing brace. */
function block(src: string, anchor: string): string {
  const start = src.indexOf(anchor)
  expect(start).toBeGreaterThan(-1)
  expect(src.indexOf(anchor, start + 1)).toBe(-1) // exactly one such block
  let depth = 0
  for (let i = src.indexOf('{', start); i >= 0 && i < src.length; i++) {
    if (src[i] === '{') depth++
    else if (src[i] === '}' && --depth === 0) return src.slice(start, i + 1)
  }
  throw new Error(`unbalanced block: ${anchor}`)
}

describe('server.ts meeting-events wiring', () => {
  test('the listener deps carry meetingEvents from the config', () => {
    expect(block(server, 'const listener = startListener({')).toContain('meetingEvents: config.meetingEvents,')
  })

  test('the meeting_event instruction appears once, inside the config.meetingEvents branch', () => {
    const sentences = server.match(/'A tag with meeting_event=[^\n]*/g) ?? []
    expect(sentences.length).toBe(1)
    expect(server).toMatch(/\.\.\.\(config\.meetingEvents\n\s*\? \[\n\s*'A tag with meeting_event=/)
  })
})

describe('parseFlag', () => {
  test("accepts '1' and 'true', trimmed and lower-cased", () => {
    expect(parseFlag(' TRUE ')).toBe(true)
    expect(parseFlag('1')).toBe(true)
  })

  test('anything else, including nothing at all, is false', () => {
    expect(parseFlag('yes')).toBe(false)
    expect(parseFlag(undefined)).toBe(false)
  })
})
