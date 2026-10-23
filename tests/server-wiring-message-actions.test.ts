/**
 * Wiring checks for the message_action meta and its instruction sentence in src/server.ts (the module starts the
 * listener and the MCP transport on import, so it is read, not imported). Each check is scoped to the block that does
 * the work, so a stray copy elsewhere in the file does not satisfy it. Source reading cannot prove behaviour; each
 * check pins the statement a one-line mutation would change. Follow-up to the #28 review.
 */
import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

const server = readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8')

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

describe('server.ts message action wiring', () => {
  test('the channel tag carries message_action when the event has one', () => {
    const notify = block(server, 'const onEvent: ChannelEventSink = event => {')
    expect(block(notify, 'meta: {')).toContain('...(event.messageAction ? { message_action: event.messageAction } : {}),')
  })

  test('the message_action sentence is in the instructions only when TEAMS_MESSAGE_ACTIONS is on', () => {
    const sentence = 'A tag with message_action="<command>" is a message action'
    expect(server.split(sentence).length).toBe(2)
    const cond = server.indexOf('...(config.messageActions')
    expect(cond).toBeGreaterThan(-1)
    expect(server.indexOf('?', cond)).toBeLessThan(server.indexOf(sentence))
    expect(server.indexOf(sentence)).toBeLessThan(server.indexOf(': []', cond))
    expect(server.slice(cond, server.indexOf(sentence))).not.toContain(']')
    // The sentence is the first entry of the branch taken when the switch is on: a negated, compared or widened
    // condition, or swapped branches, do not match.
    expect(server).toMatch(/\.\.\.\(config\.messageActions\n\s*\? \[\n\s*'A tag with message_action="<command>" is a message action/)
  })
})
