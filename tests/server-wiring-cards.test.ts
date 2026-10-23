/**
 * Wiring checks for the card meta and its instruction sentences in src/server.ts (the module starts the listener and
 * the MCP transport on import, so it is read, not imported). Each check is scoped to the block that does the work, so a
 * stray copy elsewhere in the file does not satisfy it. Source reading cannot prove behaviour; each check pins the
 * statement a one-line mutation would change. Follow-up to the #27 review.
 */
import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

const server = readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8')

/** The block that starts at `anchor`, from its first '{' or '[' up to the matching closing bracket. */
function block(src: string, anchor: string): string {
  const start = src.indexOf(anchor)
  expect(start).toBeGreaterThan(-1)
  expect(src.indexOf(anchor, start + 1)).toBe(-1) // exactly one such block
  const openAt = src.slice(start).search(/[{[]/)
  expect(openAt).toBeGreaterThan(-1)
  let depth = 0
  for (let i = start + openAt; i < src.length; i++) {
    const ch = src[i]
    if (ch === '{' || ch === '[') depth++
    else if ((ch === '}' || ch === ']') && --depth === 0) return src.slice(start, i + 1)
  }
  throw new Error(`unbalanced block: ${anchor}`)
}

describe('server.ts card wiring', () => {
  test('the channel tag carries card_action, and action_verb only when there is a verb', () => {
    const notify = block(server, 'const onEvent: ChannelEventSink = event => {')
    const meta = block(notify, 'meta: {')
    expect(meta).toContain(
      '...(event.cardAction ? { card_action: event.cardAction.kind, ...(event.cardAction.verb ? { action_verb: event.cardAction.verb } : {}) } : {}),',
    )
  })

  test('the instructions say what card_action="execute"/"submit" and card_action="dialog" mean', () => {
    const instructions = block(server, 'instructions: [')
    const card = 'A message whose tag has card_action="execute" or card_action="submit" is a click on a card this bot sent'
    const dialog = 'A message with card_action="dialog" is a dialog the user submitted'
    expect(instructions.split(card).length).toBe(2)
    expect(instructions.split(dialog).length).toBe(2)
  })
})
