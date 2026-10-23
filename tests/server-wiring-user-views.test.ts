/**
 * Wiring checks for user_cards in src/server.ts (the module starts the listener and the MCP transport on import, so it
 * is read, not imported), the same source-level way as the #29 tests. Each check is scoped to the block that does the
 * work, so a stray copy elsewhere in the file does not satisfy it. Source reading cannot prove behaviour; each check
 * pins the statement a one-line mutation would change.
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

describe('server.ts user_cards wiring', () => {
  test('the reply tool refuses a non-array user_cards and passes it on in the options', () => {
    const reply = block(server, "case 'reply': {")
    expect(reply).toMatch(
      /const rawUserCards = args\.user_cards \?\? undefined\n\s*if \(rawUserCards !== undefined && !Array\.isArray\(rawUserCards\)\) \{\n\s*throw new Error\('user_cards must be an array'\)/,
    )
    // user_cards alone also takes the branch with options.
    expect(reply).toMatch(
      /const id = aiGenerated \|\| rawCitations !== undefined \|\| rawUserCards !== undefined \|\| targetedMessageId !== undefined \|\| isPublic\n\s*\? await replier\.sendReply\(conversationId, text, card, \{/,
    )
    expect(block(reply, 'await replier.sendReply(conversationId, text, card, {')).toContain(
      '...(rawUserCards !== undefined ? { userCards: rawUserCards } : {}),',
    )
  })

  test('the reply sender and the listener share the same stores', () => {
    expect(server.split('const knownUsers = createKnownUsers()').length).toBe(2)
    expect(server.split('const userViews = createUserViewStore()').length).toBe(2)
    const sender = block(server, 'const replier = createReplySender(')
    expect(sender).toMatch(/\bknownUsers\b/)
    expect(sender).toMatch(/\buserViews\b/)
    const listener = block(server, 'const listener = startListener(')
    expect(listener).toMatch(/\n\s*knownUsers,\n/)
    expect(listener).toMatch(/\n\s*userViews,\n/)
  })

  test('the reply tool schema lists user_cards with the 60-person cap', () => {
    const schema = block(server, 'user_cards: {')
    expect(schema).toContain('maxItems: 60,')
    expect(schema).toContain("required: ['aad_object_id', 'card'],")
  })
})
