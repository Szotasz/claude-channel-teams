import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

const server = readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8')

function block(src: string, anchor: string): string {
  const start = src.indexOf(anchor)
  expect(start).toBeGreaterThan(-1)
  expect(src.indexOf(anchor, start + 1)).toBe(-1)
  let depth = 0
  for (let i = src.indexOf('{', start); i >= 0 && i < src.length; i++) {
    if (src[i] === '{') depth++
    else if (src[i] === '}' && --depth === 0) return src.slice(start, i + 1)
  }
  throw new Error(`unbalanced block: ${anchor}`)
}

describe('server.ts mentions wiring', () => {
  const reply = block(server, "case 'reply': {")
  test('non-array check', () => {
    expect(reply).toContain('const rawMentions = args.mentions ?? undefined')
    expect(reply).toContain(
      "if (rawMentions !== undefined && !Array.isArray(rawMentions)) {\n          throw new Error('mentions must be an array')",
    )
  })
  test('branch condition and options spread', () => {
    expect(reply).toContain('aiGenerated || rawCitations !== undefined || rawMentions !== undefined')
    const first = block(reply, 'await replier.sendReply(conversationId, text, card, {')
    expect(first).toContain('...(rawMentions !== undefined ? { mentions: rawMentions } : {}),')
  })
  test('tags passed to the sender', () => {
    expect(server).toContain('mentionTags: config.mentionTags')
    expect(block(server, 'createReplySender({')).toContain('mentionTags: config.mentionTags')
  })
  test('schema maxItems', () => {
    expect(block(server, 'mentions: {')).toContain('maxItems: 10')
    expect(block(server, 'mentions: {')).toContain('minItems: 1,')
  })
})
