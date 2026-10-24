/**
 * Wiring checks for code that lives inside src/server.ts closures (the module starts the listener and the MCP
 * transport on import, so it is not imported here). Like tests/files-base-url.test.ts, these read the source, scoped to
 * the block that does the work, so a stray or dead copy elsewhere in the file does not satisfy them. Source reading
 * cannot prove behaviour; each check pins the exact statements a one-line mutation would change.
 * Follow-ups to #4 (mentioned in the channel tag), #9 (a null card is no card) and #12 (the permission prompt only
 * goes to a 1:1 chat).
 */
import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

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

/** The source with leading whitespace and comments removed, so `startsWith` sees the first statement. */
function firstStatement(src: string): string {
  return src.replace(/^(\s|\/\/[^\n]*|\/\*[\s\S]*?\*\/)*/, '')
}

describe('server.ts wiring', () => {
  test('#12 the permission prompt target is the most recent 1:1 chat, re-checked against the allowlist', () => {
    const body = block(server, 'resolveTargetConversation: () => {')
    expect(body).toContain('const id = refs.mostRecentPersonalConversationId()')
    expect(body).not.toContain('mostRecentConversationId(')
    expect(body).toContain('const stored = refs.get(id)')
    expect(body).toContain('if (!allowlist.isAllowed(stored.aadObjectId)) return undefined')
    // In this order: the 1:1 lookup, the stored reference, the allowlist guard, and only then `return id`.
    const at = (s: string) => body.indexOf(s)
    expect(at('const id = refs.mostRecentPersonalConversationId()')).toBeLessThan(at('const stored = refs.get(id)'))
    expect(at('const stored = refs.get(id)')).toBeLessThan(at('if (!allowlist.isAllowed(stored.aadObjectId))'))
    expect(at('if (!allowlist.isAllowed(stored.aadObjectId))')).toBeLessThan(at('return id'))
    // Every return is a whole line: `return undefined`, or `return id` as the last statement. No line continues an
    // expression from the line before, so a return cannot carry on onto the next line.
    expect((body.match(/\breturn\b[^\n]*/g) ?? []).filter(r => r !== 'return undefined')).toEqual(['return id'])
    expect(body).toMatch(/\n\s*return id\n\s*\}$/)
    expect(body).not.toMatch(/\n\s*(\?|:|\|\||&&|\.)/)
  })

  test('#9 the reply tool treats a null card as no card, and sends that normalized card', () => {
    const reply = block(server, "case 'reply': {")
    const normalize = 'const card = args.card ?? undefined'
    expect(reply).toContain(normalize)
    // After the normalization nothing reads args.card again, and every send passes `card`, except the one that runs
    // only when there is no card.
    expect(reply.slice(reply.indexOf(normalize) + normalize.length)).not.toContain('args.card')
    const sends = reply.match(/replier\.sendReply\([^\n]*/g) ?? []
    expect(sends.length).toBe((reply.match(/\bsendReply\b/g) ?? []).length)
    expect(sends.map(c => c.replace(/^replier\.sendReply\(conversationId, text, card(, \{|\))$/, 'card')))
      .toEqual(['card', 'card', 'replier.sendReply(conversationId, text)'])
    expect(reply).toMatch(
      /: card !== undefined\s*\? await replier\.sendReply\(conversationId, text, card\)\s*: await replier\.sendReply\(conversationId, text\)\n/,
    )
  })

  test('#4 the channel tag carries mentioned as "true" / "false", only when it is known', () => {
    const spread = "...(event.mentioned !== undefined ? { mentioned: event.mentioned ? 'true' : 'false' } : {})"
    const notify = block(server, 'const onEvent: ChannelEventSink = event => {')
    if (notify.includes('meta: channelEventMeta(event)')) {
      // The meta is built by the adapter helper: check the helper that the notification's meta actually comes from.
      expect(block(read('teams', 'adapter.ts'), 'export function channelEventMeta(')).toContain(spread)
    } else {
      expect(block(notify, 'meta: {')).toContain(spread)
    }
  })

  test('invisible characters are stripped from the text before it reaches the model', () => {
    const notify = block(server, 'const onEvent: ChannelEventSink = event => {')
    const strip = 'event = withInvisibleStripped(event)'
    expect(notify).toContain(strip)
    // The strip happens before the notification is built and before the text is read.
    const at = (s: string) => notify.indexOf(s)
    expect(at(strip)).toBeLessThan(at('mcp'))
    expect(at(strip)).toBeLessThan(at('content: event.text'))
    expect(server).toContain("import { withInvisibleStripped } from './teams/invisibleChars.js'")
  })

  test('inbound debouncing: the listener pushes into the debouncer, which delivers to onEvent', () => {
    expect(server).toContain("import { createInboundDebouncer } from './teams/inboundDebounce.js'")
    // Created after the sink, delivering into it — the strip inside onEvent still runs last.
    expect(server).toContain(
      'const inboundDebouncer = createInboundDebouncer({ windowMs: config.inboundDebounceMs, deliver: onEvent })',
    )
    const listener = block(server, 'const listener = startListener({')
    expect(listener).toContain('onEvent: inboundDebouncer.push')
    // Not the bare sink shorthand — that would bypass the debouncer entirely.
    expect(listener).not.toMatch(/^\s*onEvent,\s*$/m)
  })

  test('shutdown flushes whatever the debouncer is holding before anything else is torn down', () => {
    const shutdown = block(server, 'async function shutdown(): Promise<void> {')
    const tryBody = shutdown.slice(shutdown.indexOf('try {') + 'try {'.length)
    expect(firstStatement(tryBody).startsWith('inboundDebouncer.flushAll()')).toBe(true)
    expect(shutdown.indexOf('inboundDebouncer.flushAll()')).toBeLessThan(shutdown.indexOf('permission.clear()'))
  })

  test('the onEvent block still begins with the invisible-character strip line', () => {
    const notify = block(server, 'const onEvent: ChannelEventSink = event => {')
    const body = notify.slice(notify.indexOf('{') + 1)
    expect(firstStatement(body).startsWith('event = withInvisibleStripped(event)')).toBe(true)
  })
})
