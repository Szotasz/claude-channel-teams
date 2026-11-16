import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

const server = readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8')

/** The block that starts at `anchor`, from its first '{' up to the matching closing brace. */
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

describe('server.ts create_channel wiring', () => {
  test('the tools list includes channelTools(config.channelCreateTeams)', () => {
    expect(server).toContain('    ...channelTools(config.channelCreateTeams),\n')
    expect(server.indexOf('...channelTools(config.channelCreateTeams),')).toBeGreaterThan(
      server.indexOf('...plannerTools(config.plannerPlanIds),'),
    )
  })

  test('the create_channel case validates and calls createChannel with the graph token', () => {
    const b = block(server, "case 'create_channel': {")
    expect(b).toContain('validateCreateChannelArgs(args, config.channelCreateTeams)')
    expect(b).toContain('createChannel(input, { getToken: () => graphToken.getToken() })')
    // Every check runs before a token is requested: validation comes first, and createChannel takes its result.
    expect(b.indexOf('validateCreateChannelArgs(')).toBeGreaterThan(-1)
    expect(b.indexOf('validateCreateChannelArgs(')).toBeLessThan(b.indexOf('createChannel(input'))
  })
})
