import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'node:fs'
import { readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { loadConfig, parseFlag } from '../src/config'

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

describe('server.ts status wiring', () => {
  test('tools list offers status_update through statusTools', () => {
    expect(server).toContain('...statusTools(config.statusUpdates),')
  })
  test('statusStreams only when enabled and passed to the reply sender', () => {
    expect(server).toContain('const statusStreams = config.statusUpdates ? createStatusStreams() : undefined')
    expect(server).toContain('createReplySender({ config, adapter, allowlist, refs, typingPump, statusStreams })')
    expect(server.indexOf('const statusStreams =')).toBeLessThan(server.indexOf('const replier = createReplySender'))
  })
  test("the 'status_update' case calls replier.sendStatus", () => {
    const body = block(server, "case 'status_update': {")
    expect(body).toContain('await replier.sendStatus(conversationId, text)')
    expect(body).toContain("return { content: [{ type: 'text', text: 'status shown' }] }")
  })
})

describe('TEAMS_STATUS_UPDATES', () => {
  test('parseFlag', () => {
    for (const v of ['1', 'true', ' TRUE ', 'True']) expect(parseFlag(v)).toBe(true)
    for (const v of ['yes', '', '0', 'false', undefined]) expect(parseFlag(v)).toBe(false)
  })

  test('loadConfig', () => {
    const keys = ['TEAMS_STATUS_UPDATES', 'TEAMS_PLUGIN_STATE_DIR', 'TEAMS_BOT_APP_ID', 'TEAMS_BOT_APP_PASSWORD', 'TEAMS_BOT_TENANT_ID']
    const saved = Object.fromEntries(keys.map(k => [k, process.env[k]]))
    try {
      process.env.TEAMS_PLUGIN_STATE_DIR = mkdtempSync(join(tmpdir(), 'cct-cfg-'))
      process.env.TEAMS_BOT_APP_ID = 'id'
      process.env.TEAMS_BOT_APP_PASSWORD = 'pw'
      process.env.TEAMS_BOT_TENANT_ID = 'tenant'
      for (const [v, want] of [['1', true], ['true', true], [' TRUE ', true], ['yes', false], ['', false], [undefined, false]] as const) {
        if (v === undefined) delete process.env.TEAMS_STATUS_UPDATES
        else process.env.TEAMS_STATUS_UPDATES = v
        expect(loadConfig().statusUpdates).toBe(want)
      }
    } finally {
      for (const k of keys) {
        if (saved[k] === undefined) delete process.env[k]
        else process.env[k] = saved[k]
      }
    }
  })
})
