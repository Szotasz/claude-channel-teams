/** Wiring checks for TEAMS_FEEDBACK in src/server.ts, read as source (see tests/server-wiring.test.ts). */
import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { loadConfig } from '../src/config'
import { channelEventMeta } from '../src/teams/adapter'

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

describe('feedback wiring', () => {
  test('meta carries the rating: onEvent builds it with channelEventMeta', () => {
    expect(block(server, 'const onEvent: ChannelEventSink = event => {')).toContain('meta: channelEventMeta(event),')
    const event = { text: 'x', conversationId: 'c', aadObjectId: 'a', tenantId: 't', fromName: 'n', messageId: 'm1' }
    const meta = channelEventMeta({ ...event, feedback: 'dislike' })
    expect(meta.feedback).toBe('dislike')
    expect(meta.message_id).toBe('m1')
    expect('feedback' in channelEventMeta(event)).toBe(false)
  })

  test('the instruction sentence is present once, only when config.feedback is true', () => {
    expect(server.split('A tag with feedback=').length - 1).toBe(1)
    expect(server).toMatch(/\.\.\.\(config\.feedback\n\s*\? \[\n\s*'A tag with feedback=/)
  })

  test('feedback is passed to the reply sender and the listener', () => {
    expect(server).toMatch(/createReplySender\(\{[^}]*feedback: config\.feedback[^}]*\}\)/)
    expect(block(server, 'const listener = startListener({')).toContain('feedback: config.feedback,')
  })

  test('loadConfig', () => {
    const keys = ['TEAMS_FEEDBACK', 'TEAMS_PLUGIN_STATE_DIR', 'TEAMS_BOT_APP_ID', 'TEAMS_BOT_APP_PASSWORD', 'TEAMS_BOT_TENANT_ID']
    const saved = Object.fromEntries(keys.map(k => [k, process.env[k]]))
    try {
      process.env.TEAMS_PLUGIN_STATE_DIR = mkdtempSync(join(tmpdir(), 'cct-cfg-'))
      process.env.TEAMS_BOT_APP_ID = 'id'
      process.env.TEAMS_BOT_APP_PASSWORD = 'pw'
      process.env.TEAMS_BOT_TENANT_ID = 'tenant'
      process.env.TEAMS_FEEDBACK = '1'
      expect(loadConfig().feedback).toBe(true)
      process.env.TEAMS_FEEDBACK = 'true'
      expect(loadConfig().feedback).toBe(true)
      delete process.env.TEAMS_FEEDBACK
      expect(loadConfig().feedback).toBe(false)
    } finally {
      for (const k of keys) {
        if (saved[k] === undefined) delete process.env[k]
        else process.env[k] = saved[k]
      }
    }
  })
})
