/** Wiring checks for TEAMS_CHANNEL_ATTACHMENTS in src/server.ts, read as source (see tests/server-wiring.test.ts). */
import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { loadConfig } from '../src/config'
import { stripComments } from './strip-comments'

const server = stripComments(readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8'))
const config = stripComments(readFileSync(join(import.meta.dir, '..', 'src', 'config.ts'), 'utf8'))

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

describe('channel attachments wiring', () => {
  test('sharedFiles is bound once, only when config.channelAttachments is true', () => {
    expect(server.split('sharedFiles: async').length - 1).toBe(1)
    expect(block(server, 'const listener = startListener({')).toMatch(/\.\.\.\(config\.channelAttachments\s*\?\s*\{\s*sharedFiles: async/)
  })

  test('channelAttachments is parsed from TEAMS_CHANNEL_ATTACHMENTS once', () => {
    expect(config.split('channelAttachments: parseFlag(env.TEAMS_CHANNEL_ATTACHMENTS),').length - 1).toBe(1)
  })

  test('loadConfig', () => {
    const keys = ['TEAMS_CHANNEL_ATTACHMENTS', 'TEAMS_PLUGIN_STATE_DIR', 'TEAMS_BOT_APP_ID', 'TEAMS_BOT_APP_PASSWORD', 'TEAMS_BOT_TENANT_ID']
    const saved = Object.fromEntries(keys.map(k => [k, process.env[k]]))
    try {
      process.env.TEAMS_PLUGIN_STATE_DIR = mkdtempSync(join(tmpdir(), 'cct-cfg-'))
      process.env.TEAMS_BOT_APP_ID = 'id'
      process.env.TEAMS_BOT_APP_PASSWORD = 'pw'
      process.env.TEAMS_BOT_TENANT_ID = 'tenant'
      delete process.env.TEAMS_CHANNEL_ATTACHMENTS
      expect(loadConfig().channelAttachments).toBe(false)
      process.env.TEAMS_CHANNEL_ATTACHMENTS = '1'
      expect(loadConfig().channelAttachments).toBe(true)
      process.env.TEAMS_CHANNEL_ATTACHMENTS = 'true'
      expect(loadConfig().channelAttachments).toBe(true)
      process.env.TEAMS_CHANNEL_ATTACHMENTS = 'nope'
      expect(loadConfig().channelAttachments).toBe(false)
    } finally {
      for (const k of keys) {
        if (saved[k] === undefined) delete process.env[k]
        else process.env[k] = saved[k]
      }
    }
  })
})
