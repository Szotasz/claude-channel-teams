/**
 * Download-link base URL: derived from THIS install's own public endpoint,
 * never a fixed host. A fixed host would put every install's download token
 * (and, via Microsoft's link prefetch, the request) on a third party.
 */

import { describe, expect, test } from 'bun:test'
import { readdirSync, readFileSync, statSync } from 'fs'
import { join } from 'path'

import { resolveFilesBaseUrl } from '../src/config'

describe('resolveFilesBaseUrl', () => {
  test('uses the origin of TEAMS_BOT_ENDPOINT_URL + /files', () => {
    expect(
      resolveFilesBaseUrl({ TEAMS_BOT_ENDPOINT_URL: 'https://box.tail1234.ts.net/api/messages' }),
    ).toBe('https://box.tail1234.ts.net/files')
  })

  test('keeps a non-default port (Funnel on 8443)', () => {
    expect(
      resolveFilesBaseUrl({ TEAMS_BOT_ENDPOINT_URL: 'https://box.tail1234.ts.net:8443/api/messages' }),
    ).toBe('https://box.tail1234.ts.net:8443/files')
  })

  test('drops query, fragment and any path of the endpoint', () => {
    expect(
      resolveFilesBaseUrl({ TEAMS_BOT_ENDPOINT_URL: 'https://teams-bot.example.test/bot/api/messages?x=1#y' }),
    ).toBe('https://teams-bot.example.test/files')
  })

  test('TEAMS_FILES_BASE_URL wins, trailing slashes trimmed', () => {
    expect(
      resolveFilesBaseUrl({
        TEAMS_FILES_BASE_URL: 'https://files.example.test/teams/files/',
        TEAMS_BOT_ENDPOINT_URL: 'https://box.tail1234.ts.net/api/messages',
      }),
    ).toBe('https://files.example.test/teams/files')
  })

  test('nothing set -> null (send_file disabled), never a fallback host', () => {
    expect(resolveFilesBaseUrl({})).toBeNull()
  })

  test('http or garbage -> null, and an explicit bad override does not fall through', () => {
    expect(resolveFilesBaseUrl({ TEAMS_BOT_ENDPOINT_URL: 'http://box.tail1234.ts.net/api/messages' })).toBeNull()
    expect(resolveFilesBaseUrl({ TEAMS_BOT_ENDPOINT_URL: 'not a url' })).toBeNull()
    expect(
      resolveFilesBaseUrl({
        TEAMS_FILES_BASE_URL: 'ftp://files.example.test/files',
        TEAMS_BOT_ENDPOINT_URL: 'https://box.tail1234.ts.net/api/messages',
      }),
    ).toBeNull()
  })
})

describe('wiring', () => {
  function tsFiles(dir: string): string[] {
    return readdirSync(dir).flatMap(name => {
      const p = join(dir, name)
      return statSync(p).isDirectory() ? tsFiles(p) : p.endsWith('.ts') ? [p] : []
    })
  }

  test('server.ts passes config.filesBaseUrl to the file sender', () => {
    const src = readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8')
    expect(src).toMatch(/downloadBaseUrl:\s*config\.filesBaseUrl\b/)
  })

  test('no string-literal download base URL anywhere under src/', () => {
    const offenders = tsFiles(join(import.meta.dir, '..', 'src')).filter(f =>
      /downloadBaseUrl:\s*['"`]/.test(readFileSync(f, 'utf8')),
    )
    expect(offenders).toEqual([])
  })
})
