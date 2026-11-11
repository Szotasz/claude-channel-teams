import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { chmodSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

const SCRIPT = resolve(import.meta.dir, '..', 'scripts', 'make-teams-manifest.sh')
const APP_ID = '11111111-2222-4333-8444-555555555555'

let tmp: string

beforeEach(() => {
  tmp = mkdtempSync(join(tmpdir(), 'manifest-activity-'))
  const shim = join(tmp, 'zip')
  writeFileSync(shim, '#!/usr/bin/env bash\ncp manifest.json "$3"\n')
  chmodSync(shim, 0o755)
})

afterEach(() => {
  rmSync(tmp, { recursive: true, force: true })
})

function run(...extra: string[]) {
  const out = join(tmp, 'out.zip')
  const r = Bun.spawnSync(
    ['bash', SCRIPT, '--app-id', APP_ID, '--name', 'Test', '--env-file', join(tmp, 'none.env'), '--out', out, ...extra],
    { env: { ...process.env, PATH: `${tmp}:${process.env.PATH}`, HOME: tmp } },
  )
  const raw = r.exitCode === 0 ? readFileSync(out, 'utf8') : ''
  return { code: r.exitCode, stdout: r.stdout.toString(), raw, manifest: raw ? JSON.parse(raw) : undefined }
}

const rscLines = (stdout: string) => stdout.split('\n').filter(l => l.startsWith('RSC:'))

describe('make-teams-manifest.sh --activity-feed', () => {
  test('without the flag there is no webApplicationInfo or authorization', () => {
    const r = run()
    expect(r.code).toBe(0)
    expect('webApplicationInfo' in r.manifest).toBe(false)
    expect('authorization' in r.manifest).toBe(false)
    expect(rscLines(r.stdout)).toEqual([])
  })

  test('--activity-feed alone', () => {
    const r = run('--activity-feed')
    expect(r.code).toBe(0)
    expect(r.manifest.webApplicationInfo).toEqual({ id: APP_ID, resource: 'https://RscBasedStoreApp' })
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: 'TeamsActivity.Send.Group', type: 'Application' },
    ])
    expect(r.manifest.activities).toBeUndefined()
    const lines = rscLines(r.stdout)
    expect(lines.length).toBe(1)
    expect(lines[0]).toContain('TeamsActivity.Send.Group')
  })

  test('both flags give one list with three permissions in order', () => {
    const r = run('--rsc-all-messages', '--activity-feed')
    expect(r.code).toBe(0)
    expect(r.raw.split('webApplicationInfo').length - 1).toBe(1)
    expect(r.raw.split('resourceSpecific').length - 1).toBe(1)
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: 'ChannelMessage.Read.Group', type: 'Application' },
      { name: 'ChatMessage.Read.Chat', type: 'Application' },
      { name: 'TeamsActivity.Send.Group', type: 'Application' },
    ])
    const lines = rscLines(r.stdout)
    expect(lines.length).toBe(2)
    expect(lines.some(l => l.includes('TeamsActivity.Send.Group'))).toBe(true)
    expect(lines.some(l => l.includes('ChannelMessage.Read.Group'))).toBe(true)
  })

  test('--activity-feed adds exactly webApplicationInfo and authorization', () => {
    const plain = run().manifest
    const m = run('--activity-feed').manifest
    delete m.webApplicationInfo
    delete m.authorization
    expect(m).toEqual(plain)
  })

  test('--rsc-all-messages output is unchanged by the new code', () => {
    const r = run('--rsc-all-messages')
    expect(r.manifest.webApplicationInfo).toEqual({ id: APP_ID, resource: 'https://RscBasedStoreApp' })
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: 'ChannelMessage.Read.Group', type: 'Application' },
      { name: 'ChatMessage.Read.Chat', type: 'Application' },
    ])
    expect(rscLines(r.stdout).length).toBe(1)
  })
})
