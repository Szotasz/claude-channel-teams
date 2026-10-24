import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { chmodSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

const SCRIPT = resolve(import.meta.dir, '..', 'scripts', 'make-teams-manifest.sh')
const APP_ID = '11111111-2222-4333-8444-555555555555'
const RSC_LINE = 'RSC: requests Channel.Create.Group for create_channel; the team owner consents at install.'

let tmp: string

beforeEach(() => {
  tmp = mkdtempSync(join(tmpdir(), 'manifest-channel-'))
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
  const manifest = r.exitCode === 0 ? JSON.parse(readFileSync(out, 'utf8')) : undefined
  return { code: r.exitCode, stdout: r.stdout.toString(), manifest }
}

describe('make-teams-manifest.sh --rsc-channel-create', () => {
  test('alone', () => {
    const r = run('--rsc-channel-create')
    expect(r.code).toBe(0)
    expect(r.manifest.webApplicationInfo).toEqual({ id: APP_ID, resource: 'https://RscBasedStoreApp' })
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: 'Channel.Create.Group', type: 'Application' },
    ])
  })

  test('with --rsc-all-messages', () => {
    const r = run('--rsc-all-messages', '--rsc-channel-create')
    expect(r.code).toBe(0)
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: 'ChannelMessage.Read.Group', type: 'Application' },
      { name: 'ChatMessage.Read.Chat', type: 'Application' },
      { name: 'Channel.Create.Group', type: 'Application' },
    ])
  })

  test('--rsc-all-messages alone is unchanged', () => {
    const r = run('--rsc-all-messages')
    expect(r.code).toBe(0)
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: 'ChannelMessage.Read.Group', type: 'Application' },
      { name: 'ChatMessage.Read.Chat', type: 'Application' },
    ])
    expect(r.stdout).not.toContain(RSC_LINE)
  })

  test('summary line only with the option', () => {
    const plain = run()
    expect(plain.code).toBe(0)
    expect(plain.stdout.split('\n')).not.toContain(RSC_LINE)
    expect(plain.manifest.authorization).toBeUndefined()
    const withOption = run('--rsc-channel-create')
    expect(withOption.code).toBe(0)
    expect(withOption.stdout.split('\n')).toContain(RSC_LINE)
  })
})
