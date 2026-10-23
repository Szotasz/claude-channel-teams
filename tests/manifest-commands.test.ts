import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { chmodSync, existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

const SCRIPT = resolve(import.meta.dir, '..', 'scripts', 'make-teams-manifest.sh')
const APP_ID = '11111111-2222-4333-8444-555555555555'
const SCHEMA_129 = 'https://developer.microsoft.com/en-us/json-schemas/teams/v1.29/MicrosoftTeams.schema.json'

let tmp: string

beforeEach(() => {
  tmp = mkdtempSync(join(tmpdir(), 'manifest-cmd-'))
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
  const raw = r.exitCode === 0 ? readFileSync(out, 'utf8') : undefined
  return {
    code: r.exitCode,
    stdout: r.stdout.toString(),
    stderr: r.stderr.toString(),
    raw,
    manifest: raw === undefined ? undefined : JSON.parse(raw),
    outExists: existsSync(out),
  }
}

const cmd = (v: string) => ['--command', v]
const hasCommandsLine = (s: string) => s.split('\n').some(l => l.startsWith('COMMANDS:'))

describe('make-teams-manifest.sh --command', () => {
  test('C1 no --command leaves the manifest untouched', () => {
    const a = run()
    const b = run()
    expect(a.code).toBe(0)
    expect(a.manifest.bots[0].supportsTargetedMessages).toBeUndefined()
    expect(a.manifest.bots[0].commandLists).toBeUndefined()
    expect(a.manifest.manifestVersion).toBe('1.16')
    expect(a.raw).toBe(b.raw as string)
    expect(hasCommandsLine(a.stdout)).toBe(false)
  })

  test('C2 two commands', () => {
    const r = run(...cmd('/deadlines: Upcoming tender deadlines'), ...cmd('filter: Filter tenders by keyword'))
    expect(r.code).toBe(0)
    const bot = r.manifest.bots[0]
    expect(bot.supportsTargetedMessages).toBe(true)
    expect(bot.commandLists).toEqual([
      {
        scopes: ['personal'],
        triggers: ['mention', 'slash'],
        commands: [
          { title: 'deadlines', description: 'Upcoming tender deadlines' },
          { title: 'filter', description: 'Filter tenders by keyword' },
        ],
      },
    ])
    expect(r.manifest.manifestVersion).toBe('1.29')
    expect(r.manifest['$schema']).toBe(SCHEMA_129)
    expect(hasCommandsLine(r.stdout)).toBe(true)
  })

  test('C3 quotes, backslash and Hungarian letters round-trip', () => {
    const desc = 'Szűrés "kulcsszó" szerint \\ ok'
    const r = run(...cmd(`f: ${desc}`))
    expect(r.code).toBe(0)
    expect(r.manifest.bots[0].commandLists[0].commands[0].description).toBe(desc)
  })

  test('C12 the value is split at the first ": ", later ones stay in the description', () => {
    const r = run(...cmd('find: Search: by keyword: or id'))
    expect(r.code).toBe(0)
    expect(r.manifest.bots[0].commandLists[0].commands[0]).toEqual({ title: 'find', description: 'Search: by keyword: or id' })
  })

  test('C13 shell syntax in a description stays literal text', () => {
    const desc = 'run $(echo pwned) and `echo pwned` and ${HOME}'
    const r = run(...cmd(`x: ${desc}`))
    expect(r.code).toBe(0)
    expect(r.manifest.bots[0].commandLists[0].commands[0].description).toBe(desc)
  })

  test('C4 control characters are rejected', () => {
    for (const d of ['a\tb', 'a\nb']) {
      rmSync(join(tmp, 'out.zip'), { force: true })
      const r = run(...cmd(`x: ${d}`))
      expect(r.code).toBe(2)
      expect(r.outExists).toBe(false)
    }
  })

  test('C5 bad titles', () => {
    for (const v of ['has space: x', ': x', '/: x', 'ékezet: x', `${'a'.repeat(129)}: x`]) {
      const r = run(...cmd(v))
      expect(r.code).toBe(2)
      expect(r.outExists).toBe(false)
    }
    const ok = run(...cmd(`${'a'.repeat(128)}: x`))
    expect(ok.code).toBe(0)
  })

  test('C6 12 commands accepted, 13 rejected', () => {
    const mk = (n: number) => Array.from({ length: n }, (_, i) => cmd(`c${i}: d`)).flat()
    const ok = run(...mk(12))
    expect(ok.code).toBe(0)
    expect(ok.manifest.bots[0].commandLists[0].commands.length).toBe(12)
    rmSync(join(tmp, 'out.zip'), { force: true })
    const bad = run(...mk(13))
    expect(bad.code).toBe(2)
    expect(bad.outExists).toBe(false)
  })

  test('C7 duplicate title', () => {
    const r = run(...cmd('a: one'), ...cmd('/A: two'))
    expect(r.code).toBe(2)
    expect(r.outExists).toBe(false)
  })

  test('C8 missing ": "', () => {
    expect(run(...cmd('nocolon')).code).toBe(2)
    expect(run(...cmd('a:b')).code).toBe(2)
  })

  test('C9 together with --rsc-all-messages', () => {
    const r = run('--rsc-all-messages', ...cmd('a: one'))
    expect(r.code).toBe(0)
    expect(r.manifest.webApplicationInfo).toEqual({ id: APP_ID, resource: 'https://RscBasedStoreApp' })
    expect(r.manifest.authorization.permissions.resourceSpecific.length).toBe(2)
    expect(r.manifest.bots[0].commandLists[0].commands).toEqual([{ title: 'a', description: 'one' }])
    expect(r.manifest.manifestVersion).toBe('1.29')
  })

  test('C10 the option adds nothing else', () => {
    const plain = run().manifest
    const m = run(...cmd('/deadlines: Upcoming'), ...cmd('filter: Filter')).manifest
    delete m.bots[0].supportsTargetedMessages
    delete m.bots[0].commandLists
    m.manifestVersion = plain.manifestVersion
    m['$schema'] = plain['$schema']
    expect(m).toEqual(plain)
  })

  test('C11 description byte limits', () => {
    expect(run(...cmd(`a: ${'x'.repeat(4000)}`)).code).toBe(0)
    expect(run(...cmd(`a: ${'x'.repeat(4001)}`)).code).toBe(2)
    expect(run(...cmd(`a: ${'é'.repeat(2000)}`)).code).toBe(0)
    expect(run(...cmd(`a: ${'é'.repeat(2001)}`)).code).toBe(2)
  })

  test('C14 a description that is not valid UTF-8 exits 2 and writes nothing', () => {
    // argv cannot carry a lone byte from JS, so printf builds it inside bash.
    for (const bytes of ['caf\\377', 'caf\\303', '\\355\\240\\200']) {
      const out = join(tmp, 'bad.zip')
      const r = Bun.spawnSync(
        ['bash', '-c', `exec bash "$1" --app-id "$2" --env-file "$3" --out "$4" --command "$(printf 'x: ${bytes}')"`,
          '_', SCRIPT, APP_ID, join(tmp, 'none.env'), out],
        { env: { ...process.env, PATH: `${tmp}:${process.env.PATH}`, HOME: tmp } },
      )
      expect(r.exitCode).toBe(2)
      expect(r.stderr.toString()).toContain('not valid UTF-8')
      expect(existsSync(out)).toBe(false)
    }
  })

  test('C15 --command without a value exits 2', () => {
    const r = Bun.spawnSync(['bash', SCRIPT, '--app-id', APP_ID, '--out', join(tmp, 'x.zip'), '--command'], {
      env: { ...process.env, HOME: tmp },
      cwd: tmp,
    })
    expect(r.exitCode).toBe(2)
    expect(r.stderr.toString()).toContain('--command needs a value')
  })
})
