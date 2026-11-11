import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { chmodSync, existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

const SCRIPT = resolve(import.meta.dir, '..', 'scripts', 'make-teams-manifest.sh')
const APP_ID = '11111111-2222-4333-8444-555555555555'

let tmp: string

beforeEach(() => {
  tmp = mkdtempSync(join(tmpdir(), 'manifest-ma-'))
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
    stderr: r.stderr.toString(),
    raw,
    manifest: raw === undefined ? undefined : JSON.parse(raw),
    outExists: existsSync(out),
  }
}

const ma = (v: string) => ['--message-action', v]
const command = (id: string, title: string, description: string) => ({
  id,
  type: 'action',
  context: ['message'],
  title,
  description,
  fetchTask: false,
  parameters: [{ name: 'note', title: 'Note', description: 'Optional note for the assistant', inputType: 'textarea' }],
})

function expectRejected(...args: string[]) {
  rmSync(join(tmp, 'out.zip'), { force: true })
  const r = run(...args)
  expect(r.code).toBe(2)
  expect(r.outExists).toBe(false)
  expect(r.stderr).toContain('--message-action')
}

describe('make-teams-manifest.sh --message-action', () => {
  test('M1 no option: no composeExtensions; the option adds exactly that member', () => {
    const plain = run()
    expect(plain.manifest.composeExtensions).toBeUndefined()
    const m = run(...ma('ask: Ask the owner: Send this to the assistant')).manifest
    expect(m.manifestVersion).toBe('1.16')
    delete m.composeExtensions
    expect(m).toEqual(plain.manifest)
  })

  test('M2 one and two options give the exact JSON, order kept', () => {
    const one = run(...ma('ask: Ask: Send to the assistant'))
    expect(one.manifest.composeExtensions).toEqual([
      { botId: APP_ID, commands: [command('ask', 'Ask', 'Send to the assistant')] },
    ])
    const two = run(...ma('b-2: Second: Two'), ...ma('a1: First: One: more'))
    expect(two.manifest.composeExtensions).toEqual([
      { botId: APP_ID, commands: [command('b-2', 'Second', 'Two'), command('a1', 'First', 'One: more')] },
    ])
  })

  test('M3 together with --command and --rsc-all-messages', () => {
    const r = run('--command', 'a: one', '--rsc-all-messages', ...ma('ask: Ask: Send'))
    expect(r.code).toBe(0)
    expect(r.manifest.manifestVersion).toBe('1.29')
    expect(r.manifest.bots[0].commandLists.length).toBe(1)
    expect(r.manifest.webApplicationInfo.id).toBe(APP_ID)
    expect(r.manifest.composeExtensions[0].commands.length).toBe(1)
    expect(Object.keys(r.manifest).slice(-3)).toEqual(['webApplicationInfo', 'authorization', 'composeExtensions'])
  })

  test('M4 quotes and backslashes are escaped', () => {
    const title = 'Say "hi" \\ now'
    const desc = 'A "quoted" back\\slash'
    const r = run(...ma(`q: ${title}: ${desc}`))
    expect(r.code).toBe(0)
    const c = r.manifest.composeExtensions[0].commands[0]
    expect(c.title).toBe(title)
    expect(c.description).toBe(desc)
  })

  test('M5 Hungarian title', () => {
    const r = run(...ma('kerdes: Kérdés a kiírónak: Kérdés küldése'))
    expect(r.code).toBe(0)
    expect(r.manifest.composeExtensions[0].commands[0].title).toBe('Kérdés a kiírónak')
  })

  test('M6 limits: 10 accepted, 11 rejected', () => {
    const mk = (n: number) => Array.from({ length: n }, (_, i) => ma(`a${i}: T: D`)).flat()
    const ok = run(...mk(10))
    expect(ok.code).toBe(0)
    expect(ok.manifest.composeExtensions[0].commands.length).toBe(10)
    expectRejected(...mk(11))
  })

  test('M7 bad ids', () => {
    for (const id of ['Bad', '-x', 'a'.repeat(65), 'a_b', '']) expectRejected(...ma(`${id}: T: D`))
    expect(run(...ma(`${'a'.repeat(64)}: T: D`)).code).toBe(0)
  })

  test('M8 duplicate id', () => {
    expectRejected(...ma('a: T: D'), ...ma('a: U: E'))
  })

  test('M9 title and description byte lengths', () => {
    expectRejected(...ma('a: : D'))
    expectRejected(...ma(`a: ${'x'.repeat(33)}: D`))
    expect(run(...ma(`a: ${'x'.repeat(32)}: D`)).code).toBe(0)
    expect(run(...ma(`a: ${'é'.repeat(16)}: D`)).code).toBe(0)
    expectRejected(...ma(`a: ${'é'.repeat(17)}: D`))
    expectRejected(...ma('a: T: '))
    expectRejected(...ma(`a: T: ${'x'.repeat(129)}`))
    expect(run(...ma(`a: T: ${'x'.repeat(128)}`)).code).toBe(0)
  })

  test('M10 control characters', () => {
    expectRejected(...ma('a: T\tx: D'))
    expectRejected(...ma('a: T: D\nx'))
    // C1 controls (U+0080..U+009F) are controls too; U+00A0 (no-break space) is not.
    expectRejected(...ma('a: A\u0085B: D'))
    expectRejected(...ma('a: T: D\u009fx'))
    expect(run(...ma('a: A\u00a0B: D')).code).toBe(0)
  })

  test('M11 invalid UTF-8', () => {
    for (const bytes of ['caf\\377', 'caf\\303']) {
      for (const where of ['title', 'desc']) {
        const v = where === 'title' ? `a: ${bytes}: D` : `a: T: ${bytes}`
        const out = join(tmp, 'bad.zip')
        const r = Bun.spawnSync(
          ['bash', '-c', `exec bash "$1" --app-id "$2" --env-file "$3" --out "$4" --message-action "$(printf '${v}')"`,
            '_', SCRIPT, APP_ID, join(tmp, 'none.env'), out],
          { env: { ...process.env, PATH: `${tmp}:${process.env.PATH}`, HOME: tmp } },
        )
        expect(r.exitCode).toBe(2)
        expect(r.stderr.toString()).toContain('not valid UTF-8')
        expect(existsSync(out)).toBe(false)
      }
    }
  })

  test('M12 malformed values', () => {
    expectRejected(...ma('nocolon'))
    expectRejected(...ma('a:b: c'))
    expectRejected(...ma('a: only title'))
  })

  test('M13 missing value exits 2', () => {
    const r = Bun.spawnSync(['bash', SCRIPT, '--app-id', APP_ID, '--out', join(tmp, 'x.zip'), '--message-action'], {
      env: { ...process.env, HOME: tmp },
      cwd: tmp,
    })
    expect(r.exitCode).toBe(2)
    expect(r.stderr.toString()).toContain('--message-action needs a value')
  })
})
