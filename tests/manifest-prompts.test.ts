import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { chmodSync, existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

const SCRIPT = resolve(import.meta.dir, '..', 'scripts', 'make-teams-manifest.sh')
const APP_ID = '11111111-2222-4333-8444-555555555555'

let tmp: string

beforeEach(() => {
  tmp = mkdtempSync(join(tmpdir(), 'manifest-prompt-'))
  const shim = join(tmp, 'zip')
  writeFileSync(shim, '#!/usr/bin/env bash\ncp manifest.json "$3"\n')
  chmodSync(shim, 0o755)
})

afterEach(() => {
  rmSync(tmp, { recursive: true, force: true })
})

function run(...extra: string[]) {
  const out = join(tmp, 'out.zip')
  // A zip left by an earlier run must not count as this run's output.
  rmSync(out, { force: true })
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

const prompt = (v: string) => ['--prompt', v]
const expectRejected = (r: ReturnType<typeof run>) => {
  expect(r.code).toBe(2)
  expect(r.outExists).toBe(false)
}

describe('make-teams-manifest.sh --prompt', () => {
  test('P1 two prompts only', () => {
    const r = run(...prompt('Summarize: Summarize the tender'), ...prompt('Deadlines: List the deadlines'))
    expect(r.code).toBe(0)
    expect(r.manifest.manifestVersion).toBe('1.29')
    expect(r.manifest.bots[0].commandLists).toEqual([
      {
        scopes: ['personal', 'team', 'groupChat'],
        commands: [
          { title: 'Summarize', type: 'prompt', prompt: 'Summarize the tender', description: 'Summarize the tender' },
          { title: 'Deadlines', type: 'prompt', prompt: 'List the deadlines', description: 'List the deadlines' },
        ],
      },
    ])
    expect('supportsTargetedMessages' in r.manifest.bots[0]).toBe(false)
    expect(r.stdout).toContain('PROMPTS: declares 2 prompt starter(s)')
  })

  test('P2 with one --command', () => {
    const r = run('--command', 'help: Show help', ...prompt('Summarize: Summarize the tender'))
    expect(r.code).toBe(0)
    expect(r.manifest.manifestVersion).toBe('1.29')
    expect(r.manifest.bots[0].supportsTargetedMessages).toBe(true)
    expect(r.manifest.bots[0].commandLists).toEqual([
      {
        scopes: ['personal', 'team', 'groupChat'],
        triggers: ['mention', 'slash'],
        commands: [
          { title: 'help', description: 'Show help' },
          { title: 'Summarize', type: 'prompt', prompt: 'Summarize the tender', description: 'Summarize the tender' },
        ],
      },
    ])
  })

  test('P3 non-ASCII title and escaped prompt round-trip', () => {
    const title = 'Façade ß resume'
    const text = 'Say "hi" and keep a \\ backslash'
    const r = run(...prompt(`${title}: ${text}`))
    expect(r.code).toBe(0)
    const c = r.manifest.bots[0].commandLists[0].commands[0]
    expect(c.title).toBe(title)
    expect(c.prompt).toBe(text)
    expect(c.description).toBe(text)
  })

  test('P3b quotes and a backslash in a title round-trip', () => {
    const r = run(...prompt('Say "hi" \\ now: text'))
    expect(r.code).toBe(0)
    expect(r.manifest.bots[0].commandLists[0].commands[0].title).toBe('Say "hi" \\ now')
  })

  test('P4 split at the first ": "', () => {
    const r = run(...prompt('A: b: c: d'))
    expect(r.code).toBe(0)
    const c = r.manifest.bots[0].commandLists[0].commands[0]
    expect(c.title).toBe('A')
    expect(c.prompt).toBe('b: c: d')
  })

  test('P5 errors', () => {
    expectRejected(run(...prompt('nocolon')))
    expectRejected(run(...prompt(': text')))
    expectRejected(run(...prompt(' Lead: text')))
    expectRejected(run(...prompt('Lead : text')))
    expectRejected(run(...prompt('t: ')))
    expectRejected(run(...prompt(`${'a'.repeat(129)}: text`)))
    expect(run(...prompt(`${'a'.repeat(128)}: text`)).code).toBe(0)
    expectRejected(run(...prompt(`t: ${'a'.repeat(4001)}`)))
    expect(run(...prompt(`t: ${'a'.repeat(4000)}`)).code).toBe(0)
    expectRejected(run(...prompt('t\u0001x: text')))
    expectRejected(run(...prompt('t: te\u0001xt')))
    // C1 control characters (U+0080..U+009F) are refused too.
    const c1Title = run(...prompt('A\u0085B: text'))
    expectRejected(c1Title)
    expect(c1Title.stderr).toContain('invalid --prompt title')
    const c1Prompt = run(...prompt('t: a\u009fb'))
    expectRejected(c1Prompt)
    expect(c1Prompt.stderr).toContain('invalid --prompt prompt')
    // Limits are in bytes: 'ß' is two bytes in UTF-8.
    expect(run(...prompt(`${'ß'.repeat(64)}: text`)).code).toBe(0)
    expectRejected(run(...prompt(`${'ß'.repeat(64)}a: text`)))
    expect(run(...prompt(`t: ${'ß'.repeat(2000)}`)).code).toBe(0)
    expectRejected(run(...prompt(`t: ${'ß'.repeat(2000)}a`)))
    // Bun cannot pass invalid UTF-8 in argv, so let bash build the bytes ($'\xff').
    for (const v of ["$'t\\xffx: text'", "$'t: te\\xffxt'"]) {
      const out = join(tmp, 'out.zip')
      const r = Bun.spawnSync(
        ['bash', '-c', `bash "$0" --app-id ${APP_ID} --name Test --env-file /nonexistent --out "$1" --prompt ${v}`, SCRIPT, out],
        { env: { ...process.env, PATH: `${tmp}:${process.env.PATH}`, HOME: tmp } },
      )
      expect(r.exitCode).toBe(2)
      expect(r.stderr.toString()).toContain('not valid UTF-8')
      expect(existsSync(out)).toBe(false)
    }
    expectRejected(run(...prompt('Dup: a'), ...prompt('dUP: b')))
    expectRejected(run('--command', 'help: Show help', ...prompt('HELP: x')))
    // Titles with spaces are compared whole, not word by word.
    expect(run(...prompt('Summarize tender: x'), ...prompt('tender: y')).code).toBe(0)
    expect(run(...prompt('A: x'), ...prompt('B: y'), ...prompt('A B: z')).code).toBe(0)
    expect(run('--command', 'help: Show help', ...prompt('help me: x')).code).toBe(0)
    const cmds = (n: number) => Array.from({ length: n }, (_, i) => ['--command', `c${i}: d`]).flat()
    const prompts = (n: number) => Array.from({ length: n }, (_, i) => prompt(`p${i}: x`)).flat()
    expect(run(...cmds(5), ...prompts(7)).code).toBe(0)
    expectRejected(run(...cmds(6), ...prompts(7)))
    const r = run('--prompt')
    expectRejected(r)
    expect(r.stderr).toContain('--prompt needs a value')
  }, 60_000)

  test('P6 without --prompt the manifest is unchanged', () => {
    expect(run().manifest.manifestVersion).toBe('1.16')
    expect(run().manifest.bots[0].commandLists).toBeUndefined()
  })
})
