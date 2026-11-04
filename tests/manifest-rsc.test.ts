import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { chmodSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

const SCRIPT = resolve(import.meta.dir, '..', 'scripts', 'make-teams-manifest.sh')
const APP_ID = '11111111-2222-4333-8444-555555555555'

let tmp: string

beforeEach(() => {
  tmp = mkdtempSync(join(tmpdir(), 'manifest-rsc-'))
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
  return { code: r.exitCode, stdout: r.stdout.toString(), stderr: r.stderr.toString(), manifest }
}

/** A golden manifest: the bytes the script writes for a given option set. */
const golden = (name: string) => resolve(import.meta.dir, 'fixtures', 'manifest-golden', `${name}.json`)

describe('make-teams-manifest.sh --rsc-all-messages', () => {
  test('T1 without the flag', () => {
    const r = run()
    expect(r.code).toBe(0)
    expect(r.manifest.webApplicationInfo).toBeUndefined()
    expect(r.manifest.authorization).toBeUndefined()
    expect('authorization' in r.manifest).toBe(false)
    expect(r.manifest.bots[0].botId).toBe(APP_ID)
    expect(r.manifest.id).toBe(APP_ID)
    expect(r.stdout.split('\n').some(l => l.startsWith('RSC:'))).toBe(false)
  })

  test('T2 with the flag', () => {
    const r = run('--rsc-all-messages')
    expect(r.code).toBe(0)
    expect(r.manifest.webApplicationInfo).toEqual({ id: APP_ID, resource: 'https://RscBasedStoreApp' })
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: 'ChannelMessage.Read.Group', type: 'Application' },
      { name: 'ChatMessage.Read.Chat', type: 'Application' },
    ])
    expect(r.manifest.manifestVersion).toBe('1.16')
    expect(r.stdout.split('\n').some(l => l.startsWith('RSC:'))).toBe(true)
  })

  test('T3 the flag adds exactly webApplicationInfo and authorization', () => {
    const plain = run().manifest
    const rsc = run('--rsc-all-messages').manifest
    delete rsc.webApplicationInfo
    delete rsc.authorization
    expect(rsc).toEqual(plain)
  })

  test('T4 unknown option exits 2', () => {
    const r = Bun.spawnSync(['bash', SCRIPT, '--bogus'], { env: { ...process.env, HOME: tmp } })
    expect(r.exitCode).toBe(2)
  })
})

describe('make-teams-manifest.sh --rsc-meeting-events', () => {
  const MEETING = [
    { name: 'OnlineMeeting.ReadBasic.Chat', type: 'Application' },
    { name: 'ChannelMeeting.ReadBasic.Group', type: 'Application' },
  ]
  const MESSAGES = [
    { name: 'ChannelMessage.Read.Group', type: 'Application' },
    { name: 'ChatMessage.Read.Chat', type: 'Application' },
  ]

  test('T5 alone: webApplicationInfo and exactly the two meeting permissions', () => {
    const r = run('--rsc-meeting-events')
    expect(r.code).toBe(0)
    expect(r.manifest.webApplicationInfo).toEqual({ id: APP_ID, resource: 'https://RscBasedStoreApp' })
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual(MEETING)
    expect(r.manifest.manifestVersion).toBe('1.16')
    expect(r.stdout.split('\n').some(l => l.startsWith('RSC:') && l.includes('OnlineMeeting.ReadBasic.Chat'))).toBe(true)
  })

  test('T6 with --rsc-all-messages too: all four, messages first', () => {
    const r = run('--rsc-all-messages', '--rsc-meeting-events')
    expect(r.code).toBe(0)
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual([...MESSAGES, ...MEETING])
  })

  test('T7 --rsc-all-messages alone is unchanged: exactly the two message permissions', () => {
    const r = run('--rsc-all-messages')
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual(MESSAGES)
    expect(r.stdout.split('\n').some(l => l.startsWith('RSC:') && l.includes('OnlineMeeting.ReadBasic.Chat'))).toBe(false)
  })

  test('T8 with no option there is no authorization key', () => {
    const r = run()
    expect(r.manifest.webApplicationInfo).toBeUndefined()
    expect('authorization' in r.manifest).toBe(false)
  })
})

describe('make-teams-manifest.sh --rsc-meeting-transcripts', () => {
  const TRANSCRIPT = [{ name: 'OnlineMeetingTranscript.Read.Chat', type: 'Application' }]

  test('T9 alone: webApplicationInfo and exactly the transcript permission', () => {
    const r = run('--rsc-meeting-transcripts')
    expect(r.code).toBe(0)
    expect(r.manifest.webApplicationInfo).toEqual({ id: APP_ID, resource: 'https://RscBasedStoreApp' })
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual(TRANSCRIPT)
    expect(r.manifest.manifestVersion).toBe('1.16')
    expect(r.stdout.split('\n').some(l => l.startsWith('RSC:') && l.includes('OnlineMeetingTranscript.Read.Chat'))).toBe(true)
  })

  test('T10 all three options: the five permissions, the transcript one last', () => {
    const r = run('--rsc-all-messages', '--rsc-meeting-events', '--rsc-meeting-transcripts')
    expect(r.code).toBe(0)
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: 'ChannelMessage.Read.Group', type: 'Application' },
      { name: 'ChatMessage.Read.Chat', type: 'Application' },
      { name: 'OnlineMeeting.ReadBasic.Chat', type: 'Application' },
      { name: 'ChannelMeeting.ReadBasic.Group', type: 'Application' },
      ...TRANSCRIPT,
    ])
  })

  test('T11 with --rsc-meeting-events: the three meeting permissions in order', () => {
    const r = run('--rsc-meeting-events', '--rsc-meeting-transcripts')
    expect(r.code).toBe(0)
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual([
      { name: 'OnlineMeeting.ReadBasic.Chat', type: 'Application' },
      { name: 'ChannelMeeting.ReadBasic.Group', type: 'Application' },
      ...TRANSCRIPT,
    ])
  })

  test('T12 alone it warns on stderr, and the manifest is still written', () => {
    const r = run('--rsc-meeting-transcripts')
    expect(r.code).toBe(0)
    expect(r.stderr).toContain('warning: --rsc-meeting-transcripts without --rsc-meeting-events')
    expect(r.manifest.authorization.permissions.resourceSpecific).toEqual(TRANSCRIPT)
  })

  test('T13 with --rsc-meeting-events there is no such warning', () => {
    const r = run('--rsc-meeting-events', '--rsc-meeting-transcripts')
    expect(r.code).toBe(0)
    expect(r.stderr).not.toContain('--rsc-meeting-transcripts without')
  })
})

describe('make-teams-manifest.sh without --rsc-meeting-transcripts is byte-identical', () => {
  test('G1 no option', () => {
    const r = run()
    expect(r.code).toBe(0)
    expect(readFileSync(join(tmp, 'out.zip'), 'utf8')).toBe(readFileSync(golden('none'), 'utf8'))
  })

  test('G2 --rsc-all-messages', () => {
    const r = run('--rsc-all-messages')
    expect(r.code).toBe(0)
    expect(readFileSync(join(tmp, 'out.zip'), 'utf8')).toBe(readFileSync(golden('all-messages'), 'utf8'))
  })

  test('G3 --rsc-meeting-events', () => {
    const r = run('--rsc-meeting-events')
    expect(r.code).toBe(0)
    expect(readFileSync(join(tmp, 'out.zip'), 'utf8')).toBe(readFileSync(golden('meeting-events'), 'utf8'))
  })

  test('G4 --rsc-all-messages --rsc-meeting-events', () => {
    const r = run('--rsc-all-messages', '--rsc-meeting-events')
    expect(r.code).toBe(0)
    expect(readFileSync(join(tmp, 'out.zip'), 'utf8')).toBe(
      readFileSync(golden('all-messages-meeting-events'), 'utf8'),
    )
  })

  test('G5 --command', () => {
    const r = run('--command', 'help: Show help')
    expect(r.code).toBe(0)
    expect(readFileSync(join(tmp, 'out.zip'), 'utf8')).toBe(readFileSync(golden('command'), 'utf8'))
  })
})
