import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { loadConfig } from '../src/config'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import {
  createCommandFilter,
  createHeldMessages,
  withHeldContext,
  type FilterDecision,
  type UnmentionedFilter,
  type UnmentionedFilterInput,
} from '../src/teams/unmentionedFilter'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const UNKNOWN_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const CHAN1 = '19:chan1@thread.tacv2;messageid=111'

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function makeActivity(overrides: Partial<Activity> = {}): Activity {
  return {
    type: 'message',
    text: 'hello there',
    id: 'msg-1',
    timestamp: new Date('2026-05-18T19:43:00Z'),
    from: { id: 'user-1', name: 'Adnan Khan', aadObjectId: FIXTURE_ID },
    conversation: { id: CHAN1, tenantId: TENANT_ID, conversationType: 'channel' },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/uk/',
    ...overrides,
  } as unknown as Activity
}

/** A message that @mentions the bot (recipient id 'bot-1'), like the mentioned ones in unmentioned.test.ts. */
function mentionedActivity(overrides: Partial<Activity> = {}): Activity {
  return makeActivity({
    text: '<at>Bot</at> hello',
    entities: [{ type: 'mention', text: '<at>Bot</at>', mentioned: { id: 'bot-1', name: 'Bot' } }] as any,
    ...overrides,
  })
}

function setup(unmentioned?: Set<string>, extra: Record<string, unknown> = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'cct-unmentioned-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(FIXTURE_ID)
  const refs = createConversationRefStore()
  const events: any[] = []
  const config = {
    stateDir: '/tmp/unused',
    allowlistFile: '/tmp/unused',
    appId: 'app-id',
    appPassword: 'app-pw',
    appType: 'SingleTenant' as const,
    tenantId: TENANT_ID,
    bindHost: '127.0.0.1',
    port: 0,
    receivedFilesDir: '/tmp/unused',
    ...(unmentioned ? { unmentionedConversations: unmentioned } : {}),
  }
  const turn = makeTurnHandler({
    config: config as any,
    adapter: {} as any,
    allowlist,
    refs,
    onEvent: e => events.push(e),
    ...extra,
  } as any)
  const run = (a: Activity) => turn(new TurnContext(new NoopAdapter() as any, a))
  return { run, events, refs }
}

const chan2 = (o: Partial<Activity> = {}) =>
  makeActivity({
    conversation: { id: '19:chan2@thread.tacv2;messageid=5', tenantId: TENANT_ID, conversationType: 'channel' } as any,
    ...o,
  })

const holdAll: UnmentionedFilter = async () => ({ verdict: 'hold' })

describe('unmentioned filter', () => {
  test('no filter deps: unmentioned message is delivered with its text unchanged', async () => {
    const s = setup(new Set(['*']))
    await s.run(makeActivity())
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('hello there')
  })

  test('heldMessages without unmentionedFilter changes nothing (the filter needs both)', async () => {
    const s = setup(new Set(['*']), { heldMessages: createHeldMessages(20) })
    await s.run(makeActivity())
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('hello there')
  })

  test('a hold verdict delivers nothing; the message rides along with the next mentioned one', async () => {
    const typingStarts: any[] = []
    const attachmentCalls: any[] = []
    const s = setup(new Set(['*']), {
      unmentionedFilter: holdAll,
      heldMessages: createHeldMessages(20),
      typingPump: { start: (...a: any[]) => { typingStarts.push(a) }, stop() {}, stopAll() {} },
      attachments: async (...a: any[]) => { attachmentCalls.push(a); return { annotations: [] } },
    })
    const firstText = 'see this file'
    await s.run(makeActivity({
      text: firstText,
      attachments: [{ contentType: 'text/html', content: 'x' }] as any,
    }))
    expect(s.events.length).toBe(0)
    expect(typingStarts.length).toBe(0)
    expect(attachmentCalls.length).toBe(0)
    const mentionedText = '<at>Bot</at> what do you think?'
    await s.run(mentionedActivity({ text: mentionedText, id: 'msg-2' }))
    expect(s.events.length).toBe(1)
    const expected = [
      '[Earlier messages in this conversation that were not addressed to you, oldest first]',
      `Adnan Khan: ${firstText} [1 attachment(s), not downloaded]`,
      '[End of earlier messages]',
      '',
      mentionedText.replace('<at>Bot</at>', '').trim(),
    ].join('\n')
    expect(s.events[0].text).toBe(expected)
  })

  test('a deliver verdict delivers unchanged, and the filter sees exactly the input', async () => {
    const seen: UnmentionedFilterInput[] = []
    const filter: UnmentionedFilter = async input => { seen.push(input); return { verdict: 'deliver' } }
    const s = setup(new Set(['*']), { unmentionedFilter: filter, heldMessages: createHeldMessages(20) })
    await s.run(makeActivity({ text: 'plain question' }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('plain question')
    expect(seen).toEqual([
      { text: 'plain question', fromName: 'Adnan Khan', conversationId: CHAN1, conversationType: 'channel', botName: 'Bot' },
    ])
  })

  test('a deliver verdict with a reason still delivers (fail-open)', async () => {
    const s = setup(new Set(['*']), {
      unmentionedFilter: async () => ({ verdict: 'deliver', reason: 'exit code 3' }),
      heldMessages: createHeldMessages(20),
    })
    await s.run(makeActivity())
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('hello there')
  })

  test('a filter that throws still delivers', async () => {
    const s = setup(new Set(['*']), {
      unmentionedFilter: async () => { throw new Error('boom') },
      heldMessages: createHeldMessages(20),
    })
    await s.run(makeActivity())
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('hello there')
  })

  test('the filter is never called for a mentioned, a 1:1, a non-allowlisted, or a non-opted-in message', async () => {
    const seen: UnmentionedFilterInput[] = []
    const filter: UnmentionedFilter = async input => { seen.push(input); return { verdict: 'hold' } }
    const held = createHeldMessages(20)
    const s = setup(new Set(['*']), { unmentionedFilter: filter, heldMessages: held })
    await s.run(mentionedActivity())
    expect(seen.length).toBe(0)
    await s.run(makeActivity({
      conversation: { id: 'a:personal-1', tenantId: TENANT_ID, conversationType: 'personal' } as any,
    }))
    expect(seen.length).toBe(0)
    await s.run(makeActivity({ from: { id: 'u2', name: 'X', aadObjectId: UNKNOWN_ID } as any }))
    expect(seen.length).toBe(0)
    const other = setup(new Set(['19:other@thread.tacv2']), { unmentionedFilter: filter, heldMessages: held })
    await other.run(makeActivity())
    expect(seen.length).toBe(0)
  })

  test('two held messages ride along oldest-first, and are taken once', async () => {
    const s = setup(new Set(['*']), { unmentionedFilter: holdAll, heldMessages: createHeldMessages(20) })
    await s.run(makeActivity({ text: 'a', id: 'm1' }))
    await s.run(makeActivity({ text: 'b', id: 'm2' }))
    expect(s.events.length).toBe(0)
    await s.run(mentionedActivity({ text: '<at>Bot</at> c', id: 'm3' }))
    expect(s.events.length).toBe(1)
    const text = s.events[0].text as string
    expect(text).toContain('Adnan Khan: a')
    expect(text).toContain('Adnan Khan: b')
    expect(text.indexOf('Adnan Khan: a')).toBeLessThan(text.indexOf('Adnan Khan: b'))
    await s.run(mentionedActivity({ text: '<at>Bot</at> d', id: 'm4' }))
    expect(s.events.length).toBe(2)
    expect(s.events[1].text).toBe('d')
  })

  test('held messages do not leak into another conversation', async () => {
    const s = setup(new Set(['*']), { unmentionedFilter: holdAll, heldMessages: createHeldMessages(20) })
    await s.run(makeActivity({ text: 'a' }))
    await s.run(mentionedActivity(chan2({ id: 'm2', text: '<at>Bot</at> hi' })))
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).not.toContain('[Earlier messages')
    expect(s.events[0].text).toBe('hi')
  })

  test('held messages stay in their thread: another thread of the same channel does not see them', async () => {
    const s = setup(new Set(['*']), { unmentionedFilter: holdAll, heldMessages: createHeldMessages(20) })
    await s.run(makeActivity({ text: 'a' }))
    const otherThread = { id: '19:chan1@thread.tacv2;messageid=222', tenantId: TENANT_ID, conversationType: 'channel' } as any
    await s.run(mentionedActivity({ id: 'm2', text: '<at>Bot</at> hi', conversation: otherThread }))
    expect(s.events.map(e => e.text)).toEqual(['hi'])
  })

  test('a second unmentioned message is judged after the first, and a later mention waits for both', async () => {
    const resolvers: Record<string, (d: FilterDecision) => void> = {}
    const filter: UnmentionedFilter = input => new Promise<FilterDecision>(res => { resolvers[input.text] = res })
    const s = setup(new Set(['*']), { unmentionedFilter: filter, heldMessages: createHeldMessages(20) })
    const tick = () => new Promise(r => setTimeout(r, 20))
    const runA = s.run(makeActivity({ text: 'a' }))
    const runB = s.run(makeActivity({ id: 'msg-2', text: 'b' }))
    await tick()
    expect(Object.keys(resolvers)).toEqual(['a'])
    resolvers.a!({ verdict: 'hold' })
    await runA
    await tick()
    expect(Object.keys(resolvers)).toEqual(['a', 'b'])
    const runC = s.run(mentionedActivity({ id: 'm3', text: '<at>Bot</at> c' }))
    await tick()
    expect(s.events.length).toBe(0)
    resolvers.b!({ verdict: 'hold' })
    await runB
    await runC
    expect(s.events.map(e => e.text)).toEqual([
      '[Earlier messages in this conversation that were not addressed to you, oldest first]\nAdnan Khan: a\nAdnan Khan: b\n[End of earlier messages]\n\nc',
    ])
  })

  test('createCommandFilter kills a filter that runs past the timeout', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'cct-ufilter-'))
    try {
      const pidFile = join(dir, 'pid')
      const slow = join(dir, 'slow.sh')
      writeFileSync(slow, `#!/bin/sh\necho $$ > "${pidFile}"\nexec sleep 5\n`, { mode: 0o755 })
      const d = await createCommandFilter(slow, 500)({ text: 'x', conversationId: 'c1' })
      expect(d).toEqual({ verdict: 'deliver', reason: 'timed out after 500 ms' })
      const pid = Number(readFileSync(pidFile, 'utf8').trim())
      await new Promise(r => setTimeout(r, 200))
      expect(() => process.kill(pid, 0)).toThrow()
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test('a mentioned message waits for the pending filter decision of an earlier unmentioned one', async () => {
    let resolveFilter!: (d: FilterDecision) => void
    const filter: UnmentionedFilter = () => new Promise<FilterDecision>(res => { resolveFilter = res })
    const s = setup(new Set(['*']), { unmentionedFilter: filter, heldMessages: createHeldMessages(20) })
    const unmentionedRun = s.run(makeActivity({ text: 'a' }))
    const mentionedRun = s.run(mentionedActivity({ text: '<at>Bot</at> c', id: 'm2' }))
    await new Promise(r => setTimeout(r, 20))
    expect(s.events.length).toBe(0)
    resolveFilter({ verdict: 'hold' })
    await unmentionedRun
    await mentionedRun
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toContain('Adnan Khan: a')
  })

  test('a message held while an earlier one waits for its attachments is not shown as earlier context', async () => {
    const filter: UnmentionedFilter = async input =>
      input.text === 'a' ? { verdict: 'deliver' } : { verdict: 'hold' }
    let releaseAttachments!: () => void
    const attachments = () =>
      new Promise<{ annotations: string[] }>(res => {
        releaseAttachments = () => res({ annotations: [] })
      })
    const s = setup(new Set(['*']), { unmentionedFilter: filter, heldMessages: createHeldMessages(20), attachments })
    const runA = s.run(makeActivity({ text: 'a', attachments: [{ contentType: 'text/html', content: 'x' }] as any }))
    await new Promise(r => setTimeout(r, 20))
    // 'b' arrives while 'a' is still waiting for its attachment download — held by the filter.
    await s.run(makeActivity({ id: 'msg-2', text: 'b' }))
    releaseAttachments()
    await runA
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('a')
    // 'b' waited for the NEXT message; it must not be glued onto 'a' retroactively.
    await s.run(mentionedActivity({ id: 'm3', text: '<at>Bot</at> c' }))
    expect(s.events.length).toBe(2)
    expect(s.events[1].text).toBe(
      '[Earlier messages in this conversation that were not addressed to you, oldest first]\nAdnan Khan: b\n[End of earlier messages]\n\nc',
    )
  })

  test('a held message rides along with a later unmentioned message the filter lets through', async () => {
    const filter: UnmentionedFilter = async input =>
      input.text === 'a' ? { verdict: 'hold' } : { verdict: 'deliver' }
    const s = setup(new Set(['*']), { unmentionedFilter: filter, heldMessages: createHeldMessages(20) })
    await s.run(makeActivity({ text: 'a' }))
    await s.run(makeActivity({ id: 'msg-2', text: 'b' }))
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe(
      '[Earlier messages in this conversation that were not addressed to you, oldest first]\nAdnan Khan: a\n[End of earlier messages]\n\nb',
    )
  })

  test('createHeldMessages keeps only the last N lines of a conversation, oldest first', () => {
    const held = createHeldMessages(3)
    for (const t of ['1', '2', '3', '4', '5']) held.hold('c', 'Adnan Khan', t)
    expect(held.take('c')).toEqual(['Adnan Khan: 3', 'Adnan Khan: 4', 'Adnan Khan: 5'])
    expect(held.take('c')).toEqual([])
  })

  test('createHeldMessages truncates at 1000 chars and names an empty sender Someone', () => {
    const held = createHeldMessages(20)
    const long = 'x'.repeat(1500)
    held.hold('c', 'Adnan Khan', long)
    const [line] = held.take('c')
    expect(line).toBe(`Adnan Khan: ${'x'.repeat(999)}…`)
    expect(line.slice('Adnan Khan: '.length)).toHaveLength(1000)
    held.hold('c', undefined, 'hi')
    expect(held.take('c')).toEqual(['Someone: hi'])
    // The adapter passes '' for a sender without a name (a.from?.name ?? '').
    held.hold('c', '', 'hi')
    expect(held.take('c')).toEqual(['Someone: hi'])
  })

  test('createHeldMessages puts the attachment note after the truncated text', () => {
    const held = createHeldMessages(20)
    held.hold('c', 'Adnan Khan', 'x'.repeat(1500), 2)
    expect(held.take('c')).toEqual([`Adnan Khan: ${'x'.repeat(999)}… [2 attachment(s), not downloaded]`])
    held.hold('c', 'Adnan Khan', '', 1)
    expect(held.take('c')).toEqual(['Adnan Khan: [1 attachment(s), not downloaded]'])
    held.hold('c', 'Adnan Khan', 'hi', 0)
    expect(held.take('c')).toEqual(['Adnan Khan: hi'])
  })

  test('createHeldMessages take(id, upTo) leaves the later ones', () => {
    const held = createHeldMessages(20)
    held.hold('c', 'Adnan Khan', 'a')
    const m = held.mark()
    held.hold('c', 'Adnan Khan', 'b')
    expect(held.take('c', m)).toEqual(['Adnan Khan: a'])
    expect(held.take('c')).toEqual(['Adnan Khan: b'])
    expect(held.take('c')).toEqual([])
  })

  test('createHeldMessages caps the conversation count; re-holding moves a conversation to the back', () => {
    const held = createHeldMessages(20)
    for (let i = 0; i <= 200; i++) held.hold(`c${i}`, 'Adnan Khan', 'x')
    expect(held.take('c0')).toEqual([])
    expect(held.take('c200')).toEqual(['Adnan Khan: x'])

    const fresh = createHeldMessages(20)
    for (let i = 0; i <= 199; i++) fresh.hold(`c${i}`, 'Adnan Khan', 'x')
    fresh.hold('c0', 'Adnan Khan', 'again')
    fresh.hold('c200', 'Adnan Khan', 'x')
    expect(fresh.take('c1')).toEqual([])
    expect(fresh.take('c0')).toEqual(['Adnan Khan: x', 'Adnan Khan: again'])
  })

  test('withHeldContext with no held lines is the identity', () => {
    expect(withHeldContext('t', [])).toBe('t')
  })

  test('createCommandFilter: exit codes, signals, timeouts, spawn errors and stdin', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'cct-ufilter-'))
    try {
      const write = (name: string, body: string) => {
        const p = join(dir, name)
        writeFileSync(p, body, { mode: 0o755 })
        return p
      }
      const input: UnmentionedFilterInput = {
        text: 'hi',
        fromName: 'Adnan Khan',
        conversationId: 'c1',
        conversationType: 'channel',
        botName: 'Bot',
      }
      const run = (cmd: string, timeoutMs = 5000) => createCommandFilter(cmd, timeoutMs)(input)

      const d0 = await run(write('exit0.sh', '#!/bin/sh\nexit 0\n'))
      expect(d0.verdict).toBe('deliver')
      expect('reason' in d0).toBe(false)

      expect(await run(write('exit1.sh', '#!/bin/sh\nexit 1\n'))).toEqual({ verdict: 'hold' })

      const d3 = await run(write('exit3.sh', '#!/bin/sh\nexit 3\n'))
      expect(d3.verdict).toBe('deliver')
      expect(d3.reason).toBe('exit code 3')

      const dkill = await run(write('kill.sh', '#!/bin/sh\nkill -9 $$\n'))
      expect(dkill.verdict).toBe('deliver')
      expect(dkill.reason).toBe('killed by SIGKILL')

      const startedAt = Date.now()
      const dtimed = await run(write('sleep.sh', '#!/bin/sh\nsleep 5\n'), 200)
      expect(dtimed.verdict).toBe('deliver')
      expect(dtimed.reason).toBe('timed out after 200 ms')
      expect(Date.now() - startedAt).toBeLessThan(2000)

      const dmissing = await run(join(dir, 'does-not-exist'))
      expect(dmissing.verdict).toBe('deliver')
      expect(dmissing.reason?.startsWith('spawn failed')).toBe(true)

      const inJson = join(dir, 'in.json')
      const dcat = await run(write('cat.sh', `#!/bin/sh\ncat > ${JSON.stringify(inJson)}\nexit 1\n`))
      expect(dcat.verdict).toBe('hold')
      expect(JSON.parse(readFileSync(inJson, 'utf8'))).toEqual(input)

      const dbig = await createCommandFilter(write('noread.sh', '#!/bin/sh\nexit 1\n'), 5000)({
        text: 'x'.repeat(300000),
        conversationId: 'c1',
      })
      expect(dbig.verdict).toBe('hold')
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test('loadConfig reads the unmentioned-filter variables from process.env', () => {
    const dir = mkdtempSync(join(tmpdir(), 'cct-ufilter-cfg-'))
    const keys = [
      'TEAMS_BOT_APP_ID',
      'TEAMS_BOT_APP_PASSWORD',
      'TEAMS_BOT_TENANT_ID',
      'TEAMS_STATE_DIR',
      'TEAMS_UNMENTIONED_FILTER_CMD',
      'TEAMS_UNMENTIONED_FILTER_TIMEOUT_MS',
      'TEAMS_UNMENTIONED_HISTORY_LIMIT',
    ]
    const saved: Record<string, string | undefined> = {}
    for (const k of keys) saved[k] = process.env[k]
    const set = (k: string, v: string | undefined) => {
      if (v === undefined) delete process.env[k]
      else process.env[k] = v
    }
    const base: Record<string, string> = {
      TEAMS_BOT_APP_ID: 'a',
      TEAMS_BOT_APP_PASSWORD: 'p',
      TEAMS_BOT_TENANT_ID: 't',
      TEAMS_STATE_DIR: dir,
    }
    const apply = (extra: Record<string, string>) => {
      for (const k of keys) set(k, undefined)
      for (const [k, v] of Object.entries({ ...base, ...extra })) set(k, v)
    }
    try {
      apply({})
      let cfg = loadConfig()
      expect(cfg.unmentionedFilterCmd).toBeUndefined()
      expect(cfg.unmentionedFilterTimeoutMs).toBe(10000)
      expect(cfg.unmentionedHistoryLimit).toBe(20)

      apply({ TEAMS_UNMENTIONED_FILTER_CMD: '/usr/local/bin/f' })
      cfg = loadConfig()
      expect(cfg.unmentionedFilterCmd).toBe('/usr/local/bin/f')

      apply({ TEAMS_UNMENTIONED_FILTER_CMD: 'relative/f' })
      cfg = loadConfig()
      expect(cfg.unmentionedFilterCmd).toBeUndefined()

      for (const bad of ['0', 'abc', '10abc', '1e3']) {
        apply({ TEAMS_UNMENTIONED_FILTER_TIMEOUT_MS: bad })
        expect(() => loadConfig()).toThrow(/TEAMS_UNMENTIONED_FILTER_TIMEOUT_MS must be a positive integer/)
      }

      apply({ TEAMS_UNMENTIONED_HISTORY_LIMIT: '0' })
      expect(() => loadConfig()).toThrow(/TEAMS_UNMENTIONED_HISTORY_LIMIT must be a positive integer/)

      apply({ TEAMS_UNMENTIONED_HISTORY_LIMIT: '5' })
      cfg = loadConfig()
      expect(cfg.unmentionedHistoryLimit).toBe(5)
    } finally {
      for (const k of keys) set(k, saved[k])
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
