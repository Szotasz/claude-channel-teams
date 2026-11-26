/**
 * TEAMS_BACKFILL / TEAMS_DECISION_LOG configuration: the defaults, the valid ranges, the one-warning fallback for
 * anything out of range (following tests/config-debounce.test.ts: the parsers are pure and take the raw value), and
 * that with both switches off the server wiring creates none of the new files. The server module starts the
 * listener on import, so the wiring is replayed here with the same factory calls, not booted.
 */

import { afterAll, describe, expect, test } from 'bun:test'
import { mkdtempSync, readdirSync, readFileSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import {
  backfillTiming,
  loadConfig,
  parseBackfillGraceMs,
  parseBackfillIntervalMs,
  parseBackfillMaxAgeMs,
  parseDecisionLogMaxBytes,
} from '../src/config'
import { createInboundIds } from '../src/teams/inboundIds'
import { createDecisionLog } from '../src/teams/decisionLog'
import { createBackfill } from '../src/teams/backfill'

const parse = (fn: (raw: string | undefined, warn: (m: string) => void) => number, raw: string | undefined) => {
  const warnings: string[] = []
  const value = fn(raw, m => warnings.push(m))
  return { value, warnings }
}

describe('the backfill seconds parsers', () => {
  test('unset or empty is the default, without a warning', () => {
    for (const raw of [undefined, '', '   ']) {
      expect(parse(parseBackfillIntervalMs, raw)).toEqual({ value: 180_000, warnings: [] })
      expect(parse(parseBackfillGraceMs, raw)).toEqual({ value: 90_000, warnings: [] })
      expect(parse(parseBackfillMaxAgeMs, raw)).toEqual({ value: 7_200_000, warnings: [] })
      expect(parse(parseDecisionLogMaxBytes, raw)).toEqual({ value: 1_048_576, warnings: [] })
    }
  })

  test('valid values are used as-is', () => {
    expect(parse(parseBackfillIntervalMs, '60').value).toBe(60_000)
    expect(parse(parseBackfillIntervalMs, '3600').value).toBe(3_600_000)
    expect(parse(parseBackfillGraceMs, '30').value).toBe(30_000)
    expect(parse(parseBackfillGraceMs, '900').value).toBe(900_000)
    expect(parse(parseBackfillMaxAgeMs, '600').value).toBe(600_000)
    expect(parse(parseBackfillMaxAgeMs, '86400').value).toBe(86_400_000)
    expect(parse(parseDecisionLogMaxBytes, '65536').value).toBe(65_536)
    expect(parse(parseDecisionLogMaxBytes, '16777216').value).toBe(16_777_216)
  })

  test('out of range or not a whole number falls back with exactly one warning', () => {
    expect(parse(parseBackfillIntervalMs, '59')).toEqual({
      value: 180_000,
      warnings: ['teams channel: TEAMS_BACKFILL_INTERVAL_S must be a whole number of seconds from 60 to 3600; using 180'],
    })
    expect(parse(parseBackfillIntervalMs, '3601').warnings).toHaveLength(1)
    expect(parse(parseBackfillIntervalMs, 'abc').warnings).toHaveLength(1)
    expect(parse(parseBackfillIntervalMs, '1.5').warnings).toHaveLength(1)
    // Inside the range but not a whole number: Number() accepts all three, only the digit check refuses them.
    for (const raw of ['120.5', '1e2', '0x64']) {
      const r = parse(parseBackfillIntervalMs, raw)
      expect(r.value).toBe(180_000)
      expect(r.warnings).toHaveLength(1)
    }
    expect(parse(parseBackfillGraceMs, '10')).toEqual({
      value: 90_000,
      warnings: ['teams channel: TEAMS_BACKFILL_GRACE_S must be a whole number of seconds from 30 to 900; using 90'],
    })
    expect(parse(parseBackfillGraceMs, '901').warnings).toHaveLength(1)
    expect(parse(parseBackfillMaxAgeMs, '599')).toEqual({
      value: 7_200_000,
      warnings: ['teams channel: TEAMS_BACKFILL_MAX_AGE_S must be a whole number of seconds from 600 to 86400; using 7200'],
    })
    expect(parse(parseBackfillMaxAgeMs, '100000').warnings).toHaveLength(1)
    expect(parse(parseDecisionLogMaxBytes, '1024')).toEqual({
      value: 1_048_576,
      warnings: ['teams channel: TEAMS_DECISION_LOG_MAX_BYTES must be a whole number of bytes from 65536 to 16777216; using 1048576'],
    })
    expect(parse(parseDecisionLogMaxBytes, '16777217').warnings).toHaveLength(1)
    expect(parse(parseDecisionLogMaxBytes, '1e6').warnings).toHaveLength(1)
  })
})

describe('the backfill max age covers the gap between two cycles', () => {
  const noWarn = (m: string) => {
    throw new Error(`unexpected warning: ${m}`)
  }

  test('the defaults need no change', () => {
    expect(backfillTiming({}, noWarn)).toEqual({ backfillIntervalMs: 180_000, backfillGraceMs: 90_000, backfillMaxAgeMs: 7_200_000 })
  })

  test('a max age exactly twice the interval plus the grace is kept', () => {
    const t = backfillTiming({ TEAMS_BACKFILL_INTERVAL_S: '300', TEAMS_BACKFILL_GRACE_S: '60', TEAMS_BACKFILL_MAX_AGE_S: '660' }, noWarn)
    expect(t.backfillMaxAgeMs).toBe(660_000)
  })

  test('a max age below it is raised to it, with exactly one warning naming the variable', () => {
    const warnings: string[] = []
    const t = backfillTiming({ TEAMS_BACKFILL_GRACE_S: '900', TEAMS_BACKFILL_MAX_AGE_S: '600' }, m => warnings.push(m))
    expect(t).toEqual({ backfillIntervalMs: 180_000, backfillGraceMs: 900_000, backfillMaxAgeMs: 1_260_000 })
    expect(warnings).toHaveLength(1)
    expect(warnings[0]).toContain('TEAMS_BACKFILL_MAX_AGE_S')
    expect(warnings[0]).toContain('using 1260')
  })

  test('a long interval raises the default max age too', () => {
    const warnings: string[] = []
    expect(backfillTiming({ TEAMS_BACKFILL_INTERVAL_S: '3600' }, m => warnings.push(m)).backfillMaxAgeMs).toBe(7_290_000)
    expect(warnings).toHaveLength(1)
  })
})

describe('loadConfig reads the switches', () => {
  const dir = mkdtempSync(join(tmpdir(), 'cct-backfill-cfg-'))
  const vars: Record<string, string | undefined> = {
    TEAMS_BOT_APP_ID: 'app-1',
    TEAMS_BOT_APP_PASSWORD: 'secret',
    TEAMS_BOT_TENANT_ID: 'tenant-1',
    TEAMS_STATE_DIR: dir,
  }
  const set = (k: string, v: string | undefined) => {
    if (v === undefined) delete process.env[k]
    else process.env[k] = v
  }
  const saved = Object.fromEntries(
    [
      ...Object.keys(vars),
      'TEAMS_BACKFILL',
      'TEAMS_DECISION_LOG',
      'TEAMS_BACKFILL_INTERVAL_S',
      'TEAMS_BACKFILL_GRACE_S',
      'TEAMS_BACKFILL_MAX_AGE_S',
    ].map(k => [k, process.env[k]]),
  )
  const restore = () => {
    for (const [k, v] of Object.entries(saved)) set(k, v)
  }
  // A test that writes files into the state dir, or checks that none appear, gets a fresh one: in one shared dir the
  // "switches on" test left its files for the "switches off" test whenever it ran first (bun test --randomize).
  const made: string[] = []
  const freshStateDir = () => {
    const d = mkdtempSync(join(tmpdir(), 'cct-backfill-cfg-'))
    made.push(d)
    return d
  }
  const withVars = (extra: Record<string, string | undefined>, run: () => void) => {
    for (const [k, v] of Object.entries({ ...vars, ...extra })) set(k, v)
    try {
      run()
    } finally {
      restore()
    }
  }

  test('both off unless they are 1 or true', () => {
    withVars({ TEAMS_BACKFILL: undefined, TEAMS_DECISION_LOG: undefined }, () => {
      const c = loadConfig()
      expect(c.backfill).toBe(false)
      expect(c.decisionLog).toBe(false)
      expect(c.backfillIntervalMs).toBe(180_000)
      expect(c.backfillGraceMs).toBe(90_000)
      expect(c.backfillMaxAgeMs).toBe(7_200_000)
      expect(c.decisionLogMaxBytes).toBe(1_048_576)
    })
    withVars({ TEAMS_BACKFILL: 'yes' }, () => expect(loadConfig().backfill).toBe(false))
    withVars({ TEAMS_BACKFILL: '1' }, () => expect(loadConfig().backfill).toBe(true))
    withVars({ TEAMS_BACKFILL: 'true' }, () => expect(loadConfig().backfill).toBe(true))
    withVars({ TEAMS_DECISION_LOG: '1' }, () => expect(loadConfig().decisionLog).toBe(true))
    withVars({ TEAMS_DECISION_LOG: 'true' }, () => expect(loadConfig().decisionLog).toBe(true))
    withVars({ TEAMS_DECISION_LOG: 'off' }, () => expect(loadConfig().decisionLog).toBe(false))
    withVars({ TEAMS_BACKFILL: '1', TEAMS_BACKFILL_INTERVAL_S: '300', TEAMS_BACKFILL_GRACE_S: '45', TEAMS_BACKFILL_MAX_AGE_S: '3600' }, () => {
      const c = loadConfig()
      expect(c.backfillIntervalMs).toBe(300_000)
      expect(c.backfillGraceMs).toBe(45_000)
      expect(c.backfillMaxAgeMs).toBe(3_600_000)
    })
  })

  test('loadConfig raises a max age below twice the interval plus the grace', () => {
    withVars({ TEAMS_BACKFILL_INTERVAL_S: undefined, TEAMS_BACKFILL_GRACE_S: '900', TEAMS_BACKFILL_MAX_AGE_S: '600' }, () => {
      const c = loadConfig()
      expect(c.backfillGraceMs).toBe(900_000)
      expect(c.backfillMaxAgeMs).toBe(1_260_000)
    })
  })

  test('with both switches off the server wiring writes no new file', () => {
    withVars({ TEAMS_BACKFILL: undefined, TEAMS_DECISION_LOG: undefined, TEAMS_STATE_DIR: freshStateDir() }, () => {
      const config = loadConfig()
      // Exactly what src/server.ts does (it cannot be imported here: it boots the listener).
      const inboundIds = createInboundIds()
      if (config.backfill) {
        Object.assign(inboundIds, createInboundIds({ file: join(config.stateDir, 'inbound-ids.json') }))
      }
      const decisions = config.decisionLog
        ? createDecisionLog({ file: join(config.stateDir, 'decisions.log'), maxBytes: config.decisionLogMaxBytes })
        : undefined
      const backfill = config.backfill
        ? createBackfill({
            getToken: async () => 'TOKEN',
            refs: { channelThreadRefs: () => [] },
            teamGroupFor: () => undefined,
            botAppId: config.appId,
            tenantId: config.tenantId,
            stateFile: join(config.stateDir, 'backfill-state.json'),
            intervalMs: config.backfillIntervalMs,
            graceMs: config.backfillGraceMs,
            maxAgeMs: config.backfillMaxAgeMs,
            submit: async () => {},
            decisions,
          })
        : undefined
      expect(backfill).toBeUndefined()
      expect(decisions).toBeUndefined()
      inboundIds.firstSeen('19:abc@thread.tacv2', '1700000000001')
      expect(readdirSync(config.stateDir)).toEqual([])
    })
  })

  test('with the switches on the files appear', () => {
    withVars({ TEAMS_BACKFILL: '1', TEAMS_DECISION_LOG: '1', TEAMS_STATE_DIR: freshStateDir() }, () => {
      const config = loadConfig()
      const inboundIds = createInboundIds()
      if (config.backfill) {
        Object.assign(inboundIds, createInboundIds({ file: join(config.stateDir, 'inbound-ids.json') }))
      }
      const decisions = createDecisionLog({ file: join(config.stateDir, 'decisions.log'), maxBytes: config.decisionLogMaxBytes })
      inboundIds.firstSeen('19:abc@thread.tacv2', '1700000000001')
      decisions.record({ decision: 'deliver', reason: 'ok', source: 'live' })
      expect(readdirSync(config.stateDir)).toContain('inbound-ids.json')
      expect(readdirSync(config.stateDir)).toContain('decisions.log')
    })
  })

  test('src/server.ts carries the wiring the two tests above copy', () => {
    const src = readFileSync(join(import.meta.dir, '../src/server.ts'), 'utf8')
    expect(src).toContain(
      "if (config.backfill) {\n  Object.assign(inboundIds, createInboundIds({ file: join(config.stateDir, 'inbound-ids.json') }))\n}",
    )
    expect(src).toContain(
      "const decisions = config.decisionLog\n  ? createDecisionLog({ file: join(config.stateDir, 'decisions.log'), maxBytes: config.decisionLogMaxBytes })\n  : undefined",
    )
    expect(src).toContain('const backfill = config.backfill\n  ? createBackfill({')
    expect(src).toContain('      submit: activity => listener.submitActivity(activity),')
  })

  afterAll(() => {
    for (const d of [dir, ...made]) rmSync(d, { recursive: true, force: true })
  })
})
