import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { createSendTracker, sendTracker, SHUTDOWN_SEND_WAIT_MS, withSendRetry } from '../src/teams/sendRetry.js'

const deps = (tracker: ReturnType<typeof createSendTracker>) => ({
  tracker,
  sleep: async () => {},
  random: () => 0,
  log: () => {},
})

/** Races a promise against a timer, so the tests compare against the tracker's own timeout instead of wall-clock time. */
const before = <T>(p: Promise<T>, ms: number) =>
  Promise.race([p, new Promise<'timer'>(r => setTimeout(() => r('timer'), ms))])

describe('createSendTracker', () => {
  test('idle resolves at once with nothing in flight', async () => {
    const t = createSendTracker()
    // Resolved before any timer can fire, so it does not wait for its own 1000 ms.
    expect(await before(t.idle(1000), 0)).toBe(true)
  })
  test('idle resolves false after the timeout', async () => {
    const t = createSendTracker()
    t.start()
    const t0 = performance.now()
    expect(await t.idle(30)).toBe(false)
    expect(performance.now() - t0).toBeGreaterThanOrEqual(25)
  })
  test('idle resolves true when the send finishes, before its own timeout', async () => {
    const t = createSendTracker()
    const done = t.start()
    setTimeout(done, 20)
    expect(await before(t.idle(5000), 2000)).toBe(true)
  })
  test('done twice counts once', () => {
    const t = createSendTracker()
    const done = t.start()
    done()
    done()
    expect(t.inFlight()).toBe(0)
  })
  test('two sends: idle waits for both', async () => {
    const t = createSendTracker()
    const d1 = t.start()
    const d2 = t.start()
    let resolved = false
    const p = t.idle(1000).then(v => {
      resolved = true
      return v
    })
    d1()
    await new Promise(r => setTimeout(r, 10))
    expect(resolved).toBe(false)
    d2()
    expect(await p).toBe(true)
  })
})

describe('withSendRetry tracking', () => {
  test('in flight while fn runs, 0 after success', async () => {
    const t = createSendTracker()
    let seen = -1
    const r = await withSendRetry('x', async () => {
      seen = t.inFlight()
      return 7
    }, deps(t))
    expect(r).toBe(7)
    expect(seen).toBe(1)
    expect(t.inFlight()).toBe(0)
  })
  test('0 after non-retryable failure', async () => {
    const t = createSendTracker()
    await expect(withSendRetry('x', async () => { throw new Error('boom') }, deps(t))).rejects.toThrow('boom')
    expect(t.inFlight()).toBe(0)
  })
  test('0 after three 412 failures', async () => {
    const t = createSendTracker()
    let calls = 0
    await expect(withSendRetry('x', async () => { calls++; throw Object.assign(new Error('412'), { statusCode: 412 }) }, deps(t))).rejects.toThrow('412')
    expect(calls).toBe(3)
    expect(t.inFlight()).toBe(0)
  })
  test('still counted while waiting between attempts', async () => {
    const t = createSendTracker()
    const seen: number[] = []
    let calls = 0
    await expect(withSendRetry('x', async () => { calls++; throw Object.assign(new Error('412'), { statusCode: 412 }) }, {
      ...deps(t),
      // The wait between attempts: the send is still in progress here.
      sleep: async () => { seen.push(t.inFlight()) },
    })).rejects.toThrow('412')
    expect(seen).toEqual([1, 1])
    expect(t.inFlight()).toBe(0)
  })
  test('without an injected tracker the module-level one counts (the one shutdown waits for)', async () => {
    const before = sendTracker.inFlight()
    let inside = -1
    await withSendRetry('x', async () => { inside = sendTracker.inFlight() }, { sleep: async () => {}, random: () => 0, log: () => {} })
    expect(inside).toBe(before + 1)
    expect(sendTracker.inFlight()).toBe(before)
  })
  test('module-level tracker untouched when one is injected', async () => {
    const t = createSendTracker()
    await withSendRetry('x', async () => {
      expect(sendTracker.inFlight()).toBe(0)
    }, deps(t))
    expect(sendTracker.inFlight()).toBe(0)
  })
})

test('SHUTDOWN_SEND_WAIT_MS is 5000', () => {
  expect(SHUTDOWN_SEND_WAIT_MS).toBe(5000)
})

test('shutdown() waits for sends after the listener stops and before exiting', () => {
  const src = readFileSync(new URL('../src/server.ts', import.meta.url), 'utf8')
  const start = src.indexOf('async function shutdown()')
  expect(start).toBeGreaterThan(-1)
  // Only the function's own body: up to the first line that closes a top-level block after it.
  const body = src.slice(start, src.indexOf('\n}\n', start) + 3)
  const stop = body.indexOf('await listener.stop()')
  const idle = body.indexOf('await sendTracker.idle(SHUTDOWN_SEND_WAIT_MS)')
  const exit = body.indexOf('setTimeout(() => process.exit(0), 500)')
  expect(stop).toBeGreaterThan(-1)
  expect(idle).toBeGreaterThan(stop)
  expect(exit).toBeGreaterThan(idle)
  // When the cap runs out, one line says how many sends were still running.
  const line = body.indexOf('process.stderr.write(`teams channel: exiting with ${sendTracker.inFlight()} send(s) still in progress after ${SHUTDOWN_SEND_WAIT_MS} ms\\n`)')
  expect(line).toBeGreaterThan(idle)
  expect(exit).toBeGreaterThan(line)
})
