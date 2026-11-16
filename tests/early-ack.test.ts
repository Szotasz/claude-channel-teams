import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { ACK_AFTER_MS, makeEarlyAckTurn } from '../src/teams/earlyAck.js'

const msg = (conv: string, id: string) => ({ type: 'message', id, conversation: { id: conv } }) as any
const sleep = (ms: number) => new Promise(r => setTimeout(r, ms))

/** Polls until `cond` holds; fails the test after `ms` instead of relying on a fixed wait. */
async function waitFor(cond: () => boolean, ms = 2000): Promise<void> {
  const until = Date.now() + ms
  while (!cond()) {
    if (Date.now() > until) throw new Error('condition not met in time')
    await sleep(5)
  }
}

function setup(d: Record<string, number> = {}, fail = new Set<string>()) {
  const events: string[] = []
  const lines: string[] = []
  const tails = new Map<string, Promise<void>>()
  const turn = async ({ activity }: { activity: any }) => {
    events.push(`${activity.id}:start`)
    // A turn without a delay finishes in microtasks, before any timer can fire.
    if (d[activity.id]) await sleep(d[activity.id]!)
    events.push(`${activity.id}:end`)
    if (fail.has(activity.id)) throw new Error('boom')
  }
  const earlyAck = makeEarlyAckTurn(turn as any, { ackAfterMs: 30, log: l => { lines.push(l) }, tails })
  return { events, lines, tails, earlyAck }
}

describe('early ack', () => {
  test('a) fast turn: answered after the turn, no line, no chain entry left', async () => {
    const { events, lines, tails, earlyAck } = setup()
    await earlyAck({ activity: msg('c', 'A') })
    expect(events).toEqual(['A:start', 'A:end'])
    await sleep(60)
    expect(lines).toEqual([])
    expect(tails.size).toBe(0)
  })

  test('b) slow turn: answered early, the turn goes on', async () => {
    const { events, lines, earlyAck } = setup({ A: 120 })
    await earlyAck({ activity: msg('c', 'A') })
    expect(events).toEqual(['A:start'])
    expect(lines.length).toBe(1)
    expect(lines[0]).toContain('still running after 30 ms')
    await waitFor(() => events.includes('A:end'))
    expect(lines.length).toBe(1)
  })

  test('c) slow failing turn: the failure is logged, not thrown', async () => {
    const { events, lines, earlyAck } = setup({ A: 120 }, new Set(['A']))
    let n = 0
    const h = () => { n++ }
    process.on('unhandledRejection', h)
    try {
      await earlyAck({ activity: msg('c', 'A') })
      await waitFor(() => lines.length === 2)
      await sleep(10)
    } finally {
      process.off('unhandledRejection', h)
    }
    expect(events).toEqual(['A:start', 'A:end'])
    expect(lines[1]).toBe('teams channel: turn failed after the request was answered: boom\n')
    expect(n).toBe(0)
  })

  test('c2) an unprintable failure value or a throwing logger does not become an unhandled rejection', async () => {
    let n = 0
    const h = () => { n++ }
    process.on('unhandledRejection', h)
    try {
      // An unprintable value (String() throws on it) with a normal logger, then a printable one with a logger that throws.
      for (const [value, throwingLog] of [[Object.create(null), false], [new Error('turn broke'), true]] as const) {
        const lines: string[] = []
        const turn = async () => { await sleep(60); throw value }
        const log = (l: string) => {
          lines.push(l)
          if (throwingLog && l.includes('failed')) throw new Error('log broke')
        }
        const earlyAck = makeEarlyAckTurn(turn as any, { ackAfterMs: 30, log })
        await earlyAck({ activity: msg('c', 'A') })
        await sleep(100)
        expect(lines[0]).toContain('still running')
        // The throwing logger was really reached with the failure line.
        if (throwingLog) expect(lines[1]).toBe('teams channel: turn failed after the request was answered: turn broke\n')
      }
      await sleep(10)
    } finally {
      process.off('unhandledRejection', h)
    }
    expect(n).toBe(0)
  })

  test('d) fast failing turn: the error reaches the caller', async () => {
    const { lines, tails, earlyAck } = setup({}, new Set(['A']))
    await expect(earlyAck({ activity: msg('c', 'A') })).rejects.toThrow('boom')
    expect(lines).toEqual([])
    await waitFor(() => tails.size === 0)
  })

  test('e) one conversation runs in arrival order; the chain entry is removed', async () => {
    const { events, tails, earlyAck } = setup({ A: 120 })
    void earlyAck({ activity: msg('c', 'A') })
    void earlyAck({ activity: msg('c', 'B') })
    await waitFor(() => events.length === 4)
    expect(events).toEqual(['A:start', 'A:end', 'B:start', 'B:end'])
    await waitFor(() => tails.size === 0)
  })

  test('e2) a message that comes after the first ended still waits for the second', async () => {
    const { events, earlyAck } = setup({ A: 40, B: 160 })
    void earlyAck({ activity: msg('c', 'A') })
    void earlyAck({ activity: msg('c', 'B') })
    await waitFor(() => events.includes('A:end'))
    await sleep(10)
    void earlyAck({ activity: msg('c', 'C') })
    await waitFor(() => events.includes('C:end'))
    expect(events).toEqual(['A:start', 'A:end', 'B:start', 'B:end', 'C:start', 'C:end'])
  })

  test('e3) a failed turn does not hold up the next message of the conversation', async () => {
    const { events, lines, tails, earlyAck } = setup({ A: 60 }, new Set(['A']))
    void earlyAck({ activity: msg('c', 'A') })
    await earlyAck({ activity: msg('c', 'B') })
    await waitFor(() => events.includes('B:end'))
    expect(events).toEqual(['A:start', 'A:end', 'B:start', 'B:end'])
    await waitFor(() => tails.size === 0)
    expect(lines.some(l => l.includes('turn failed after the request was answered: boom'))).toBe(true)
  })

  test('e4) a message without a conversation id passes straight through', async () => {
    const { events, tails, earlyAck } = setup({ A: 120, B: 60 })
    void earlyAck({ activity: { type: 'message', id: 'A', conversation: { id: '' } } as any })
    const p = earlyAck({ activity: { type: 'message', id: 'B', conversation: { id: '' } } as any })
    // Not queued behind A, and no chain entry under ''.
    expect(events).toContain('B:start')
    expect(tails.size).toBe(0)
    await p
    await waitFor(() => events.includes('A:end'))
  })

  test('f) other conversations do not wait', async () => {
    const { events, earlyAck } = setup({ A: 120 })
    void earlyAck({ activity: msg('c1', 'A') })
    await earlyAck({ activity: msg('c2', 'B') })
    expect(events).toEqual(['A:start', 'B:start', 'B:end'])
    await waitFor(() => events.includes('A:end'))
  })

  test('g) non-message activities pass straight through', async () => {
    const { events, lines, earlyAck } = setup({ A: 120, U: 120 })
    void earlyAck({ activity: msg('c', 'A') })
    const upd = { type: 'messageUpdate', id: 'U', conversation: { id: 'c' } } as any
    const p = earlyAck({ activity: upd })
    // Not queued behind the running message A of the same conversation.
    expect(events).toContain('U:start')
    await p
    // No early answer for the update: its call ends only after its turn.
    expect(events).toContain('U:end')
    // The only "still running" line is message A's.
    expect(lines.filter(l => l.includes('still running')).length).toBe(1)
    await waitFor(() => events.includes('A:end'))
  })

  test('h) a turn that starts after the request was answered does not touch the revoked context', async () => {
    const events: string[] = []
    let got: unknown
    const turn = async ({ activity }: { activity: any }) => {
      if (activity.id === 'B') got = activity
      events.push(`${activity.id}:start`)
      if (activity.id === 'A') await sleep(120)
      events.push(`${activity.id}:end`)
    }
    const earlyAck = makeEarlyAckTurn(turn as any, { ackAfterMs: 30, log: () => {} })
    const actB = msg('c', 'B')
    let revoked = false
    const ctxB = { get activity() { if (revoked) throw new Error('revoked'); return actB } }
    void earlyAck({ activity: msg('c', 'A') })
    // B waits behind A, so its request is answered early, before B's turn starts.
    await earlyAck(ctxB as any)
    expect(events).not.toContain('B:start')
    revoked = true
    await waitFor(() => events.includes('B:end'))
    expect(got).toBe(actB)
  })

  test('the request is answered after 5 seconds', () => {
    expect(ACK_AFTER_MS).toBe(5000)
  })

  test('adapter wires the early ack', () => {
    const src = readFileSync(new URL('../src/teams/adapter.ts', import.meta.url), 'utf8')
    expect(src).toContain('const turn = makeEarlyAckTurn(makeTurnHandler(deps))')
  })
})
