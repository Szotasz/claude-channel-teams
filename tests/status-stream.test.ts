import { describe, expect, test } from 'bun:test'
import {
  STATUS_TOOL,
  createStatusStreams,
  isStreamClosedError,
  statusTools,
} from '../src/teams/statusStream'

function setup() {
  let t = 0
  const sleeps: number[] = []
  const streams = createStatusStreams({
    now: () => t,
    sleep: async ms => {
      sleeps.push(ms)
      t += ms
    },
  })
  return { streams, sleeps, advance: (ms: number) => void (t += ms) }
}

const info = (extra: Record<string, unknown>) => ({ type: 'streaminfo', streamType: 'informative', ...extra })

describe('status streams', () => {
  test('first update, then a second right away waits 1000 ms', async () => {
    const { streams, sleeps } = setup()
    const sent: any[] = []
    await streams.update('c', 'Searching', async a => (sent.push(a), 's-1'))
    expect(sent[0]).toEqual({ type: 'typing', text: 'Searching', entities: [info({ streamSequence: 1 })] })
    expect(streams.size()).toBe(1)
    await streams.update('c', 'More', async a => (sent.push(a), undefined))
    expect(sleeps).toEqual([1000])
    expect(sent[1].entities).toEqual([info({ streamSequence: 2, streamId: 's-1' })])
  })

  test('an update 1500 ms later does not sleep', async () => {
    const { streams, sleeps, advance } = setup()
    await streams.update('c', 'a', async () => 's-1')
    advance(1500)
    await streams.update('c', 'b', async () => undefined)
    expect(sleeps).toEqual([])
  })

  test('text errors', async () => {
    const { streams } = setup()
    for (const bad of ['', '   ', 'x'.repeat(1001), 5 as any]) {
      await expect(streams.update('c', bad, async () => 's')).rejects.toThrow('status text must be 1-1000 characters')
    }
    await streams.update('c', 'x'.repeat(1000), async () => 's')
  })

  test('no stream id from the first send', async () => {
    for (const r of [undefined, '']) {
      const { streams } = setup()
      await expect(streams.update('c', 'a', async () => r)).rejects.toThrow('Teams returned no stream id')
      expect(streams.size()).toBe(0)
    }
  })

  test('2 minute limit: refused, and the expired stream stays until the reply', async () => {
    const { streams, advance, sleeps } = setup()
    await streams.update('c', 'a', async () => 's-1')
    advance(119_999)
    await streams.update('c', 'b', async () => undefined)
    advance(1)
    sleeps.length = 0
    let called = false
    const over = 'status updates for this reply are over (2 minutes); send the reply'
    await expect(streams.update('c', 'c', async () => ((called = true), undefined))).rejects.toThrow(over)
    // A later update does not open a second stream for the same reply.
    await expect(streams.update('c', 'd', async () => ((called = true), 's-2'))).rejects.toThrow(over)
    expect(called).toBe(false)
    // An expired stream is refused at once, without waiting for a slot first.
    expect(sleeps).toEqual([])
    expect(streams.size()).toBe(1)
    expect(await streams.withStream('c', async s => s)).toBeUndefined()
    expect(streams.size()).toBe(0)
  })

  test('the 2 minutes are checked again after the one-second wait', async () => {
    const { streams, advance } = setup()
    await streams.update('c', 'a', async () => 's-1')
    advance(119_500)
    await streams.update('c', 'b', async () => undefined) // lastSentAt 119500
    advance(100) // 119600: the next update waits 900 ms, to 120500
    let called = false
    await expect(streams.update('c', 'c', async () => ((called = true), undefined))).rejects.toThrow(
      'status updates for this reply are over (2 minutes); send the reply',
    )
    expect(called).toBe(false)
  })

  test('send errors: a closed stream refuses later updates until the reply; others keep it', async () => {
    for (const err of [{ code: 'ContentStreamNotAllowed' }, { body: { error: { code: 'ContentStreamNotAllowed' } } }]) {
      const { streams } = setup()
      await streams.update('c', 'a', async () => 's-1')
      await expect(streams.update('c', 'b', async () => { throw err })).rejects.toBe(err)
      let called = false
      await expect(streams.update('c', 'c', async () => ((called = true), 's-2'))).rejects.toThrow(
        'Teams closed the status line for this reply; send the reply',
      )
      expect(called).toBe(false)
      expect(await streams.withStream('c', async s => s)).toBeUndefined()
      expect(streams.size()).toBe(0)
    }
    const { streams, sleeps } = setup()
    await streams.update('c', 'a', async () => 's-1')
    const e = { statusCode: 500 }
    await expect(streams.update('c', 'b', async () => { throw e })).rejects.toBe(e)
    expect(streams.size()).toBe(1)
    // The failed request counts for the one-second rule too.
    sleeps.length = 0
    await streams.update('c', 'c', async () => undefined)
    expect(sleeps).toEqual([1000])
  })

  test('updates for one conversation run in order', async () => {
    const { streams } = setup()
    const log: string[] = []
    let release!: () => void
    const gate = new Promise<void>(r => (release = r))
    const p1 = streams.update('c', 'a', async a => {
      log.push('start1')
      await gate
      log.push('end1')
      return 's-1'
    })
    const p2 = streams.update('c', 'b', async (a: any) => {
      log.push('seq' + a.entities[0].streamSequence)
      return undefined
    })
    // Let every pending microtask run: the first send has started, the second still waits for it.
    await new Promise(r => setTimeout(r, 0))
    expect(log).toEqual(['start1'])
    release()
    await Promise.all([p1, p2])
    expect(log).toEqual(['start1', 'end1', 'seq2'])
  })

  test('a failed first update does not block the second, but the second waits for the slot', async () => {
    const { streams, sleeps } = setup()
    const p1 = streams.update('c', 'a', async () => { throw new Error('boom') })
    const p2 = streams.update('c', 'b', async (a: any) => {
      expect(a.entities[0].streamSequence).toBe(1)
      return 's-2'
    })
    await expect(p1).rejects.toThrow('boom')
    await p2
    expect(sleeps).toEqual([1000])
    expect(await streams.withStream('c', async s => s)).toEqual({ streamId: 's-2' })
  })

  test('an update right after the closing reply waits for the slot', async () => {
    const { streams, sleeps } = setup()
    await streams.update('c', 'a', async () => 's-1')
    await streams.withStream('c', async s => s)
    sleeps.length = 0
    await streams.update('c', 'b', async () => 's-2')
    expect(sleeps).toEqual([1000])
  })

  test('withStream waits for an update still in flight', async () => {
    const { streams } = setup()
    let release!: () => void
    const gate = new Promise<void>(r => (release = r))
    const p1 = streams.update('c', 'a', async () => {
      await gate
      return 's-1'
    })
    const closed = streams.withStream('c', async s => s)
    release()
    await p1
    expect(await closed).toEqual({ streamId: 's-1' })
    expect(streams.size()).toBe(0)
  })

  test('withStream keeps the one-second rule before the final message', async () => {
    const { streams, sleeps, advance } = setup()
    await streams.update('c', 'a', async () => 's-1')
    advance(10)
    sleeps.length = 0
    expect(await streams.withStream('c', async s => s)).toEqual({ streamId: 's-1' })
    expect(sleeps).toEqual([990])
  })

  test('withStream checks the 2 minutes again after the wait', async () => {
    const { streams, advance } = setup()
    await streams.update('c', 'a', async () => 's-1')
    advance(119_500)
    await streams.update('c', 'b', async () => undefined)
    advance(300) // 119800: waits 700 ms, to 120500
    expect(await streams.withStream('c', async s => s)).toBeUndefined()
    expect(streams.size()).toBe(0)
  })

  test('no update starts while a reply closes the stream', async () => {
    const { streams } = setup()
    await streams.update('c', 'a', async () => 's-1')
    const log: string[] = []
    let release!: () => void
    const gate = new Promise<void>(r => (release = r))
    const reply = streams.withStream('c', async s => {
      log.push('final ' + s?.streamId)
      await gate
      log.push('final done')
    })
    const next = streams.update('c', 'b', async (a: any) => {
      log.push('update seq' + a.entities[0].streamSequence + ' ' + (a.entities[0].streamId ?? 'new'))
      return 's-2'
    })
    await new Promise(r => setTimeout(r, 0))
    expect(log).toEqual(['final s-1'])
    release()
    await Promise.all([reply, next])
    expect(log).toEqual(['final s-1', 'final done', 'update seq1 new'])
  })

  test('the stream is gone after a reply, also when the reply fails', async () => {
    const { streams } = setup()
    await streams.update('c', 'a', async () => 's-1')
    await expect(streams.withStream('c', async () => { throw new Error('send failed') })).rejects.toThrow('send failed')
    expect(streams.size()).toBe(0)
    expect(await streams.withStream('c', async s => s)).toBeUndefined()
  })

  test('the 2 minutes count from the request that opened the stream', async () => {
    const { streams, advance } = setup()
    await streams.update('c', 'a', async () => {
      advance(5000)
      return 's-1'
    })
    advance(115_000)
    expect(await streams.withStream('c', async s => s)).toBeUndefined()
  })
})

test('isStreamClosedError', () => {
  expect(isStreamClosedError({ code: 'ContentStreamNotAllowed' })).toBe(true)
  expect(isStreamClosedError({ body: { error: { code: 'ContentStreamNotAllowed' } } })).toBe(true)
  for (const v of [{ statusCode: 403 }, {}, null, 'x']) expect(isStreamClosedError(v)).toBe(false)
})

test('statusTools', () => {
  expect(statusTools(true)).toEqual([STATUS_TOOL])
  expect(statusTools(false)).toEqual([])
})
