/**
 * Inbound debounce — coalescing of quick successive messages from one sender
 * in one conversation (src/teams/inboundDebounce.ts).
 *
 * All timing goes through the injectable now / setTimer / clearTimer with a
 * hand-rolled fake clock, so no test waits on a real timer: advance(ms) runs
 * the timers that came due, in time order, and nothing beyond.
 */

import { describe, expect, test } from 'bun:test'

import { createInboundDebouncer } from '../src/teams/inboundDebounce'
import type { ChannelEventSink } from '../src/teams/adapter'

type Ev = Parameters<ChannelEventSink>[0]

let seq = 0
/** A normal (bufferable) message event; every field can be overridden. */
function fragment(overrides: Partial<Ev> = {}): Ev {
  seq += 1
  return {
    text: `t${seq}`,
    conversationId: 'conv-1',
    aadObjectId: 'aad-1',
    tenantId: 'tenant-1',
    fromName: 'Sender',
    messageId: `m${seq}`,
    ts: `2026-10-24T10:00:${String(seq % 60).padStart(2, '0')}.000Z`,
    ...overrides,
  }
}

function fakeClock() {
  let t = 0
  let nextId = 0
  const timers = new Map<number, { at: number; fn: () => void; cancelled: boolean }>()
  return {
    now: () => t,
    setTimer(fn: () => void, ms: number) {
      nextId += 1
      timers.set(nextId, { at: t + ms, fn, cancelled: false })
      return nextId
    },
    clearTimer(handle: unknown) {
      const entry = timers.get(handle as number)
      if (entry) entry.cancelled = true
    },
    /** Timers still armed — 0 once everything pending was flushed and cleared. */
    live() {
      let n = 0
      for (const entry of timers.values()) if (!entry.cancelled) n += 1
      return n
    },
    advance(ms: number) {
      const target = t + ms
      for (;;) {
        let due: { id: number; at: number; fn: () => void } | undefined
        for (const [id, entry] of timers) {
          if (entry.cancelled) {
            timers.delete(id)
            continue
          }
          if (entry.at <= target && (!due || entry.at < due.at)) due = { id, at: entry.at, fn: entry.fn }
        }
        if (!due) break
        timers.delete(due.id)
        t = due.at
        due.fn()
      }
      t = target
    },
  }
}

function makeHarness(opts: { windowMs: number; maxWaitMs?: number; maxMergedIds?: number }) {
  const clock = fakeClock()
  const delivered: Ev[] = []
  const debouncer = createInboundDebouncer({
    ...opts,
    deliver: event => delivered.push(event),
    now: clock.now,
    setTimer: clock.setTimer,
    clearTimer: clock.clearTimer,
  })
  return { clock, delivered, debouncer }
}

describe('inbound debounce', () => {
  test('B1 windowMs <= 0 delivers immediately and unchanged, buffering nothing', () => {
    const { clock, delivered, debouncer } = makeHarness({ windowMs: 0 })
    const a = fragment({ text: 'a' })
    const b = fragment({ text: 'b' })
    debouncer.push(a)
    expect(delivered).toEqual([a])
    expect(delivered[0]).toBe(a)
    debouncer.push(b)
    expect(delivered[1]).toBe(b)
    expect(debouncer.pending()).toBe(0)
    clock.advance(60000)
    expect(delivered.length).toBe(2)
    debouncer.flushAll()
    expect(delivered.length).toBe(2)
  })

  test('B2 three fragments 100 ms apart arrive as one merged event after the window', () => {
    const { clock, delivered, debouncer } = makeHarness({ windowMs: 1500 })
    const a = fragment({ text: 'a', messageId: 'm1', ts: '2026-10-24T10:00:01.000Z' })
    const b = fragment({ text: 'b', messageId: 'm2', ts: '2026-10-24T10:00:02.000Z' })
    const c = fragment({ text: 'c', messageId: 'm3', ts: '2026-10-24T10:00:03.000Z' })
    debouncer.push(a)
    clock.advance(100)
    debouncer.push(b)
    clock.advance(100)
    debouncer.push(c)
    expect(delivered).toEqual([])
    expect(debouncer.pending()).toBe(3)
    clock.advance(1499)
    expect(delivered).toEqual([]) // 1 ms before the window after the last fragment ends
    clock.advance(1) // exactly 1500 ms after the third fragment
    expect(delivered.length).toBe(1)
    const merged = delivered[0]!
    // The last message is the one to answer: its id and timestamp survive.
    expect(merged.text).toBe('a\nb\nc')
    expect(merged.messageId).toBe('m3')
    expect(merged.ts).toBe('2026-10-24T10:00:03.000Z')
    expect(merged.merged).toEqual({ count: 3, messageIds: ['m1', 'm2', 'm3'] })
    expect(debouncer.pending()).toBe(0)
    expect(clock.live()).toBe(0)
  })

  test('B3 another sender in the same conversation, or the same sender in another conversation, never merge', () => {
    const { clock, delivered, debouncer } = makeHarness({ windowMs: 1500 })
    const sameConvOtherSender = fragment({ text: 'other sender', conversationId: 'conv-1', aadObjectId: 'aad-2' })
    const otherConvSameSender = fragment({ text: 'other conv', conversationId: 'conv-2', aadObjectId: 'aad-1' })
    const lone = fragment({ text: 'alone', conversationId: 'conv-1', aadObjectId: 'aad-1' })
    debouncer.push(lone)
    debouncer.push(sameConvOtherSender)
    debouncer.push(otherConvSameSender)
    expect(debouncer.pending()).toBe(3)
    clock.advance(1500)
    expect(delivered.length).toBe(3)
    // Each was alone in its buffer: delivered unchanged, as the same object, unmerged.
    expect(delivered).toEqual([lone, sameConvOtherSender, otherConvSameSender])
    for (const event of delivered) expect(event.merged).toBeUndefined()
    // The sender id is matched case-insensitively, so the same person in the
    // same conversation does merge — the id is an AAD Object ID, not a name.
    const upper = fragment({ text: 'upper', conversationId: 'conv-3', aadObjectId: 'AAD-1' })
    const lower = fragment({ text: 'lower', conversationId: 'conv-3', aadObjectId: 'aad-1' })
    debouncer.push(upper)
    clock.advance(100)
    debouncer.push(lower)
    clock.advance(1500)
    expect(delivered.length).toBe(4)
    expect(delivered[3]!.text).toBe('upper\nlower')
    expect(delivered[3]!.merged).toEqual({ count: 2, messageIds: [upper.messageId, lower.messageId] })
  })

  test('B4 a fragment every 1000 ms with a 1500 ms window is delivered at the 3x cap, not later', () => {
    const { clock, delivered, debouncer } = makeHarness({ windowMs: 1500 })
    for (let i = 0; i < 5; i++) {
      if (i > 0) clock.advance(1000)
      debouncer.push(fragment({ text: `f${i}`, messageId: `m${i}` }))
    }
    // t = 4000. Each fragment re-arms the window, which alone would wait until
    // 5500 — but maxWaitMs (3 * 1500) caps the hold at 4500 from the first.
    expect(delivered).toEqual([])
    clock.advance(499)
    expect(delivered).toEqual([]) // 4499
    clock.advance(1) // 4500 = first fragment + 3 * windowMs
    expect(delivered.length).toBe(1)
    expect(delivered[0]!.text).toBe('f0\nf1\nf2\nf3\nf4')
    expect(delivered[0]!.merged).toEqual({
      count: 5,
      messageIds: ['m0', 'm1', 'm2', 'm3', 'm4'],
    })
    clock.advance(60000)
    expect(delivered.length).toBe(1)
  })

  test('B5 edits, attachments and slash commands flush what is pending, then go through unchanged', () => {
    const immediates: { label: string; event: Ev }[] = [
      { label: 'edit', event: fragment({ text: '[Message edited] hi', messageEvent: 'edited' }) },
      { label: 'attachments', event: fragment({ text: 'see this', hasAttachments: true }) },
      { label: 'slash command', event: fragment({ text: '/status' }) },
      { label: 'indented slash command', event: fragment({ text: '   /status' }) },
    ]
    for (const { label, event } of immediates) {
      const { clock, delivered, debouncer } = makeHarness({ windowMs: 1500 })
      const pending = fragment({ text: 'first fragment', messageId: 'p1' })
      debouncer.push(pending)
      clock.advance(100)
      expect(delivered).toEqual([])
      debouncer.push(event)
      // The pending fragment is delivered first (alone, so unchanged), the
      // immediate one right after it, unchanged and in order.
      expect(delivered.length).toBe(2)
      expect(delivered[0]).toBe(pending)
      expect(delivered[1]).toBe(event)
      expect(delivered[1]!.merged).toBeUndefined()
      expect(debouncer.pending()).toBe(0)
      expect(clock.live()).toBe(0)
      clock.advance(60000) // the flushed timer never fires again
      expect(delivered.length).toBe(2)
      // With two fragments waiting, the flush delivers them merged — the
      // immediate event still goes through last and unchanged.
      const first = fragment({ text: 'look at this', messageId: 'q1' })
      const second = fragment({ text: 'a link', messageId: 'q2' })
      debouncer.push(first)
      clock.advance(50)
      debouncer.push(second)
      clock.advance(50)
      const before = delivered.length
      debouncer.push(label === 'edit' ? fragment({ text: '[Message deleted]', messageEvent: 'deleted' }) : event)
      expect(delivered.length).toBe(before + 2)
      expect(delivered[delivered.length - 2]!.text).toBe('look at this\na link')
      expect(delivered[delivered.length - 2]!.merged).toEqual({ count: 2, messageIds: ['q1', 'q2'] })
      expect(delivered[delivered.length - 1]!.merged).toBeUndefined()
      expect(debouncer.pending()).toBe(0)
    }
  })

  test('B6 one fragment alone is delivered unchanged when its window ends', () => {
    const { clock, delivered, debouncer } = makeHarness({ windowMs: 1500 })
    const lone = fragment({ text: 'just me' })
    debouncer.push(lone)
    clock.advance(1499)
    expect(delivered).toEqual([])
    clock.advance(1)
    expect(delivered).toEqual([lone])
    expect(delivered[0]).toBe(lone) // the same object, not a copy
    expect(delivered[0]!.merged).toBeUndefined()
    expect(debouncer.pending()).toBe(0)
  })

  test('B7 flushAll delivers every pending key in order, clears the timers, empties the buffers', () => {
    const { clock, delivered, debouncer } = makeHarness({ windowMs: 1500 })
    const a1 = fragment({ text: 'a1', conversationId: 'conv-1', aadObjectId: 'aad-1' })
    const b1 = fragment({ text: 'b1', conversationId: 'conv-2', aadObjectId: 'aad-1' })
    const a2 = fragment({ text: 'a2', conversationId: 'conv-1', aadObjectId: 'aad-1' })
    const c1 = fragment({ text: 'c1', conversationId: 'conv-3', aadObjectId: 'aad-2' })
    debouncer.push(a1)
    debouncer.push(b1)
    debouncer.push(a2)
    debouncer.push(c1)
    expect(debouncer.pending()).toBe(4)
    debouncer.flushAll()
    // Key order is the order the keys were first buffered: conv-1, conv-2, conv-3.
    expect(delivered.map(event => event.text)).toEqual(['a1\na2', 'b1', 'c1'])
    expect(delivered[1]).toBe(b1)
    expect(delivered[2]).toBe(c1)
    expect(delivered[0]!.merged).toEqual({ count: 2, messageIds: [a1.messageId, a2.messageId] })
    expect(debouncer.pending()).toBe(0)
    expect(clock.live()).toBe(0)
    clock.advance(60000) // a timer tick after flushAll delivers nothing more
    expect(delivered.length).toBe(3)
  })

  test('B8 at most 20 merged message ids (the first ones); fragments without an id count but add no id', () => {
    const { clock, delivered, debouncer } = makeHarness({ windowMs: 1500 })
    // 25 fragments; the 13th and the last carry no message id.
    for (let i = 0; i < 25; i++) {
      debouncer.push(fragment({ text: `f${i}`, messageId: i === 12 || i === 24 ? undefined : `m${i}` }))
    }
    expect(debouncer.pending()).toBe(25)
    clock.advance(1500)
    expect(delivered.length).toBe(1)
    const merged = delivered[0]!
    expect(merged.merged!.count).toBe(25) // the id-less fragments are counted
    expect(merged.merged!.messageIds).toEqual([
      'm0', 'm1', 'm2', 'm3', 'm4', 'm5', 'm6', 'm7', 'm8', 'm9',
      'm10', 'm11', 'm13', 'm14', 'm15', 'm16', 'm17', 'm18', 'm19', 'm20',
    ])
    expect(merged.merged!.messageIds.length).toBe(20)
    expect(merged.merged!.messageIds).not.toContain('m21') // has an id, but sits past the 20-id cap
  })
})
