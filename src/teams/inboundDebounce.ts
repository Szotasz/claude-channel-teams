/**
 * Inbound debounce — opt-in coalescing of quick successive messages from one
 * sender in one conversation.
 *
 * People on a phone often split one request into two or three quick messages
 * ("look at this" + a link + "urgent"). Delivered straight away, the first
 * fragment alone starts a full agent turn on half a request. Claude Code
 * already batches messages that arrive WHILE the agent is busy; this closes
 * the gap for the opening fragments.
 *
 * Each event is held for `windowMs`, and every new fragment from the same
 * sender re-arms that window, so a burst lands as one event. A message that
 * must not wait — an edit / delete / restore, a message with attachments, a
 * slash command, a card button click, a message action — flushes what is pending for that sender first and then goes
 * through unchanged, so per-sender order is preserved. The merged event is a
 * shallow copy of the LAST message (that is the one to answer) with the texts
 * joined by '\n' and `merged = { count, messageIds }`. `mentioned` is true when
 * any fragment mentioned the bot, and `mentions` is the union of the fragments'
 * names, bounded like the adapter's list.
 *
 * Off unless configured: windowMs <= 0 delivers every event immediately and
 * unchanged. See TEAMS_INBOUND_DEBOUNCE_MS in config.ts and the design notes.
 */

import { MAX_MENTION_NAMES, MAX_MENTIONS_CHARS } from './adapter.js'
import type { ChannelEventSink } from './adapter.js'
import { stripInvisible } from './invisibleChars.js'

type Ev = Parameters<ChannelEventSink>[0]

/** The buffer key: one sender in one conversation. '\u0000' is in neither id. */
function bufferKey(event: Ev): string {
  return `${event.conversationId}\u0000${event.aadObjectId.toLowerCase()}`
}

/** True when the event must not be held back (see the module comment). */
function isImmediate(event: Ev): boolean {
  return (
    event.messageEvent !== undefined ||
    event.cardAction !== undefined ||
    // A merged copy is the last fragment's: a message action held in a burst would lose its message_action tag.
    event.messageAction !== undefined ||
    event.hasAttachments === true ||
    // Read the command the way the model will see it: onEvent strips invisible characters after this, so a
    // zero-width character in front of the slash must not turn a command into a held fragment.
    stripInvisible(event.text).text.trimStart().startsWith('/')
  )
}

/** The union of the fragments' mention names in arrival order, bounded like the adapter's own list. */
function mergedMentions(events: Ev[]): string[] | undefined {
  const names: string[] = []
  let len = 0
  for (const event of events) {
    for (const name of event.mentions ?? []) {
      if (names.includes(name)) continue
      const next = names.length === 0 ? name.length : len + 2 + name.length
      if (next > MAX_MENTIONS_CHARS) return names.length > 0 ? names : undefined
      names.push(name)
      len = next
      if (names.length >= MAX_MENTION_NAMES) return names
    }
  }
  return names.length > 0 ? names : undefined
}

interface KeyBuffer {
  /** Arrival order; the last one is the message to answer. */
  events: Ev[]
  /** now() when the first event of this window was buffered. */
  firstAt: number
  timer: unknown
}

export function createInboundDebouncer(opts: {
  windowMs: number
  /** Default 3 * windowMs — the cap on how long one merged message can be held. */
  maxWaitMs?: number
  /** Default 20 — bound on merged.messageIds. */
  maxMergedIds?: number
  deliver: (event: Ev) => void
  now?: () => number
  setTimer?: (fn: () => void, ms: number) => unknown
  clearTimer?: (handle: unknown) => void
}): { push(event: Ev): void; flushAll(): void; stop(): void; pending(): number } {
  const deliver = opts.deliver
  const maxMergedIds = opts.maxMergedIds ?? 20
  const setTimer = opts.setTimer ?? ((fn: () => void, ms: number) => setTimeout(fn, ms))
  const clearTimer = opts.clearTimer ?? ((handle: unknown) => clearTimeout(handle as ReturnType<typeof setTimeout>))

  if (opts.windowMs <= 0) {
    return {
      push(event) {
        deliver(event)
      },
      flushAll() {},
      stop() {},
      pending() {
        return 0
      },
    }
  }

  const windowMs = opts.windowMs
  const maxWaitMs = opts.maxWaitMs ?? 3 * windowMs
  // A monotonic clock: a wall clock set back (NTP step, manual change) must not stretch the cap.
  const now = opts.now ?? (() => performance.now())
  // Insertion order is the order the keys were first buffered; flushAll walks it.
  const buffers = new Map<string, KeyBuffer>()

  /** Deliver one key's buffer as a single event (or the lone event unchanged) and drop it. */
  function flushKey(key: string): void {
    const buffered = buffers.get(key)
    if (!buffered) return
    // Take the buffer out and drop its timer before delivering, so neither a
    // late timer tick nor a re-entrant push can deliver an event twice.
    buffers.delete(key)
    clearTimer(buffered.timer)
    const events = buffered.events
    if (events.length === 1) {
      deliver(events[0]!)
      return
    }
    const messageIds: string[] = []
    for (const event of events) {
      if (event.messageId !== undefined && messageIds.length < maxMergedIds) messageIds.push(event.messageId)
    }
    const last = events[events.length - 1]!
    const mentions = mergedMentions(events)
    deliver({
      ...last,
      text: events.map(event => event.text).join('\n'),
      ...(events.some(event => event.mentioned === true) ? { mentioned: true } : {}),
      ...(mentions !== undefined || last.mentions !== undefined ? { mentions } : {}),
      merged: { count: events.length, messageIds },
    })
  }

  // After stop() (shutdown) nothing is held any more: a message that was still in flight in the listener goes straight
  // through instead of waiting for a timer that will not fire before the process exits.
  let stopped = false

  // A flush that throws (the notification could not be built) must neither escape a timer callback nor stop the
  // flush of the other senders.
  function safeFlush(key: string): void {
    try {
      flushKey(key)
    } catch (err) {
      process.stderr.write(`teams channel: inbound debounce flush failed: ${err instanceof Error ? err.message : String(err)}\n`)
    }
  }

  return {
    push(event) {
      const key = bufferKey(event)
      if (stopped || isImmediate(event)) {
        // What is already pending for this sender goes first, the urgent one
        // right behind it. A pending burst that fails to go out must not take the urgent one with it.
        safeFlush(key)
        deliver(event)
        return
      }
      let buffered = buffers.get(key)
      if (!buffered) {
        buffered = { events: [], firstAt: now(), timer: undefined }
        buffers.set(key, buffered)
      }
      buffered.events.push(event)
      clearTimer(buffered.timer)
      // The window resets with every fragment, but a merged message never waits
      // longer than maxWaitMs counted from the first one.
      const remaining = buffered.firstAt + maxWaitMs - now()
      if (remaining <= 0) {
        // Past the cap already (a late event loop): hand the burst over now instead of queueing another timer.
        safeFlush(key)
        return
      }
      buffered.timer = setTimer(() => safeFlush(key), Math.min(windowMs, remaining))
    },
    flushAll() {
      for (const key of [...buffers.keys()]) safeFlush(key)
    },
    stop() {
      stopped = true
      for (const key of [...buffers.keys()]) safeFlush(key)
    },
    pending() {
      let count = 0
      for (const buffered of buffers.values()) count += buffered.events.length
      return count
    },
  }
}
