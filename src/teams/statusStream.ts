/**
 * Status updates for 1:1 chats as Teams streaming "informative updates".
 *
 * Kept in memory per conversation: the first update opens a stream (Teams
 * returns its id), later updates carry that id and a rising sequence number.
 * The next reply closes the stream (see reply.ts, `take`).
 *
 * Limits (Teams): 1:1 chats only, at most one request per second (update waits),
 * a stream lasts at most 2 minutes, informative text at most 1000 characters.
 */

export const MAX_STATUS_CHARS = 1000
export const MIN_INTERVAL_MS = 1000
export const MAX_STREAM_MS = 120_000

/** Teams answers 403 ContentStreamNotAllowed when a stream is over (timed out, stopped, closed). */
export function isStreamClosedError(err: unknown): boolean {
  if (typeof err !== 'object' || err === null) return false
  const e = err as { code?: unknown; body?: { error?: { code?: unknown } } }
  return e.code === 'ContentStreamNotAllowed' || e.body?.error?.code === 'ContentStreamNotAllowed'
}

export type StatusSend = (activity: Record<string, unknown>) => Promise<string | undefined>

/** closed: Teams ended the stream (403); kept until the next reply, so no new stream opens for this reply. */
type Stream = { streamId: string; sequence: number; startedAt: number; closed: boolean }

export function createStatusStreams(
  opts: { now?: () => number; sleep?: (ms: number) => Promise<void> } = {},
) {
  const now = opts.now ?? Date.now
  const sleep = opts.sleep ?? ((ms: number) => new Promise<void>(r => setTimeout(r, ms)))
  const streams = new Map<string, Stream>()
  const chains = new Map<string, Promise<unknown>>()
  /** Last streaming request per conversation, failed ones included: the one-per-second rule is per chat. */
  const lastRequestAt = new Map<string, number>()

  const expired = (st: Stream) => now() - st.startedAt >= MAX_STREAM_MS
  async function waitForSlot(conversationId: string): Promise<void> {
    const last = lastRequestAt.get(conversationId)
    if (last === undefined) return
    const wait = last + MIN_INTERVAL_MS - now()
    if (wait > 0) await sleep(wait)
  }

  /** Runs jobs for one conversation one after another; a failed job does not block the next. */
  function enqueue<T>(conversationId: string, job: () => Promise<T>): Promise<T> {
    const prev = chains.get(conversationId) ?? Promise.resolve()
    const next = prev.then(job)
    const tail = next.then(
      () => {},
      () => {},
    )
    chains.set(conversationId, tail)
    void tail.then(() => {
      if (chains.get(conversationId) === tail) chains.delete(conversationId)
    })
    return next
  }

  async function run(conversationId: string, text: string, send: StatusSend): Promise<void> {
    const st = streams.get(conversationId)
    // An expired or closed stream stays until the next reply: no second stream opens for the same reply.
    const over = () => new Error('status updates for this reply are over (2 minutes); send the reply')
    if (st?.closed) throw new Error('Teams closed the status line for this reply; send the reply')
    if (st && expired(st)) throw over()
    await waitForSlot(conversationId)
    if (st && expired(st)) throw over()
    const seq = st ? st.sequence + 1 : 1
    const activity = {
      type: 'typing',
      text,
      entities: [
        {
          type: 'streaminfo',
          streamType: 'informative',
          streamSequence: seq,
          ...(st ? { streamId: st.streamId } : {}),
        },
      ],
    }
    let id: string | undefined
    // Teams counts the 2 minutes from the request that opens the stream, so take the time before sending.
    const sentAt = now()
    try {
      id = await send(activity)
    } catch (err) {
      if (st && isStreamClosedError(err)) st.closed = true
      throw err
    } finally {
      lastRequestAt.set(conversationId, now())
    }
    if (!st) {
      if (typeof id !== 'string' || id.length === 0) throw new Error('Teams returned no stream id')
      streams.set(conversationId, { streamId: id, sequence: 1, startedAt: sentAt, closed: false })
    } else {
      st.sequence = seq
    }
  }

  return {
    update(conversationId: string, text: string, send: StatusSend): Promise<void> {
      if (typeof text !== 'string' || text.trim().length === 0 || text.length > MAX_STATUS_CHARS) {
        return Promise.reject(new Error(`status text must be 1-${MAX_STATUS_CHARS} characters`))
      }
      return enqueue(conversationId, () => run(conversationId, text, send))
    },
    /**
     * Runs a reply for this conversation in the same queue as the updates: it waits for an update in flight and for
     * the one-second slot, and no update starts while it runs. `fn` gets the stream to close, or undefined when
     * there is none (or it is over or closed). The stream is gone afterwards, whatever `fn` does.
     */
    withStream<T>(conversationId: string, fn: (stream: { streamId: string } | undefined) => Promise<T>): Promise<T> {
      return enqueue(conversationId, async () => {
        const st = streams.get(conversationId)
        try {
          if (!st || st.closed || expired(st)) return await fn(undefined)
          await waitForSlot(conversationId)
          if (expired(st)) return await fn(undefined)
          try {
            return await fn({ streamId: st.streamId })
          } finally {
            // The final message is a streaming request too.
            lastRequestAt.set(conversationId, now())
          }
        } finally {
          if (streams.get(conversationId) === st) streams.delete(conversationId)
        }
      })
    },
    size(): number {
      return streams.size
    },
  }
}

export type StatusStreams = ReturnType<typeof createStatusStreams>

export const STATUS_TOOL = {
  name: 'status_update',
  description:
    "1:1 chats only. Shows a short line about what you are doing (for example 'Searching the tender list...') in Teams' progress indicator where your reply will appear; your next reply replaces it. Plain text, up to 1000 characters. At most one update per second (the call waits). After 2 minutes no more updates are shown; then just send the reply.",
  inputSchema: {
    type: 'object',
    properties: {
      conversation_id: { type: 'string', description: 'The conversation_id of the 1:1 chat.' },
      text: { type: 'string', maxLength: 1000, description: 'What you are doing now.' },
    },
    required: ['conversation_id', 'text'],
  },
}

export function statusTools(enabled: boolean): (typeof STATUS_TOOL)[] {
  return enabled ? [STATUS_TOOL] : []
}
