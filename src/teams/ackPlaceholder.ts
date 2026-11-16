/**
 * Acknowledgement placeholder: one acknowledgement per conversation at a time.
 * It is sent after a delay unless a reply came first; the next reply
 * overwrites it in place (see reply.ts). An acknowledgement older than
 * ACK_MAX_AGE_MS is not overwritten.
 */

export const ACK_MAX_AGE_MS = 10 * 60 * 1000

export interface AckPlaceholders {
  /** A message was delivered to Claude: send the acknowledgement after the delay, unless one is pending or a reply comes first. */
  arm(conversationId: string): void
  /** A reply is about to be sent: cancels a pending timer, waits for a send in flight, and returns the id of an
   *  acknowledgement younger than ACK_MAX_AGE_MS (it is then consumed); otherwise undefined. */
  take(conversationId: string): Promise<string | undefined>
  /** Another kind of send (a file): cancels a pending timer; an acknowledgement already sent is left alone. */
  cancel(conversationId: string): void
  /** Clears every pending timer and arms no new one (shutdown). */
  stopAll(): void
}

export interface AckDeps {
  delayMs: number
  /** Sends the acknowledgement and returns its message id (undefined when Teams gave none). */
  send: (conversationId: string) => Promise<string | undefined>
  now?: () => number
  setTimer?: (fn: () => void, ms: number) => unknown
  clearTimer?: (handle: unknown) => void
  log?: (line: string) => void
}

type State =
  | { kind: 'timer'; handle: unknown }
  | { kind: 'sending'; done: Promise<void> }
  | { kind: 'sent'; id: string; at: number }

export function createAckPlaceholders(deps: AckDeps): AckPlaceholders {
  const now = deps.now ?? Date.now
  const setTimer = deps.setTimer ?? ((fn: () => void, ms: number) => setTimeout(fn, ms))
  const clearTimer = deps.clearTimer ?? ((h: unknown) => clearTimeout(h as ReturnType<typeof setTimeout>))
  const log = deps.log ?? ((line: string) => process.stderr.write(line))
  const states = new Map<string, State>()
  let stopped = false

  function fire(c: string): void {
    const sending = { kind: 'sending' as const, done: Promise.resolve() }
    states.set(c, sending)
    sending.done = deps.send(c).then(
      id => {
        if (states.get(c) === sending) {
          if (id) states.set(c, { kind: 'sent', id, at: now() })
          else states.delete(c)
        }
      },
      err => {
        if (states.get(c) === sending) states.delete(c)
        try {
          log(`teams channel: acknowledgement failed conv=${c.slice(0, 20)}: ${err instanceof Error ? err.message : String(err)}\n`)
        } catch {}
      },
    )
  }

  return {
    arm(c) {
      if (stopped) return
      const s = states.get(c)
      if (s?.kind === 'timer' || s?.kind === 'sending') return
      if (s?.kind === 'sent' && now() - s.at < ACK_MAX_AGE_MS) return
      states.set(c, { kind: 'timer', handle: setTimer(() => fire(c), deps.delayMs) })
    },
    async take(c) {
      let s = states.get(c)
      if (s?.kind === 'timer') {
        clearTimer(s.handle)
        states.delete(c)
        return undefined
      }
      if (s?.kind === 'sending') {
        await s.done
        s = states.get(c)
      }
      if (s?.kind === 'sent') {
        states.delete(c)
        return now() - s.at < ACK_MAX_AGE_MS ? s.id : undefined
      }
      return undefined
    },
    cancel(c) {
      const s = states.get(c)
      if (s?.kind === 'timer') {
        clearTimer(s.handle)
        states.delete(c)
      }
    },
    stopAll() {
      stopped = true
      for (const [c, s] of states) {
        if (s.kind === 'timer') {
          clearTimer(s.handle)
          states.delete(c)
        }
      }
    },
  }
}
