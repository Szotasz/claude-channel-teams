// A message turn still running after ACK_AFTER_MS gets its request answered and goes on in the
// background (so Teams does not retry it); messages of one conversation run one after another in
// arrival order; the turn gets only the activity, because it may run after botbuilder revoked
// the context.
import type { TurnContext } from 'botbuilder'

export const ACK_AFTER_MS = 5_000
export type ActivityTurn = (ctx: Pick<TurnContext, 'activity'>) => Promise<void>
export interface EarlyAckOptions {
  ackAfterMs?: number
  log?: (line: string) => void
  /** The per-conversation chain; injectable for tests. */
  tails?: Map<string, Promise<void>>
}
export function makeEarlyAckTurn(turn: ActivityTurn, opts: EarlyAckOptions = {}): ActivityTurn {
  const ackAfterMs = opts.ackAfterMs ?? ACK_AFTER_MS
  const log = opts.log ?? ((line: string) => { process.stderr.write(line) })
  const tails = opts.tails ?? new Map<string, Promise<void>>()
  return async function earlyAckTurn(ctx) {
    const activity = ctx.activity
    const conversation = activity.conversation?.id
    if (activity.type !== 'message' || typeof conversation !== 'string' || conversation.length === 0) {
      await turn(ctx)
      return
    }
    const own = { activity }
    const previous = tails.get(conversation) ?? Promise.resolve()
    const run = previous.then(() => turn(own))
    const tail = run.then(() => {}, () => {})
    tails.set(conversation, tail)
    void tail.then(() => {
      if (tails.get(conversation) === tail) tails.delete(conversation)
    })
    let timer: ReturnType<typeof setTimeout> | undefined
    const late = new Promise<'late'>(resolve => {
      timer = setTimeout(() => resolve('late'), ackAfterMs)
    })
    try {
      const outcome = await Promise.race([run.then(() => 'done' as const), late])
      if (outcome === 'late') {
        log(`teams channel: turn still running after ${ackAfterMs} ms, answering the request now conv=${conversation.slice(0, 20)}\n`)
        run.catch(err => {
          try {
            log(`teams channel: turn failed after the request was answered: ${err instanceof Error ? err.message : String(err)}\n`)
          } catch {
            // A failure value that cannot be printed, or a logger that throws, must not become an unhandled rejection.
          }
        })
      }
    } finally {
      clearTimeout(timer)
    }
  }
}
