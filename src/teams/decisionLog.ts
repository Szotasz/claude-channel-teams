/**
 * Durable, TEXT-FREE decision log (opt-in, TEAMS_DECISION_LOG).
 *
 * One JSON object per line in `<stateDir>/decisions.log` (mode 0600, append), so a lost or dropped message can be
 * traced afterwards without keeping message content on disk. A line holds ONLY: the timestamp, the decision, a short
 * fixed reason code, the source (live or backfill), and — when known — the message id and the channel part of the
 * conversation id. Never text, display names, AAD object ids or tokens: those identify people and this file is meant
 * to be kept around and read.
 *
 * The file is rotated by size: before an append that would grow it past `maxBytes`, the current file becomes
 * `decisions.log.1` (replacing an older one) and a fresh file starts. Anything that can go wrong writing it is one
 * stderr line per process and never an exception into the turn handler.
 */

import { appendFileSync, existsSync, mkdirSync, renameSync, statSync } from 'fs'
import { dirname } from 'path'

export type Decision =
  | 'deliver'
  | 'drop'
  | 'ignore'
  | 'held'
  | 'notify_ok'
  | 'notify_failed'
  | 'backfill_submit'
  | 'backfill_skip'
  | 'backfill_error'

export interface DecisionLog {
  record(e: {
    decision: Decision
    reason: string
    source: 'live' | 'backfill'
    messageId?: string
    conversationId?: string
  }): void
}

/** A reason is a short fixed code, nothing else (so no message text can leak into it). */
const REASON_RE = /^[a-z0-9_]{1,40}$/

/** The channel / chat part of a conversation id (drops ';messageid=...'). Same rule as the adapter's helper. */
function channelPart(conversationId: string): string {
  const i = conversationId.indexOf(';')
  return i === -1 ? conversationId : conversationId.slice(0, i)
}

export function createDecisionLog(opts: { file: string; maxBytes: number; now?: () => number }): DecisionLog {
  const now = opts.now ?? Date.now
  const rotated = `${opts.file}.1`
  let size: number | undefined
  let warned = false

  const warn = (err: unknown): void => {
    if (warned) return
    warned = true
    process.stderr.write(`teams channel: decision log write failed (${err instanceof Error ? err.message : String(err)})\n`)
  }

  return {
    record(e) {
      try {
        const line =
          JSON.stringify({
            ts: new Date(now()).toISOString(),
            decision: e.decision,
            reason: REASON_RE.test(e.reason) ? e.reason : 'other',
            source: e.source,
            ...(typeof e.messageId === 'string' && e.messageId.length > 0 ? { message_id: e.messageId.slice(0, 64) } : {}),
            ...(typeof e.conversationId === 'string' && e.conversationId.length > 0
              ? { channel: channelPart(e.conversationId).slice(0, 200) }
              : {}),
          }) + '\n'
        if (size === undefined) {
          try {
            size = statSync(opts.file).size
          } catch {
            const dir = dirname(opts.file)
            if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 })
            size = 0
          }
        }
        // Rotate before this line would grow the file past the cap; a line longer than the cap still gets written,
        // into a fresh file, rather than dropped.
        if (size > 0 && size + line.length > opts.maxBytes) {
          try {
            renameSync(opts.file, rotated)
          } catch (err) {
            warn(err)
            return
          }
          size = 0
        }
        appendFileSync(opts.file, line, { mode: 0o600 })
        size += line.length
      } catch (err) {
        warn(err)
      }
    },
  }
}
