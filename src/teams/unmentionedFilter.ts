import { spawn, type ChildProcess } from 'node:child_process'

/** What the filter decided for one unmentioned message. */
export type FilterVerdict = 'deliver' | 'hold'

export interface UnmentionedFilterInput {
  text: string
  fromName?: string
  conversationId: string
  conversationType?: string
  botName?: string
}

/** reason is set only when the filter failed and the message is delivered anyway. */
export interface FilterDecision {
  verdict: FilterVerdict
  reason?: string
}

export type UnmentionedFilter = (input: UnmentionedFilterInput) => Promise<FilterDecision>

/** Held messages kept per conversation is configurable; these two bounds are not. */
export const HELD_MAX_CONVERSATIONS = 200
export const HELD_MAX_CHARS = 1000

/**
 * Runs `cmd` (an absolute path; no shell, no arguments) with the input as one JSON line on stdin.
 * Exit 0 = deliver, exit 1 = hold. Anything else -- another exit code, a signal, a spawn error,
 * or no exit within timeoutMs (the process is then killed with SIGKILL) -- = deliver, with the
 * reason. Fail-open on purpose: a broken filter gives today's behaviour, it never swallows a
 * message. Never rejects.
 */
export function createCommandFilter(cmd: string, timeoutMs: number): UnmentionedFilter {
  return input =>
    new Promise<FilterDecision>(resolve => {
      let settled = false
      let timer: ReturnType<typeof setTimeout> | undefined
      const done = (d: FilterDecision) => {
        if (settled) return
        settled = true
        if (timer) clearTimeout(timer)
        resolve(d)
      }
      let child: ChildProcess
      try {
        child = spawn(cmd, [], { stdio: ['pipe', 'ignore', 'ignore'] })
      } catch (err) {
        done({ verdict: 'deliver', reason: `spawn failed: ${err instanceof Error ? err.message : String(err)}` })
        return
      }
      timer = setTimeout(() => {
        child.kill('SIGKILL')
        done({ verdict: 'deliver', reason: `timed out after ${timeoutMs} ms` })
      }, timeoutMs)
      child.on('error', err => done({ verdict: 'deliver', reason: `spawn failed: ${err.message}` }))
      child.on('exit', (code, signal) => {
        if (code === 0) done({ verdict: 'deliver' })
        else if (code === 1) done({ verdict: 'hold' })
        else done({ verdict: 'deliver', reason: signal ? `killed by ${signal}` : `exit code ${code}` })
      })
      // A filter that exits without reading its stdin must not crash the plugin (EPIPE).
      child.stdin?.on('error', () => {})
      child.stdin?.end(JSON.stringify(input) + '\n')
    })
}

export interface HeldMessages {
  hold(conversationId: string, fromName: string | undefined, text: string): void
  /** Returns the held lines of this conversation, oldest first, and forgets them. */
  take(conversationId: string): string[]
}

export function createHeldMessages(limitPerConversation: number): HeldMessages {
  const byConversation = new Map<string, string[]>()
  return {
    hold(conversationId, fromName, text) {
      const body = text.length > HELD_MAX_CHARS ? `${text.slice(0, HELD_MAX_CHARS)}…` : text
      const line = `${fromName && fromName.length > 0 ? fromName : 'Someone'}: ${body}`
      const list = byConversation.get(conversationId) ?? []
      // Re-insert, so the Map's first key is always the conversation held into longest ago.
      byConversation.delete(conversationId)
      list.push(line)
      while (list.length > limitPerConversation) list.shift()
      byConversation.set(conversationId, list)
      while (byConversation.size > HELD_MAX_CONVERSATIONS) {
        const oldest = byConversation.keys().next().value
        if (oldest === undefined) break
        byConversation.delete(oldest)
      }
    },
    take(conversationId) {
      const list = byConversation.get(conversationId) ?? []
      byConversation.delete(conversationId)
      return list
    },
  }
}

/** Puts the held lines in front of a delivered message's text; no held lines = the text unchanged. */
export function withHeldContext(text: string, held: string[]): string {
  if (held.length === 0) return text
  return [
    '[Earlier messages in this conversation that were not addressed to you, oldest first]',
    ...held,
    '[End of earlier messages]',
    '',
    text,
  ].join('\n')
}
