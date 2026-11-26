// Retry of outbound Teams sends (TEAMSSDK1126). botframework-connector 4.23.3 retried 408, 5xx (except 501 and 505)
// and a 429 with a usable retry header inside its HTTP pipeline, and this module retried only what it left (412, a 429
// without such a header). The Microsoft 365 Agents SDK (@microsoft/agents-hosting 1.10.2) retries nothing: its
// HttpClient throws an HttpError (`status`, `response.headers`) on the first non-2xx answer, measured. So every one of
// those answers is retried here now, and a 429's retry header is honoured as the wait when it fits the budget.

export const SEND_RETRY_ATTEMPTS = 3
export const SEND_RETRY_BUDGET_MS = 10_000
export const SEND_RETRY_BASE_MS = 1_000
/** How long shutdown waits for sends in progress. Shorter than SEND_RETRY_BUDGET_MS on purpose: a send still retrying
 *  after this is cut off, and the exit line says so. */
export const SHUTDOWN_SEND_WAIT_MS = 5_000

export interface SendTracker {
  /** Number of withSendRetry calls that have not finished yet. */
  inFlight(): number
  /** Resolves true as soon as no send is in flight, or false after `ms` with sends still running. */
  idle(ms: number): Promise<boolean>
}

export function createSendTracker(): SendTracker & { start(): () => void } {
  let count = 0
  let waiters: Array<() => void> = []
  return {
    inFlight: () => count,
    start() {
      count++
      let finished = false
      return () => {
        if (finished) return
        finished = true
        count--
        if (count === 0) {
          const ws = waiters
          waiters = []
          for (const w of ws) w()
        }
      }
    },
    idle(ms: number): Promise<boolean> {
      if (count === 0) return Promise.resolve(true)
      return new Promise<boolean>(resolve => {
        const waiter = () => {
          clearTimeout(timer)
          resolve(true)
        }
        const timer = setTimeout(() => {
          waiters = waiters.filter(w => w !== waiter)
          resolve(false)
        }, ms)
        waiters.push(waiter)
      })
    },
  }
}

export const sendTracker = createSendTracker()

/**
 * The wait a throttling answer asks for, in ms, from 'retry-after-ms', 'x-ms-retry-after-ms' or 'Retry-After' (seconds
 * or an HTTP date), the first that parses deciding, as the old SDK pipeline read them. Undefined when none is usable.
 */
export function retryAfterMs(headers: unknown, now: () => number = Date.now): number | undefined {
  if (typeof headers !== 'object' || headers === null) return undefined
  const get = (headers as { get?: unknown }).get
  if (typeof get !== 'function') return undefined
  try {
    const read = (name: string): unknown => get.call(headers, name)
    for (const name of ['retry-after-ms', 'x-ms-retry-after-ms', 'retry-after']) {
      const v = read(name)
      if (typeof v === 'string' && v !== '' && !Number.isNaN(Number(v))) {
        const ms = Number(v) * (name === 'retry-after' ? 1000 : 1)
        return Number.isFinite(ms) ? Math.max(0, ms) : undefined
      }
    }
    const v = read('retry-after')
    if (typeof v !== 'string' || v === '') return undefined
    const at = Date.parse(v)
    return Number.isFinite(at) ? Math.max(0, at - now()) : undefined
  } catch {
    return undefined
  }
}

/** The HTTP status of a failed send: botframework-connector's RestError had `statusCode`, the Agents SDK's HttpError `status`. */
export function errorStatus(err: unknown): number | undefined {
  if (typeof err !== 'object' || err === null) return undefined
  const s = (err as { statusCode?: unknown }).statusCode ?? (err as { status?: unknown }).status
  return typeof s === 'number' ? s : undefined
}

/** 408, 412, 429 and 5xx except 501 and 505: what the old pipeline plus this module retried between them. */
export function isPluginRetryable(err: unknown): boolean {
  const status = errorStatus(err)
  if (status === undefined) return false
  if (status === 408 || status === 412 || status === 429) return true
  return status >= 500 && status <= 599 && status !== 501 && status !== 505
}

function throttleWaitMs(err: unknown, now: () => number): number | undefined {
  if (errorStatus(err) !== 429) return undefined
  let headers: unknown
  try {
    headers = (err as any).response?.headers
  } catch {
    headers = undefined
  }
  return retryAfterMs(headers, now)
}

export interface SendRetryDeps {
  sleep?: (ms: number) => Promise<void>
  random?: () => number
  now?: () => number
  log?: (line: string) => void
  tracker?: ReturnType<typeof createSendTracker>
}

export async function withSendRetry<T>(label: string, fn: () => Promise<T>, deps: SendRetryDeps = {}): Promise<T> {
  const sleep = deps.sleep ?? ((ms: number) => new Promise<void>(r => setTimeout(r, ms)))
  const random = deps.random ?? Math.random
  const now = deps.now ?? Date.now
  const log = deps.log ?? ((line: string) => process.stderr.write(line))
  const done = (deps.tracker ?? sendTracker).start()
  try {
    const start = now()
    for (let attempt = 1; ; attempt++) {
      try {
        return await fn()
      } catch (err) {
        if (attempt >= SEND_RETRY_ATTEMPTS || !isPluginRetryable(err)) throw err
        const base = SEND_RETRY_BASE_MS * 2 ** (attempt - 1)
        const backoff = Math.round(base / 2 + random() * (base / 2))
        // A 429 that names its wait gets at least that wait; the budget below still caps it.
        const delay = Math.max(backoff, throttleWaitMs(err, now) ?? 0)
        if (now() - start + delay > SEND_RETRY_BUDGET_MS) throw err
        await sleep(delay)
        // The timer may fire late: no attempt starts after the budget.
        if (now() - start > SEND_RETRY_BUDGET_MS) throw err
        log(`teams channel: ${label} got HTTP ${errorStatus(err)}, attempt ${attempt + 1}/${SEND_RETRY_ATTEMPTS} after ${delay} ms\n`)
      }
    }
  } finally {
    done()
  }
}
