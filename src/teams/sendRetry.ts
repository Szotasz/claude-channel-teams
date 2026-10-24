// Retry of outbound Teams sends. Measured on botframework-connector 4.23.3: the SDK
// (via @azure/core-rest-pipeline) already retries 408 and 5xx (except 501, 505) answers, and 429
// answers with a usable 'retry-after-ms', 'x-ms-retry-after-ms' or 'Retry-After'
// header (up to 3 retries). It does not retry 412, nor a 429 without such a header.
// Only those two cases are retried here, so nothing is retried twice.

export const SEND_RETRY_ATTEMPTS = 3
export const SEND_RETRY_BUDGET_MS = 10_000
export const SEND_RETRY_BASE_MS = 1_000

/** True when the SDK itself would retry a throttling answer with these headers. */
export function sdkRetriesThrottle(headers: unknown): boolean {
  if (typeof headers !== 'object' || headers === null) return false
  const get = (headers as { get?: unknown }).get
  if (typeof get !== 'function') return false
  try {
    const read = (name: string): unknown => get.call(headers, name)
    for (const name of ['retry-after-ms', 'x-ms-retry-after-ms', 'retry-after']) {
      const v = read(name)
      // Like the SDK: the first header that parses as a number decides, and only a finite wait is retried.
      if (typeof v === 'string' && v !== '' && !Number.isNaN(Number(v))) {
        return Number.isFinite(Number(v) * (name === 'retry-after' ? 1000 : 1))
      }
    }
    const v = read('retry-after')
    return typeof v === 'string' && v !== '' && Number.isFinite(Date.parse(v))
  } catch {
    return false
  }
}

export function isPluginRetryable(err: unknown): boolean {
  if (typeof err !== 'object' || err === null) return false
  const status = (err as { statusCode?: unknown }).statusCode
  if (typeof status !== 'number') return false
  if (status === 412) return true
  if (status !== 429) return false
  let headers: unknown
  try {
    headers = (err as any).response?.headers
  } catch {
    headers = undefined
  }
  return !sdkRetriesThrottle(headers)
}

export interface SendRetryDeps {
  sleep?: (ms: number) => Promise<void>
  random?: () => number
  now?: () => number
  log?: (line: string) => void
}

export async function withSendRetry<T>(label: string, fn: () => Promise<T>, deps: SendRetryDeps = {}): Promise<T> {
  const sleep = deps.sleep ?? ((ms: number) => new Promise<void>(r => setTimeout(r, ms)))
  const random = deps.random ?? Math.random
  const now = deps.now ?? Date.now
  const log = deps.log ?? ((line: string) => process.stderr.write(line))
  const start = now()
  for (let attempt = 1; ; attempt++) {
    try {
      return await fn()
    } catch (err) {
      if (attempt >= SEND_RETRY_ATTEMPTS || !isPluginRetryable(err)) throw err
      const base = SEND_RETRY_BASE_MS * 2 ** (attempt - 1)
      const delay = Math.round(base / 2 + random() * (base / 2))
      if (now() - start + delay > SEND_RETRY_BUDGET_MS) throw err
      await sleep(delay)
      // The timer may fire late: no attempt starts after the budget.
      if (now() - start > SEND_RETRY_BUDGET_MS) throw err
      log(`teams channel: ${label} got HTTP ${(err as { statusCode: number }).statusCode}, attempt ${attempt + 1}/${SEND_RETRY_ATTEMPTS} after ${delay} ms\n`)
    }
  }
}
