/**
 * `create_planner_task` implementation: creates a Microsoft Planner task
 * through Microsoft Graph with the bot's own app-only (client-credentials)
 * token.
 *
 * Permission: Graph application permission Tasks.ReadWrite.All with admin
 * consent. It covers every plan in the tenant, so the tool only writes to the
 * plans the operator lists in TEAMS_PLANNER_PLAN_IDS (docs/planner.md).
 *
 * The client secret goes only to the token endpoint. No token or secret is
 * ever logged or put in an error message, and error text from Graph or from a
 * failed fetch (which can contain URLs) is never passed on. No request follows
 * a redirect.
 */

const G = 'https://graph.microsoft.com/v1.0'
const GRAPH_PREFIX = `${G}/`
const MAX_ITEMS = 20
const MAX_PAGES = 50
const CODE_RE = /^[A-Za-z0-9_.]{1,64}$/
const GUID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/

type FetchFn = typeof fetch

/** Every request goes through here: no redirects, and no fetch error text is passed on. */
async function safeFetch(
  fetchImpl: FetchFn | undefined,
  url: string,
  init: RequestInit,
): Promise<Response> {
  try {
    return await (fetchImpl ?? globalThis.fetch)(url, { ...init, redirect: 'error' })
  } catch (err) {
    const code = (err as { code?: unknown } | null)?.code
    const safe = typeof code === 'string' && /^[A-Za-z_]{1,40}$/.test(code) ? code : 'network error'
    throw new Error(`Graph request failed (${safe})`)
  }
}

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

export function createGraphTokenProvider(deps: {
  tenantId: string
  appId: string
  appPassword: string
  fetch?: typeof fetch
  now?: () => number
}): { getToken(): Promise<string> } {
  const now = deps.now ?? Date.now
  let cached: { token: string; until: number } | null = null
  return {
    async getToken() {
      if (cached && now() < cached.until) return cached.token
      const res = await safeFetch(
        deps.fetch,
        `https://login.microsoftonline.com/${encodeURIComponent(deps.tenantId)}/oauth2/v2.0/token`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({
            grant_type: 'client_credentials',
            client_id: deps.appId,
            client_secret: deps.appPassword,
            scope: 'https://graph.microsoft.com/.default',
          }).toString(),
        },
      )
      let json: unknown
      try {
        json = await res.json()
      } catch {
        json = undefined
      }
      if (
        res.status === 200 &&
        isObject(json) &&
        typeof json.access_token === 'string' &&
        typeof json.expires_in === 'number'
      ) {
        const until = now() + (json.expires_in - 300) * 1000
        cached = until > now() ? { token: json.access_token, until } : null
        return json.access_token
      }
      let code: string | undefined
      if (isObject(json)) {
        const e = json.error
        if (typeof e === 'string') code = e
        else if (isObject(e) && typeof e.code === 'string') code = e.code
      }
      if (code !== undefined && !CODE_RE.test(code)) code = undefined
      throw new Error(`Graph token request failed (HTTP ${res.status}${code ? ' ' + code : ''})`)
    },
  }
}

/** Graph request helper shared with other Graph tools: Graph URLs only, no redirects, error codes filtered. */
export async function graph(
  getToken: () => Promise<string>,
  fetchImpl: FetchFn | undefined,
  method: string,
  url: string,
  body?: unknown,
  extraHeaders?: Record<string, string>,
): Promise<unknown> {
  if (!url.startsWith(GRAPH_PREFIX)) throw new Error('refusing a non-Graph URL')
  const token = await getToken()
  const headers: Record<string, string> = { Authorization: `Bearer ${token}`, ...extraHeaders }
  if (body !== undefined) headers['Content-Type'] = 'application/json'
  const res = await safeFetch(fetchImpl, url, {
    method,
    headers,
    body: body === undefined ? undefined : JSON.stringify(body),
  })
  if (res.status >= 200 && res.status < 300) {
    if (res.status === 204) return null
    try {
      return await res.json()
    } catch {
      throw new Error('Graph returned an invalid response')
    }
  }
  let code: string | undefined
  try {
    const json: unknown = await res.json()
    if (isObject(json) && isObject(json.error) && typeof json.error.code === 'string') {
      code = json.error.code
    }
  } catch {}
  if (code !== undefined && !CODE_RE.test(code)) code = undefined
  throw new Error(`Graph HTTP ${res.status}${code ? ' ' + code : ''}`)
}

async function graphPaged(
  getToken: () => Promise<string>,
  fetchImpl: FetchFn | undefined,
  firstUrl: string,
): Promise<unknown[]> {
  const items: unknown[] = []
  let url: string | undefined = firstUrl
  for (let page = 0; url !== undefined; page++) {
    if (page >= MAX_PAGES) throw new Error('too many pages')
    const json = await graph(getToken, fetchImpl, 'GET', url)
    if (isObject(json) && Array.isArray(json.value)) items.push(...json.value)
    const next = isObject(json) ? json['@odata.nextLink'] : undefined
    if (next === undefined || next === null) {
      url = undefined
    } else if (typeof next === 'string' && next.startsWith(GRAPH_PREFIX)) {
      url = next
    } else {
      throw new Error('Graph next link outside graph.microsoft.com/v1.0')
    }
  }
  return items
}

export interface PlannerTaskInput {
  planId: string
  title: string
  bucket?: string
  assigneeIds: string[]
  dueDate?: string
  priority?: 1 | 3 | 5 | 9
  description?: string
  link?: string
  checklist: string[]
}

const PRIORITIES = { urgent: 1, important: 3, medium: 5, low: 9 } as const

/** Checks the tool arguments; throws on the first problem. No network. */
export function validatePlannerArgs(
  args: Record<string, unknown>,
  allowedPlans: ReadonlySet<string>,
): PlannerTaskInput {
  if (allowedPlans.size === 0) {
    throw new Error('create_planner_task is off (TEAMS_PLANNER_PLAN_IDS is not set)')
  }
  let planId: string
  if (args.plan_id === undefined || args.plan_id === null) {
    if (allowedPlans.size !== 1) {
      throw new Error('plan_id is required when more than one plan is configured')
    }
    planId = [...allowedPlans][0]!
  } else if (typeof args.plan_id === 'string' && allowedPlans.has(args.plan_id)) {
    planId = args.plan_id
  } else {
    throw new Error('plan not allowed')
  }

  if (typeof args.title !== 'string' || args.title.trim() === '') {
    throw new Error('title is required')
  }
  const input: PlannerTaskInput = {
    planId,
    title: args.title.trim(),
    assigneeIds: [],
    checklist: [],
  }

  if (args.bucket !== undefined && args.bucket !== null) {
    if (typeof args.bucket !== 'string' || args.bucket.trim() === '') {
      throw new Error('bucket must be a non-empty string')
    }
    input.bucket = args.bucket.trim()
  }

  if (args.assignee_ids !== undefined && args.assignee_ids !== null) {
    const raw = args.assignee_ids
    if (!Array.isArray(raw) || !raw.every(x => typeof x === 'string' && GUID_RE.test(x))) {
      throw new Error('assignee_ids must be AAD object ids')
    }
    input.assigneeIds = [...new Set((raw as string[]).map(x => x.toLowerCase()))]
    if (input.assigneeIds.length > MAX_ITEMS) throw new Error('at most 20 assignees')
  }

  if (args.due_date !== undefined && args.due_date !== null) {
    const d = args.due_date
    const m = typeof d === 'string' ? /^(\d{4})-(\d{2})-(\d{2})$/.exec(d) : null
    let ok = false
    if (m) {
      const [y, mo, da] = [Number(m[1]), Number(m[2]), Number(m[3])]
      const t = new Date(Date.UTC(y, mo - 1, da))
      ok = t.getUTCFullYear() === y && t.getUTCMonth() === mo - 1 && t.getUTCDate() === da
    }
    if (!ok) throw new Error('due_date must be a real YYYY-MM-DD date')
    input.dueDate = d as string
  }

  if (args.priority !== undefined && args.priority !== null) {
    const p = args.priority
    if (typeof p !== 'string' || !Object.hasOwn(PRIORITIES, p)) {
      throw new Error('priority must be urgent, important, medium or low')
    }
    input.priority = PRIORITIES[p as keyof typeof PRIORITIES]
  }

  if (args.description !== undefined && args.description !== null) {
    if (typeof args.description !== 'string') throw new Error('description must be a string')
    if (args.description !== '') input.description = args.description
  }

  if (args.link !== undefined && args.link !== null) {
    let ok = false
    if (typeof args.link === 'string') {
      try {
        ok = new URL(args.link).protocol === 'https:'
      } catch {}
    }
    if (!ok) throw new Error('link must be an https URL')
    input.link = args.link as string
  }

  if (args.checklist !== undefined && args.checklist !== null) {
    const c = args.checklist
    if (
      !Array.isArray(c) ||
      c.length > MAX_ITEMS ||
      !c.every(x => typeof x === 'string' && x.trim() !== '')
    ) {
      throw new Error('checklist must be at most 20 non-empty strings')
    }
    input.checklist = (c as string[]).map(x => x.trim())
  }
  return input
}

export async function createPlannerTask(
  input: PlannerTaskInput,
  deps: {
    getToken: () => Promise<string>
    fetch?: typeof fetch
    randomUUID?: () => string
  },
): Promise<{ taskId: string }> {
  const { getToken, fetch: fetchImpl } = deps
  const call = (method: string, url: string, body?: unknown, headers?: Record<string, string>) =>
    graph(getToken, fetchImpl, method, url, body, headers)

  let bucketId: string | undefined
  if (input.bucket !== undefined) {
    const buckets = (await graphPaged(
      getToken,
      fetchImpl,
      `${G}/planner/plans/${encodeURIComponent(input.planId)}/buckets`,
    )).filter(isObject)
    const hit =
      buckets.find(b => b.id === input.bucket) ?? buckets.find(b => b.name === input.bucket)
    if (!hit || typeof hit.id !== 'string') throw new Error('bucket not found')
    bucketId = hit.id
  }

  const body: Record<string, unknown> = { planId: input.planId, title: input.title }
  if (bucketId !== undefined) body.bucketId = bucketId
  if (input.assigneeIds.length > 0) {
    body.assignments = Object.fromEntries(
      input.assigneeIds.map(id => [
        id,
        { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' },
      ]),
    )
  }
  if (input.dueDate !== undefined) body.dueDateTime = `${input.dueDate}T12:00:00Z`
  if (input.priority !== undefined) body.priority = input.priority

  const created = await call('POST', `${G}/planner/tasks`, body)
  const id = isObject(created) ? created.id : undefined
  if (typeof id !== 'string' || id === '') throw new Error('Graph returned no task id')

  if (input.description || input.link || input.checklist.length > 0) {
    try {
      const url = `${G}/planner/tasks/${encodeURIComponent(id)}/details`
      const details = await call('GET', url)
      const etag = isObject(details) ? details['@odata.etag'] : undefined
      if (typeof etag !== 'string' || etag === '') throw new Error('Graph returned no details etag')
      const patch: Record<string, unknown> = {}
      const parts = [input.description, input.link ? 'Thread: ' + input.link : undefined].filter(
        (p): p is string => !!p,
      )
      if (parts.length > 0) patch.description = parts.join('\n\n')
      if (input.checklist.length > 0) {
        const uuid = deps.randomUUID ?? (() => crypto.randomUUID())
        patch.checklist = Object.fromEntries(
          input.checklist.map((title, i) => [
            uuid(),
            {
              '@odata.type': 'microsoft.graph.plannerChecklistItem',
              title,
              isChecked: false,
              orderHint: ' !'.repeat(i + 1),
            },
          ]),
        )
      }
      await call('PATCH', url, patch, { 'If-Match': etag })
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err)
      throw new Error(`task created (task_id ${id}) but setting its details failed: ${msg}`)
    }
  }
  return { taskId: id }
}

export const PLANNER_TOOL = {
  name: 'create_planner_task',
  description:
    "Creates a task in one of the operator's allow-listed Microsoft Planner plans. " +
    'Returns {"task_id"}.',
  inputSchema: {
    type: 'object',
    properties: {
      plan_id: {
        type: 'string',
        description: 'Planner plan id. Optional when exactly one plan is configured.',
      },
      title: { type: 'string', description: 'Task title.' },
      bucket: { type: 'string', description: 'Bucket id or exact bucket name.' },
      assignee_ids: {
        type: 'array',
        items: { type: 'string' },
        maxItems: 20,
        description: 'AAD object ids of the assignees, at most 20.',
      },
      due_date: {
        type: 'string',
        description: 'Due date, YYYY-MM-DD; stored as 12:00 UTC.',
      },
      priority: {
        type: 'string',
        enum: ['urgent', 'important', 'medium', 'low'],
        description: 'Task priority.',
      },
      description: { type: 'string', description: 'Task description.' },
      link: {
        type: 'string',
        description: "https URL; added to the description as 'Thread: <link>'.",
      },
      checklist: {
        type: 'array',
        items: { type: 'string' },
        maxItems: 20,
        description: 'Checklist item titles, in order, at most 20.',
      },
    },
    required: ['title'],
  },
}

export function plannerTools(planIds: ReadonlySet<string>): (typeof PLANNER_TOOL)[] {
  return planIds.size === 0 ? [] : [PLANNER_TOOL]
}
