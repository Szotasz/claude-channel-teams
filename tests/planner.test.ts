import { describe, expect, test } from 'bun:test'

import { parsePlannerPlanIds } from '../src/config'
import {
  PLANNER_TOOL,
  createGraphTokenProvider,
  createPlannerTask,
  plannerTools,
  validatePlannerArgs,
} from '../src/teams/planner'

const TENANT = '00000000-0000-4000-8000-000000000002'
const APP = '00000000-0000-4000-8000-000000000001'
const SECRET = 'SECRET-XYZ'
const G = 'https://graph.microsoft.com/v1.0'
const U1 = '11111111-1111-4111-8111-111111111111'
const U2 = 'AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA'

interface Call {
  url: string
  method: string
  headers: Record<string, string>
  redirect: unknown
  body: string | undefined
}
type Answer = { status?: number; json?: unknown; text?: string } | Error

/** A fake fetch: records every call and answers by 'METHOD url' (or 'METHOD prefix*'). */
function makeFetch(routes: Record<string, Answer | (() => Answer)>) {
  const calls: Call[] = []
  const fn = (async (url: any, init: any) => {
    const call: Call = {
      url: String(url),
      method: init?.method ?? 'GET',
      headers: { ...(init?.headers ?? {}) },
      redirect: init?.redirect,
      body: init?.body,
    }
    calls.push(call)
    const key = `${call.method} ${call.url}`
    const hit = routes[key] ?? routes[Object.keys(routes).find(k => k.endsWith('*') && key.startsWith(k.slice(0, -1))) ?? '']
    if (hit === undefined) return new Response('{}', { status: 404 })
    const a = typeof hit === 'function' ? hit() : hit
    if (a instanceof Error) throw a
    const status = a.status ?? 200
    if (status === 204) return new Response(null, { status })
    return new Response(a.text ?? JSON.stringify(a.json ?? {}), { status })
  }) as unknown as typeof fetch
  return { fn, calls }
}

const TOKEN_URL = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`
const tokenOk = (expires = 3600): Answer => ({ json: { access_token: 'TOK', expires_in: expires } })
const provider = (f: typeof fetch, now?: () => number) =>
  createGraphTokenProvider({ tenantId: TENANT, appId: APP, appPassword: SECRET, fetch: f, now })

const base = (over: Partial<Parameters<typeof createPlannerTask>[0]> = {}) => ({
  planId: 'p1',
  title: 'T',
  assigneeIds: [] as string[],
  checklist: [] as string[],
  ...over,
})

describe('parsePlannerPlanIds', () => {
  test('undefined and empty entries give an empty set', () => {
    expect(parsePlannerPlanIds(undefined)!.size).toBe(0)
    expect(parsePlannerPlanIds(' , ')!.size).toBe(0)
  })
  test('valid list', () => {
    expect([...parsePlannerPlanIds('p1, p2')!]).toEqual(['p1', 'p2'])
  })
  test('any invalid entry gives null', () => {
    expect(parsePlannerPlanIds('p1,bad id')).toBeNull()
    expect(parsePlannerPlanIds('p1,' + 'x'.repeat(65))).toBeNull()
  })
})

describe('plannerTools', () => {
  test('off when empty, on otherwise', () => {
    expect(plannerTools(new Set())).toEqual([])
    expect(plannerTools(new Set(['p1']))).toEqual([PLANNER_TOOL])
    expect(PLANNER_TOOL.name).toBe('create_planner_task')
  })
})

describe('validatePlannerArgs', () => {
  const one = new Set(['p1'])
  const two = new Set(['p1', 'P2'])
  const err = (args: Record<string, unknown>, plans: ReadonlySet<string> = one) => {
    try {
      validatePlannerArgs(args, plans)
    } catch (e) {
      return (e as Error).message
    }
    return null
  }

  test('off', () => {
    expect(err({ title: 'x' }, new Set())).toBe('create_planner_task is off (TEAMS_PLANNER_PLAN_IDS is not set)')
  })
  test('plan selection', () => {
    expect(validatePlannerArgs({ title: 'x' }, one).planId).toBe('p1')
    expect(validatePlannerArgs({ title: 'x', plan_id: null }, one).planId).toBe('p1')
    expect(err({ title: 'x' }, two)).toBe('plan_id is required when more than one plan is configured')
    expect(err({ title: 'x', plan_id: 'p9' })).toBe('plan not allowed')
    expect(err({ title: 'x', plan_id: 5 })).toBe('plan not allowed')
    expect(err({ title: 'x', plan_id: 'p2' }, two)).toBe('plan not allowed')
    expect(validatePlannerArgs({ title: 'x', plan_id: 'P2' }, two).planId).toBe('P2')
  })
  test('title', () => {
    expect(err({})).toBe('title is required')
    expect(err({ title: '  ' })).toBe('title is required')
    expect(err({ title: 3 })).toBe('title is required')
    expect(validatePlannerArgs({ title: ' hi ' }, one).title).toBe('hi')
  })
  test('bucket', () => {
    expect(err({ title: 'x', bucket: ' ' })).toBe('bucket must be a non-empty string')
    expect(err({ title: 'x', bucket: 1 })).toBe('bucket must be a non-empty string')
    expect(validatePlannerArgs({ title: 'x', bucket: ' B ' }, one).bucket).toBe('B')
  })
  test('assignees', () => {
    expect(err({ title: 'x', assignee_ids: 'a' })).toBe('assignee_ids must be AAD object ids')
    expect(err({ title: 'x', assignee_ids: ['nope'] })).toBe('assignee_ids must be AAD object ids')
    expect(err({ title: 'x', assignee_ids: [U1, 5] })).toBe('assignee_ids must be AAD object ids')
    const r = validatePlannerArgs({ title: 'x', assignee_ids: [U2, U1, U2.toLowerCase()] }, one)
    expect(r.assigneeIds).toEqual([U2.toLowerCase(), U1])
    const many = Array.from({ length: 21 }, (_, i) => `00000000-0000-4000-8000-${String(i).padStart(12, '0')}`)
    expect(err({ title: 'x', assignee_ids: many })).toBe('at most 20 assignees')
    expect(err({ title: 'x', assignee_ids: many.slice(0, 20) })).toBeNull()
  })
  test('due_date', () => {
    const m = 'due_date must be a real YYYY-MM-DD date'
    expect(err({ title: 'x', due_date: '2026-02-30' })).toBe(m)
    expect(err({ title: 'x', due_date: '2026-2-3' })).toBe(m)
    expect(err({ title: 'x', due_date: 20260101 })).toBe(m)
    expect(validatePlannerArgs({ title: 'x', due_date: '2028-02-29' }, one).dueDate).toBe('2028-02-29')
  })
  test('priority', () => {
    expect(err({ title: 'x', priority: 'high' })).toBe('priority must be urgent, important, medium or low')
    expect(err({ title: 'x', priority: 3 })).toBe('priority must be urgent, important, medium or low')
    const p = (priority: string) => validatePlannerArgs({ title: 'x', priority }, one).priority
    expect([p('urgent'), p('important'), p('medium'), p('low')]).toEqual([1, 3, 5, 9])
  })
  test('description', () => {
    expect(err({ title: 'x', description: 5 })).toBe('description must be a string')
    expect(validatePlannerArgs({ title: 'x', description: '' }, one).description).toBeUndefined()
  })
  test('link', () => {
    expect(err({ title: 'x', link: 'http://x' })).toBe('link must be an https URL')
    expect(err({ title: 'x', link: 'not a url' })).toBe('link must be an https URL')
    expect(err({ title: 'x', link: 5 })).toBe('link must be an https URL')
    expect(err({ title: 'x', link: 'https://example.com/t' })).toBeNull()
  })
  test('checklist', () => {
    const m = 'checklist must be at most 20 non-empty strings'
    expect(err({ title: 'x', checklist: Array(21).fill('a') })).toBe(m)
    expect(err({ title: 'x', checklist: ['a', ' '] })).toBe(m)
    expect(err({ title: 'x', checklist: 'a' })).toBe(m)
    expect(validatePlannerArgs({ title: 'x', checklist: [' a '] }, one).checklist).toEqual(['a'])
  })

  test('an invalid input makes no network call, not even the token request', async () => {
    const bad: Record<string, unknown>[] = [
      {},
      { title: 'x', plan_id: 'nope' },
      { title: 'x', bucket: '' },
      { title: 'x', assignee_ids: ['x'] },
      { title: 'x', due_date: '2026-02-30' },
      { title: 'x', priority: 'high' },
      { title: 'x', description: 1 },
      { title: 'x', link: 'http://x' },
      { title: 'x', checklist: Array(21).fill('a') },
    ]
    for (const args of bad) {
      const f = makeFetch({ [`POST ${TOKEN_URL}`]: tokenOk() })
      const p = provider(f.fn)
      try {
        await createPlannerTask(validatePlannerArgs(args, one), { getToken: () => p.getToken(), fetch: f.fn })
      } catch {}
      expect(f.calls.length).toBe(0)
    }
  })
})

describe('createPlannerTask', () => {
  const happyRoutes = () => ({
    [`POST ${TOKEN_URL}`]: tokenOk(),
    [`GET ${G}/planner/plans/p1/buckets`]: {
      json: { value: [{ id: 'b0', name: 'Other' }], '@odata.nextLink': `${G}/planner/plans/p1/buckets?page=2` },
    },
    [`GET ${G}/planner/plans/p1/buckets?page=2`]: { json: { value: [{ id: 'b1', name: 'Todo' }] } },
    [`POST ${G}/planner/tasks`]: { status: 201, json: { id: 'T1' } },
    [`GET ${G}/planner/tasks/T1/details`]: { json: { '@odata.etag': 'W/"abc"' } },
    [`PATCH ${G}/planner/tasks/T1/details`]: { status: 204 },
  })

  test('full happy path', async () => {
    const f = makeFetch(happyRoutes())
    const p = provider(f.fn)
    let n = 0
    const input = validatePlannerArgs(
      {
        title: 'Do it',
        bucket: 'Todo',
        assignee_ids: [U1],
        due_date: '2026-11-05',
        priority: 'important',
        description: 'Desc',
        link: 'https://example.com/t',
        checklist: ['c', 'a', 'b'],
      },
      new Set(['p1']),
    )
    const r = await createPlannerTask(input, {
      getToken: () => p.getToken(),
      fetch: f.fn,
      randomUUID: () => `uuid-${n++}`,
    })
    expect(r).toEqual({ taskId: 'T1' })

    const [tok, b1, b2, post, get, patch] = f.calls
    expect(f.calls.length).toBe(6)
    expect(tok!.url).toBe(TOKEN_URL)
    const form = new URLSearchParams(tok!.body)
    expect(form.get('grant_type')).toBe('client_credentials')
    expect(form.get('client_id')).toBe(APP)
    expect(form.get('client_secret')).toBe(SECRET)
    expect(form.get('scope')).toBe('https://graph.microsoft.com/.default')
    expect(b1!.headers.Authorization).toBe('Bearer TOK')
    expect(b2!.url).toBe(`${G}/planner/plans/p1/buckets?page=2`)

    expect(post!.headers['Content-Type']).toBe('application/json')
    expect(JSON.parse(post!.body!)).toEqual({
      planId: 'p1',
      title: 'Do it',
      bucketId: 'b1',
      assignments: { [U1]: { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' } },
      dueDateTime: '2026-11-05T12:00:00Z',
      priority: 3,
    })
    expect(get!.method).toBe('GET')
    expect(patch!.headers['If-Match']).toBe('W/"abc"')
    const pb = JSON.parse(patch!.body!)
    expect(Object.keys(pb).sort()).toEqual(['checklist', 'description'])
    expect(pb.description).toBe('Desc\n\nThread: https://example.com/t')
    const items = Object.values(pb.checklist) as any[]
    expect(items.every(i => i['@odata.type'] === 'microsoft.graph.plannerChecklistItem' && i.isChecked === false)).toBe(true)
    items.sort((x, y) => (x.orderHint < y.orderHint ? -1 : 1))
    expect(items.map(i => i.title)).toEqual(['c', 'a', 'b'])
    expect(items.map(i => i.orderHint)).toEqual([' !', ' ! !', ' ! ! !'])
    expect(Object.keys(pb.checklist)).toEqual(['uuid-0', 'uuid-1', 'uuid-2'])

    for (const c of f.calls) expect(c.redirect).toBe('error')
    for (const c of f.calls.slice(1)) {
      expect(c.body ?? '').not.toContain(SECRET)
      expect(JSON.stringify(c.headers)).not.toContain(SECRET)
      expect(c.url).not.toContain(SECRET)
    }
  })

  test('bucket id wins over a bucket whose name equals that id', async () => {
    const routes = {
      ...happyRoutes(),
      [`GET ${G}/planner/plans/p1/buckets`]: {
        json: { value: [{ id: 'x1', name: 'b9' }, { id: 'b9', name: 'real' }] },
      },
    }
    const f = makeFetch(routes)
    const p = provider(f.fn)
    await createPlannerTask(base({ bucket: 'b9' }), { getToken: () => p.getToken(), fetch: f.fn })
    const post = f.calls.find(c => c.method === 'POST' && c.url.endsWith('/planner/tasks'))!
    expect(JSON.parse(post.body!).bucketId).toBe('b9')
  })

  test('unknown bucket creates nothing', async () => {
    const f = makeFetch(happyRoutes())
    const p = provider(f.fn)
    await expect(
      createPlannerTask(base({ bucket: 'missing' }), { getToken: () => p.getToken(), fetch: f.fn }),
    ).rejects.toThrow('bucket not found')
    expect(f.calls.some(c => c.method === 'POST' && c.url.endsWith('/planner/tasks'))).toBe(false)
  })

  test('no details requests without description, link or checklist', async () => {
    const f = makeFetch(happyRoutes())
    const p = provider(f.fn)
    await createPlannerTask(base(), { getToken: () => p.getToken(), fetch: f.fn })
    expect(f.calls.some(c => c.url.includes('/details'))).toBe(false)
    const post = f.calls.find(c => c.method === 'POST' && c.url.endsWith('/planner/tasks'))!
    expect(JSON.parse(post.body!)).toEqual({ planId: 'p1', title: 'T' })
  })

  test('link alone becomes the description', async () => {
    const f = makeFetch(happyRoutes())
    const p = provider(f.fn)
    await createPlannerTask(base({ link: 'https://example.com/t' }), { getToken: () => p.getToken(), fetch: f.fn })
    const patch = f.calls.find(c => c.method === 'PATCH')!
    expect(JSON.parse(patch.body!)).toEqual({ description: 'Thread: https://example.com/t' })
  })

  test('token is cached when long-lived, not when expires_in <= 300', async () => {
    for (const [expires, tokens] of [[3600, 1], [100, 2]] as const) {
      const f = makeFetch({ ...happyRoutes(), [`POST ${TOKEN_URL}`]: tokenOk(expires) })
      const p = provider(f.fn)
      const deps = { getToken: () => p.getToken(), fetch: f.fn }
      await createPlannerTask(base(), deps)
      await createPlannerTask(base(), deps)
      expect(f.calls.filter(c => c.url === TOKEN_URL).length).toBe(tokens)
    }
  })

  test('token cache expires with the injected clock', async () => {
    const f = makeFetch({ [`POST ${TOKEN_URL}`]: tokenOk(3600) })
    let t = 0
    const p = provider(f.fn, () => t)
    await p.getToken()
    t = 3299_000
    await p.getToken()
    expect(f.calls.length).toBe(1)
    t = 3300_000
    await p.getToken()
    expect(f.calls.length).toBe(2)
  })

  test('token errors never carry answer text', async () => {
    const run = async (error: string) => {
      const f = makeFetch({
        [`POST ${TOKEN_URL}`]: { status: 401, json: { error, error_description: 'contains SECRET-XYZ' } },
      })
      return provider(f.fn).getToken().catch(e => (e as Error).message)
    }
    expect(await run('invalid_client')).toBe('Graph token request failed (HTTP 401 invalid_client)')
    expect(await run('bad code!')).toBe('Graph token request failed (HTTP 401)')
  })

  test('a 200 token answer of the wrong shape fails', async () => {
    const f = makeFetch({ [`POST ${TOKEN_URL}`]: { json: { access_token: 'x', expires_in: '3600' } } })
    await expect(provider(f.fn).getToken()).rejects.toThrow('Graph token request failed (HTTP 200)')
  })

  test('Graph error shows the status and a safe code only', async () => {
    const f = makeFetch({
      ...happyRoutes(),
      [`POST ${G}/planner/tasks`]: { status: 403, json: { error: { code: 'Forbidden', message: 'leak SECRET-XYZ' } } },
    })
    const p = provider(f.fn)
    const e = await createPlannerTask(base(), { getToken: () => p.getToken(), fetch: f.fn }).catch(e => e as Error)
    expect((e as Error).message).toBe('Graph HTTP 403 Forbidden')
    expect((e as Error).message).not.toContain('leak')
  })

  test('a Graph error code outside the safe pattern is dropped', async () => {
    const f = makeFetch({
      ...happyRoutes(),
      [`POST ${G}/planner/tasks`]: { status: 403, json: { error: { code: 'bad code!', message: 'x' } } },
    })
    const p = provider(f.fn)
    const e = await createPlannerTask(base(), { getToken: () => p.getToken(), fetch: f.fn }).catch(e => e as Error)
    expect((e as Error).message).toBe('Graph HTTP 403')
  })

  test('fetch failures never pass on the error message', async () => {
    const run = async (err: Error) => {
      const f = makeFetch({ [`POST ${TOKEN_URL}`]: err })
      return provider(f.fn).getToken().catch(e => (e as Error).message)
    }
    expect(await run(new Error('boom https://graph.microsoft.com/v1.0/x'))).toBe(
      'Graph request failed (network error)',
    )
    expect(await run(Object.assign(new Error('x'), { code: 'ECONNRESET' }))).toBe(
      'Graph request failed (ECONNRESET)',
    )
  })

  test('a next link outside Graph is not requested', async () => {
    const f = makeFetch({
      ...happyRoutes(),
      [`GET ${G}/planner/plans/p1/buckets`]: { json: { value: [], '@odata.nextLink': 'https://evil.example/next' } },
    })
    const p = provider(f.fn)
    await expect(
      createPlannerTask(base({ bucket: 'Todo' }), { getToken: () => p.getToken(), fetch: f.fn }),
    ).rejects.toThrow('Graph next link outside graph.microsoft.com/v1.0')
    expect(f.calls.some(c => c.url.includes('evil.example'))).toBe(false)
  })

  test('endless paging stops', async () => {
    const f = makeFetch({
      [`POST ${TOKEN_URL}`]: tokenOk(),
      [`GET ${G}/planner/plans/p1/buckets*`]: {
        json: { value: [], '@odata.nextLink': `${G}/planner/plans/p1/buckets?more` },
      },
    })
    const p = provider(f.fn)
    await expect(
      createPlannerTask(base({ bucket: 'Todo' }), { getToken: () => p.getToken(), fetch: f.fn }),
    ).rejects.toThrow('too many pages')
    expect(f.calls.filter(c => c.method === 'GET').length).toBe(50)
  })

  test('POST answer without id', async () => {
    const f = makeFetch({ ...happyRoutes(), [`POST ${G}/planner/tasks`]: { status: 201, json: {} } })
    const p = provider(f.fn)
    await expect(createPlannerTask(base(), { getToken: () => p.getToken(), fetch: f.fn })).rejects.toThrow(
      'Graph returned no task id',
    )
  })

  test('details failures report the created task id', async () => {
    const prefix = 'task created (task_id T1) but setting its details failed:'
    const cases: Record<string, Answer>[] = [
      { [`GET ${G}/planner/tasks/T1/details`]: { status: 500, json: {} } },
      { [`GET ${G}/planner/tasks/T1/details`]: { json: {} } },
      { [`PATCH ${G}/planner/tasks/T1/details`]: { status: 412, json: { error: { code: 'PreconditionFailed' } } } },
    ]
    for (const over of cases) {
      const f = makeFetch({ ...happyRoutes(), ...over })
      const p = provider(f.fn)
      const e = await createPlannerTask(base({ checklist: ['a'] }), {
        getToken: () => p.getToken(),
        fetch: f.fn,
      }).catch(e => e as Error)
      expect((e as Error).message.startsWith(prefix)).toBe(true)
    }
  })
})
