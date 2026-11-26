/**
 * TEAMSSDK1126: the error a failed call throws changed shape. botframework-connector's RestError had `statusCode` and
 * `body`; the Microsoft 365 Agents SDK's HttpError has `status` and `response.data`. Every reader of an error's status
 * must take both.
 */
import { describe, expect, test } from 'bun:test'
import { isNotMemberError } from '../src/teams/mentions'
import { connectorRequest } from '../src/teams/sdk'

const httpError = (status: number, data?: unknown) =>
  Object.assign(new Error(`Request failed with status ${status}`), { status, response: { status, data } })

describe('isNotMemberError', () => {
  test('both shapes', () => {
    expect(isNotMemberError(httpError(404))).toBe(true)
    expect(isNotMemberError({ statusCode: 404 })).toBe(true)
    expect(isNotMemberError(httpError(400, { error: { code: 'MemberNotFoundInConversation' } }))).toBe(true)
    expect(isNotMemberError({ body: { error: { code: 'MemberNotFoundInConversation' } } })).toBe(true)
    expect(isNotMemberError(httpError(403))).toBe(false)
    expect(isNotMemberError(httpError(500, { error: { code: 'ServiceError' } }))).toBe(false)
    expect(isNotMemberError(null)).toBe(false)
  })
})

describe('connectorRequest', () => {
  const KEY = Symbol('k')
  const ctx = (request: (c: any) => Promise<{ status: number }>) => ({ adapter: { ConnectorClientKey: KEY }, turnState: new Map([[KEY, { httpClient: { request } }]]) })

  test('the status of a success and of an HttpError; another error passes through', async () => {
    expect(await connectorRequest(ctx(async () => ({ status: 201 })), { method: 'PUT', url: 'https://x/y' })).toBe(201)
    expect(await connectorRequest(ctx(async () => { throw httpError(429) }), { method: 'PUT', url: 'https://x/y' })).toBe(429)
    const boom = new Error('socket hang up')
    await expect(connectorRequest(ctx(async () => { throw boom }), { method: 'PUT', url: 'https://x/y' })).rejects.toBe(boom)
  })

  test('a JSON body goes as data with its content type, a bodiless call sends neither', async () => {
    const seen: any[] = []
    await connectorRequest(ctx(async c => { seen.push(c); return { status: 200 } }), { method: 'POST', url: 'https://x/y', data: { a: 1 } })
    await connectorRequest(ctx(async c => { seen.push(c); return { status: 200 } }), { method: 'DELETE', url: 'https://x/z' })
    expect(seen[0]).toEqual({ method: 'POST', url: 'https://x/y', data: { a: 1 }, headers: { 'Content-Type': 'application/json' } })
    expect(seen[1]).toEqual({ method: 'DELETE', url: 'https://x/z' })
  })

  test('no connector client on the turn: a plain error', async () => {
    await expect(connectorRequest({ adapter: { ConnectorClientKey: KEY }, turnState: new Map() }, { method: 'PUT', url: 'https://x' })).rejects.toThrow('no connector client for this turn')
  })
})
