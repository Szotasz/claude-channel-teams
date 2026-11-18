/**
 * TEAMSAPPROVECARD1118: a click on a Marveen approval card is forwarded to the dashboard and never reaches the agent
 * session; the dashboard's card comes back as the invoke response. The verb cannot be used in a model-supplied card.
 */
import { describe, expect, test, mock } from 'bun:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  APPROVAL_CARD_VERB, NOT_DECIDED_CARD, isApprovalCardInvoke, parseApprovalUrl, createApprovalForwarder,
} from '../src/teams/approvalCard.ts'
import { validateAdaptiveCard, ADAPTIVE_CARD_CONTENT_TYPE } from '../src/teams/cards.ts'
import { handleMessagesPost } from '../src/teams/adapter.ts'

const TENANT = '11111111-2222-3333-4444-555555555555'
const OWNER = '6f1c2d3e-0000-4000-8000-000000000001'
const URL_OK = 'http://127.0.0.1:3420/api/approvals/teams-card-action'
const RESULT_CARD = { type: 'AdaptiveCard', version: '1.4', body: [{ type: 'TextBlock', text: 'Jóváhagyva (10:00).' }] }

function click(over: Record<string, unknown> = {}, verb: string = APPROVAL_CARD_VERB) {
  return {
    type: 'invoke', name: 'adaptiveCard/action', id: 'click-1', channelId: 'msteams', serviceUrl: 'https://smba.trafficmanager.net/emea/',
    replyToId: 'card-1', from: { id: '29:x', aadObjectId: OWNER }, conversation: { id: 'a:1owner', tenantId: TENANT },
    channelData: { tenant: { id: TENANT } },
    value: { action: { type: 'Action.Execute', verb, data: { approvalId: 'appr-1', decision: 'approve', nonce: 'ab'.repeat(32) } } },
    ...over,
  }
}

describe('isApprovalCardInvoke', () => {
  test('only an adaptiveCard/action invoke with the approval verb', () => {
    expect(isApprovalCardInvoke(click())).toBe(true)
    expect(isApprovalCardInvoke(click({}, 'something-else'))).toBe(false)
    expect(isApprovalCardInvoke(click({ type: 'message' }))).toBe(false)
    expect(isApprovalCardInvoke(click({ name: 'task/fetch' }))).toBe(false)
    expect(isApprovalCardInvoke(null)).toBe(false)
  })
})

describe('parseApprovalUrl: loopback only', () => {
  test('127.0.0.1, ::1 and localhost, http or https', () => {
    expect(parseApprovalUrl(URL_OK)).toBe(URL_OK)
    expect(parseApprovalUrl('http://[::1]:3420/x')).toBe('http://[::1]:3420/x')
    expect(parseApprovalUrl(' https://localhost:3420/x ')).toBe('https://localhost:3420/x')
  })
  test('anything else is not wired', () => {
    for (const bad of [undefined, '', 'http://10.0.0.5:3420/x', 'https://dashboard.example.com/x', 'http://127.0.0.1.evil.example/x',
      'file:///etc/passwd', 'http://user:pw@127.0.0.1:3420/x', 'not a url']) {
      expect(parseApprovalUrl(bad)).toBeUndefined()
    }
  })
})

function forwarder(over: Partial<Parameters<typeof createApprovalForwarder>[0]> = {}, answer: () => Response | Promise<Response> = () => new Response(JSON.stringify({ card: RESULT_CARD }), { status: 200 })) {
  const calls: Array<{ url: string; init: RequestInit }> = []
  const f = mock(async (url: unknown, init?: RequestInit) => { calls.push({ url: String(url), init: init ?? {} }); return answer() })
  const fwd = createApprovalForwarder({ url: URL_OK, tenantId: TENANT, isAllowed: (id) => id === OWNER, fetch: f as unknown as typeof fetch, log: () => {}, ...over })
  return { fwd, calls }
}

describe('the forwarder', () => {
  test('forwards the click with its Authorization header and the activity, and returns the dashboard card', async () => {
    const { fwd, calls } = forwarder()
    const act = click()
    expect(await fwd('Bearer channel.jwt.here', act)).toEqual(RESULT_CARD)
    expect(calls).toHaveLength(1)
    expect(calls[0]!.url).toBe(URL_OK)
    expect(calls[0]!.init.method).toBe('POST')
    expect((calls[0]!.init.headers as Record<string, string>).authorization).toBe('Bearer channel.jwt.here')
    expect(JSON.parse(String(calls[0]!.init.body))).toEqual(act)
    expect(calls[0]!.init.redirect).toBe('error')
  })
  test('not wired, another tenant, or a sender not on the allowlist: no request, the fixed card', async () => {
    for (const [over, act] of [
      [{ url: undefined }, click()],
      [{}, click({ conversation: { id: 'a:1owner', tenantId: 'other-tenant' } })],
      [{}, click({ conversation: { id: 'a:1owner' } })],
      [{}, click({ from: { id: '29:y', aadObjectId: '6f1c2d3e-0000-4000-8000-000000000002' } })],
      [{}, click({ from: { id: '29:y' } })],
    ] as const) {
      const { fwd, calls } = forwarder(over as never)
      expect(await fwd('Bearer t', act)).toEqual(NOT_DECIDED_CARD)
      expect(calls).toHaveLength(0)
    }
  })
  test('the dashboard refuses, is down, or answers without a usable card: the fixed card', async () => {
    const answers: Array<() => Response | Promise<Response>> = [
      () => new Response(JSON.stringify({ error: 'unauthorized' }), { status: 401 }),
      () => { throw new Error('ECONNREFUSED') },
      () => new Response('not json', { status: 200 }),
      () => new Response(JSON.stringify({ card: { ...RESULT_CARD, refresh: { action: { type: 'Action.Execute' } } } }), { status: 200 }),
      () => new Response(JSON.stringify({ card: { type: 'Nope' } }), { status: 200 }),
    ]
    for (const a of answers) expect(await forwarder({}, a).fwd('Bearer t', click())).toEqual(NOT_DECIDED_CARD)
  })
  test('a card the dashboard sends with a non-200 status (e.g. 503 "decide on the dashboard") is still shown', async () => {
    const { fwd } = forwarder({}, () => new Response(JSON.stringify({ card: RESULT_CARD }), { status: 503 }))
    expect(await fwd('Bearer t', click())).toEqual(RESULT_CARD)
  })
})

describe('POST /api/messages: an approval-card click never reaches CloudAdapter or the turn handler', () => {
  function req(body: unknown, auth = 'Bearer channel.jwt.here') {
    return new Request('http://127.0.0.1:3979/api/messages', { method: 'POST', body: JSON.stringify(body), headers: { 'content-type': 'application/json', authorization: auth } })
  }
  const accept = async () => ({ ok: true as const, claims: {} })
  test('an accepted click: forwarded, answered with the card as the invoke response, the adapter never called', async () => {
    const process = mock(async () => {})
    const turn = mock(async () => {})
    const seen: Array<[string | undefined, unknown]> = []
    const approvalCard = async (h: string | undefined, a: unknown) => { seen.push([h, a]); return RESULT_CARD }
    const act = click()
    const res = await handleMessagesPost(req(act), { adapter: { process } as never, inboundAuth: accept, approvalCard }, turn)
    expect(res.status).toBe(200)
    expect(await res.json()).toEqual({ statusCode: 200, type: ADAPTIVE_CARD_CONTENT_TYPE, value: RESULT_CARD })
    expect(seen).toEqual([['Bearer channel.jwt.here', act]])
    expect(process).not.toHaveBeenCalled()
    expect(turn).not.toHaveBeenCalled()
  })
  test('a refused token: 401 before the approval path', async () => {
    const approvalCard = mock(async () => RESULT_CARD)
    const res = await handleMessagesPost(req(click()), { adapter: { process: mock(async () => {}) } as never, inboundAuth: async () => ({ ok: false as const, reason: 'x' }), approvalCard }, async () => {})
    expect(res.status).toBe(401)
    expect(approvalCard).not.toHaveBeenCalled()
  })
  test('any other card click still goes to the adapter (and from there to the agent), not to the dashboard', async () => {
    const process = mock(async (_r: unknown, res: { status(n: number): void; end(b?: string): void }) => { res.status(200); res.end() })
    const approvalCard = mock(async () => RESULT_CARD)
    await handleMessagesPost(req(click({}, 'my-own-verb')), { adapter: { process } as never, inboundAuth: accept, approvalCard }, async () => {})
    expect(process).toHaveBeenCalledTimes(1)
    expect(approvalCard).not.toHaveBeenCalled()
  })
})

describe('a model-supplied card cannot use the approval verb', () => {
  test('refused anywhere in the card, other verbs pass', () => {
    const act = (verb: string) => ({ type: 'Action.Execute', title: 'OK', verb })
    expect(validateAdaptiveCard({ type: 'AdaptiveCard', version: '1.4', actions: [act(APPROVAL_CARD_VERB)] }).ok).toBe(false)
    expect(validateAdaptiveCard({ type: 'AdaptiveCard', version: '1.4', body: [{ type: 'ActionSet', actions: [act(APPROVAL_CARD_VERB)] }] }).ok).toBe(false)
    expect(validateAdaptiveCard({ type: 'AdaptiveCard', version: '1.4', actions: [act('approve-it')] }).ok).toBe(true)
  })
})

describe('server.ts wiring', () => {
  const server = readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8')
  test('the listener gets the forwarder: the loopback-checked MARVEEN_APPROVAL_URL, the bot tenant, the allowlist', () => {
    const start = server.indexOf('const listener = startListener({')
    expect(start).toBeGreaterThan(-1)
    const block = server.slice(start, server.indexOf('\n})', start))
    expect(block).toContain('approvalCard: createApprovalForwarder({ url: parseApprovalUrl(process.env.MARVEEN_APPROVAL_URL), tenantId: config.tenantId, isAllowed: (id) => allowlist.isAllowed(id) }),')
  })
})
