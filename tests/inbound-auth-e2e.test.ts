/**
 * TEAMSSDK1126: the inbound path end to end on the Microsoft 365 Agents SDK, with the REAL channel-token verifier, the
 * REAL CloudAdapter the server builds (createCloudAdapter) and the REAL turn handler. Only the outbound token fetch is
 * stubbed (the adapter's connector and user-token clients), and the verifier's key fetch is injected, so nothing
 * leaves the process.
 *
 * Every forged or broken request is a 401 and the handler never runs; a malformed activity with a good token is a 400;
 * a good token from another tenant is dropped by the tenant pin; a good one is delivered, and the turn carries the
 * verified claims as its identity (the Agents SDK runs a turn without request.user as anonymous).
 */
import { describe, expect, mock, test } from 'bun:test'
import { generateKeyPairSync, sign as cryptoSign, type KeyObject } from 'node:crypto'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createChannelTokenVerifier, CHANNEL_OPENID_METADATA_URL, CHANNEL_TOKEN_ISSUER } from '../src/teams/channelToken.ts'
import { handleMessagesPost, makeTurnHandler } from '../src/teams/adapter.ts'
import { createCloudAdapter } from '../src/teams/auth.ts'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'

const APP = '00000000-aaaa-bbbb-cccc-000000000001'
const TENANT = '00000000-0000-4000-8000-000000000002'
const OTHER_TENANT = '00000000-0000-4000-8000-000000000003'
const SENDER = '00000000-0000-4000-8000-0000000000aa'
const SURL = 'https://smba.trafficmanager.net/emea/'
const JWKS_URL = 'https://login.botframework.com/v1/.well-known/keys'
const NOW = Date.now()

function rsa() {
  const { privateKey, publicKey } = generateKeyPairSync('rsa', { modulusLength: 2048 })
  return { privateKey, jwk: publicKey.export({ format: 'jwk' }) as { n: string; e: string } }
}
const K1 = rsa()
const K2 = rsa()
const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url')
function token(over: Record<string, unknown> = {}, key: KeyObject = K1.privateKey): string {
  const t = Math.floor(NOW / 1000)
  const h = b64({ alg: 'RS256', kid: 'k1', typ: 'JWT' })
  const p = b64({ iss: CHANNEL_TOKEN_ISSUER, aud: APP, exp: t + 3600, nbf: t - 10, serviceurl: SURL, ...over })
  return `${h}.${p}.${cryptoSign('RSA-SHA256', Buffer.from(`${h}.${p}`), key).toString('base64url')}`
}
const keysFetch = (async (input: unknown) => {
  const url = String(input)
  if (url === CHANNEL_OPENID_METADATA_URL) return new Response(JSON.stringify({ jwks_uri: JWKS_URL }), { status: 200 })
  if (url === JWKS_URL) return new Response(JSON.stringify({ keys: [{ kty: 'RSA', kid: 'k1', ...K1.jwk, endorsements: ['msteams'] }] }), { status: 200 })
  return new Response('no', { status: 404 })
}) as unknown as typeof fetch

const message = (over: Record<string, unknown> = {}) => ({
  type: 'message',
  id: 'm-1',
  text: 'hello',
  channelId: 'msteams',
  serviceUrl: SURL,
  from: { id: '29:u', name: 'Op', aadObjectId: SENDER },
  recipient: { id: `28:${APP}`, name: 'Bot' },
  conversation: { id: 'a:dm-1', conversationType: 'personal', tenantId: TENANT },
  channelData: { tenant: { id: TENANT } },
  ...over,
})

function setup() {
  const dir = mkdtempSync(join(tmpdir(), 'cct-auth-e2e-'))
  const config = {
    stateDir: dir, allowlistFile: join(dir, 'allowlist.json'), appId: APP, appPassword: 'pw', appType: 'SingleTenant',
    tenantId: TENANT, bindHost: '127.0.0.1', port: 0, receivedFilesDir: join(dir, 'files'),
  } as any
  const adapter = createCloudAdapter(config)
  // No network: the outbound clients the SDK would build with an MSAL token for this turn.
  ;(adapter as any).createConnectorClientWithIdentity = async () => ({})
  ;(adapter as any).createUserTokenClient = async () => ({})
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(SENDER)
  const events: any[] = []
  const handler = makeTurnHandler({ config, adapter, allowlist, refs: createConversationRefStore(), onEvent: e => events.push(e) })
  const turns: any[] = []
  const turn = mock(async (ctx: any) => { turns.push({ identity: ctx.identity, activityId: ctx.activity.id }); await handler(ctx) })
  const inboundAuth = createChannelTokenVerifier(APP, { fetch: keysFetch, now: () => NOW })
  const post = (body: unknown, auth?: string) => handleMessagesPost(
    new Request('http://127.0.0.1:3979/api/messages', { method: 'POST', body: JSON.stringify(body), headers: { 'content-type': 'application/json', ...(auth ? { authorization: auth } : {}) } }),
    { adapter, inboundAuth, approvalCard: async () => ({}) },
    turn,
  )
  return { post, turn, turns, events, cleanup: () => rmSync(dir, { recursive: true, force: true }) }
}

describe('inbound auth end to end on the Agents SDK (TEAMSSDK1126)', () => {
  const refused: Array<[string, () => string | undefined]> = [
    ['no Authorization header', () => undefined],
    ['not a bearer token', () => `Basic ${Buffer.from('a:b').toString('base64')}`],
    ['signed by another key (bad signature)', () => `Bearer ${token({}, K2.privateKey)}`],
    ['another audience (another app id)', () => `Bearer ${token({ aud: '00000000-aaaa-bbbb-cccc-000000000009' })}`],
    ['another issuer (an Entra / emulator token)', () => `Bearer ${token({ iss: `https://sts.windows.net/${TENANT}/` })}`],
    ['expired', () => `Bearer ${token({ exp: Math.floor(NOW / 1000) - 3600 })}`],
    ['not yet valid', () => `Bearer ${token({ nbf: Math.floor(NOW / 1000) + 3600 })}`],
    ['serviceurl claim of another service', () => `Bearer ${token({ serviceurl: 'https://evil.example/' })}`],
  ]
  for (const [name, auth] of refused) {
    test(`401 and no turn: ${name}`, async () => {
      const s = setup()
      try {
        const res = await s.post(message(), auth())
        expect(res.status).toBe(401)
        expect(s.turn).not.toHaveBeenCalled()
        expect(s.events).toEqual([])
      } finally {
        s.cleanup()
      }
    })
  }

  test('a good token: delivered, and the turn identity is the verified claims (not anonymous)', async () => {
    const s = setup()
    try {
      const res = await s.post(message(), `Bearer ${token()}`)
      expect(res.status).toBe(200)
      expect(s.turn).toHaveBeenCalledTimes(1)
      expect(s.turns[0].identity?.aud).toBe(APP)
      expect(s.turns[0].identity?.iss).toBe(CHANNEL_TOKEN_ISSUER)
      expect(s.events.length).toBe(1)
    } finally {
      s.cleanup()
    }
  })

  test('a good token from another tenant: dropped by the tenant pin, nothing delivered', async () => {
    const s = setup()
    try {
      const res = await s.post(message({ conversation: { id: 'a:dm-1', conversationType: 'personal', tenantId: OTHER_TENANT }, channelData: { tenant: { id: OTHER_TENANT } } }), `Bearer ${token()}`)
      // The channel token carries no tenant: the request is valid, the handler runs, and its tenant pin drops it.
      expect(res.status).toBe(200)
      expect(s.turn).toHaveBeenCalledTimes(1)
      expect(s.events).toEqual([])
    } finally {
      s.cleanup()
    }
  })

  test('a good token but a malformed activity: 400, no turn', async () => {
    const s = setup()
    try {
      const res = await s.post(message({ conversation: 5 }), `Bearer ${token()}`)
      expect(res.status).toBe(400)
      expect(s.turn).not.toHaveBeenCalled()
    } finally {
      s.cleanup()
    }
  })

  test('a turn that throws is not answered in the chat: onTurnError hands the error back (no SDK default message)', async () => {
    const adapter = createCloudAdapter({ appId: APP, appPassword: 'pw', tenantId: TENANT, appType: 'SingleTenant' } as any)
    const sent: unknown[] = []
    const ctx = { sendActivity: async (a: unknown) => { sent.push(a) }, sendTraceActivity: async () => {} } as any
    const boom = new Error('boom')
    await expect(adapter.onTurnError!(ctx, boom)).rejects.toBe(boom)
    expect(sent).toEqual([])
  })

  test('no adapter without the app id, password and tenant (the SDK would run anonymous without a client id)', () => {
    for (const miss of ['appId', 'appPassword', 'tenantId']) {
      const config = { appId: APP, appPassword: 'pw', tenantId: TENANT, appType: 'SingleTenant', [miss]: '' } as any
      expect(() => createCloudAdapter(config)).toThrow('required')
    }
  })
})
