/**
 * TEAMSEMUTOKEN1118: only Bot Framework channel tokens reach CloudAdapter. Test vectors are signed with a key made
 * here; the OpenID metadata and the JWKS are served by an injected fetch, so nothing leaves the process. Every rule of
 * the verifier has a case that breaks only that rule.
 */
import { describe, expect, test, mock } from 'bun:test'
import { generateKeyPairSync, sign as cryptoSign, type KeyObject } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  createChannelTokenVerifier,
  CHANNEL_OPENID_METADATA_URL,
  CHANNEL_TOKEN_ISSUER,
  CLOCK_SKEW_S,
  REFETCH_MIN_MS,
} from '../src/teams/channelToken.ts'
import { handleMessagesPost } from '../src/teams/adapter.ts'

const APP = '00000000-aaaa-bbbb-cccc-000000000001'
const SURL = 'https://smba.trafficmanager.net/emea/'
const JWKS_URL = 'https://login.botframework.com/v1/.well-known/keys'
const NOW = 1_795_000_000_000

function rsa() {
  const { privateKey, publicKey } = generateKeyPairSync('rsa', { modulusLength: 2048 })
  return { privateKey, jwk: publicKey.export({ format: 'jwk' }) as { n: string; e: string } }
}
const K1 = rsa()
const K2 = rsa()

function b64(o: unknown): string { return Buffer.from(JSON.stringify(o)).toString('base64url') }
function token(over: Record<string, unknown> = {}, head: Record<string, unknown> = {}, key: KeyObject = K1.privateKey): string {
  const t = Math.floor(NOW / 1000)
  const h = b64({ alg: 'RS256', kid: 'k1', typ: 'JWT', ...head })
  const p = b64({ iss: CHANNEL_TOKEN_ISSUER, aud: APP, exp: t + 3600, nbf: t - 10, serviceurl: SURL, ...over })
  const s = cryptoSign('RSA-SHA256', Buffer.from(`${h}.${p}`), key).toString('base64url')
  return `${h}.${p}.${s}`
}
const activity = (over: Record<string, unknown> = {}) => ({ type: 'message', channelId: 'msteams', serviceUrl: SURL, ...over })

function keysFetch(keys: unknown[] = [{ kty: 'RSA', kid: 'k1', ...K1.jwk, endorsements: ['msteams', 'webchat'] }], jwksUri = JWKS_URL) {
  const urls: string[] = []
  const f = mock(async (input: unknown) => {
    const url = String(input)
    urls.push(url)
    if (url === CHANNEL_OPENID_METADATA_URL) return new Response(JSON.stringify({ jwks_uri: jwksUri }), { status: 200 })
    if (url === jwksUri) return new Response(JSON.stringify({ keys }), { status: 200 })
    return new Response('no', { status: 404 })
  })
  return { f: f as unknown as typeof fetch, urls }
}
function verifier(keys?: unknown[], jwksUri?: string, now: () => number = () => NOW) {
  const k = keysFetch(keys, jwksUri)
  return { verify: createChannelTokenVerifier(APP, { fetch: k.f, now }), urls: k.urls }
}

describe('channel-token-only verifier', () => {
  test('a channel token signed by an msteams-endorsed key passes', async () => {
    const { verify } = verifier()
    const r = await verify(`Bearer ${token()}`, activity())
    expect(r.ok).toBe(true)
  })

  test('an emulator-shaped token for the bot itself (the bot password path) is refused before any key fetch', async () => {
    const { verify, urls } = verifier()
    const t = Math.floor(NOW / 1000)
    const emu = token({ iss: 'https://sts.windows.net/11111111-2222-3333-4444-555555555555/', aud: APP, appid: APP, ver: '1.0', exp: t + 3600 })
    expect(await verify(`Bearer ${emu}`, activity())).toEqual({ ok: false, reason: 'not a Bot Framework channel token (issuer)' })
    const v2 = token({ iss: 'https://login.microsoftonline.com/11111111-2222-3333-4444-555555555555/v2.0', aud: APP, azp: APP, ver: '2.0' })
    expect((await verify(`Bearer ${v2}`, activity())).ok).toBe(false)
    expect(urls).toEqual([])
  })

  test('header shape: no bearer, malformed, other algorithms, no kid', async () => {
    const { verify } = verifier()
    expect(await verify(undefined, activity())).toEqual({ ok: false, reason: 'no bearer token' })
    expect(await verify(token(), activity())).toEqual({ ok: false, reason: 'no bearer token' })
    expect(await verify('Bearer a.b', activity())).toEqual({ ok: false, reason: 'malformed token' })
    expect(await verify('Bearer a.b.c+', activity())).toEqual({ ok: false, reason: 'malformed token' })
    for (const alg of ['none', 'HS256', 'RS384', 'PS256']) {
      expect(await verify(`Bearer ${token({}, { alg })}`, activity())).toEqual({ ok: false, reason: 'algorithm is not RS256' })
    }
    expect(await verify(`Bearer ${token({}, { kid: undefined })}`, activity())).toEqual({ ok: false, reason: 'no key id' })
  })

  test('audience: another app, an array, missing', async () => {
    const { verify } = verifier()
    for (const aud of ['another-app', [APP], undefined]) {
      expect(await verify(`Bearer ${token({ aud })}`, activity())).toEqual({ ok: false, reason: 'audience is not this bot' })
    }
  })

  // The skew is pinned as a number, 300 s (the SDK's 5 minutes), not read back from the constant it tests.
  test('time: expired beyond the 300 s skew is refused, inside it passes; nbf beyond the skew is refused', async () => {
    expect(CLOCK_SKEW_S).toBe(300)
    const { verify } = verifier()
    const t = Math.floor(NOW / 1000)
    expect(await verify(`Bearer ${token({ exp: t - 301 })}`, activity())).toEqual({ ok: false, reason: 'token expired' })
    expect((await verify(`Bearer ${token({ exp: t - 295 })}`, activity())).ok).toBe(true)
    expect(await verify(`Bearer ${token({ exp: undefined })}`, activity())).toEqual({ ok: false, reason: 'token expired' })
    expect(await verify(`Bearer ${token({ nbf: t + 305 })}`, activity())).toEqual({ ok: false, reason: 'token not yet valid' })
    expect((await verify(`Bearer ${token({ nbf: t + 295 })}`, activity())).ok).toBe(true)
  })

  test('serviceUrl: the claim must equal the activity serviceUrl, and the activity must carry one', async () => {
    const { verify } = verifier()
    const reason = 'serviceUrl claim does not match the activity'
    expect(await verify(`Bearer ${token({ serviceurl: 'https://evil.example/' })}`, activity())).toEqual({ ok: false, reason })
    expect(await verify(`Bearer ${token()}`, activity({ serviceUrl: undefined }))).toEqual({ ok: false, reason })
    expect(await verify(`Bearer ${token({ serviceurl: undefined })}`, activity())).toEqual({ ok: false, reason })
  })

  test('channel: only msteams activities', async () => {
    const { verify } = verifier()
    expect(await verify(`Bearer ${token()}`, activity({ channelId: 'emulator' }))).toEqual({ ok: false, reason: 'activity is not from the msteams channel' })
  })

  test('keys: unknown kid, a key not endorsed for msteams, a key with no endorsements', async () => {
    expect(await verifier().verify(`Bearer ${token({}, { kid: 'k9' })}`, activity())).toEqual({ ok: false, reason: 'unknown signing key' })
    const notTeams = verifier([{ kty: 'RSA', kid: 'k1', ...K1.jwk, endorsements: ['webchat'] }])
    expect(await notTeams.verify(`Bearer ${token()}`, activity())).toEqual({ ok: false, reason: 'signing key is not endorsed for msteams' })
    const none = verifier([{ kty: 'RSA', kid: 'k1', ...K1.jwk }])
    expect(await none.verify(`Bearer ${token()}`, activity())).toEqual({ ok: false, reason: 'signing key is not endorsed for msteams' })
  })

  test('signature: another key under the same kid, a payload changed after signing', async () => {
    const { verify } = verifier()
    expect(await verify(`Bearer ${token({}, {}, K2.privateKey)}`, activity())).toEqual({ ok: false, reason: 'bad signature' })
    const [h, , s] = token().split('.')
    const forged = `${h}.${b64({ iss: CHANNEL_TOKEN_ISSUER, aud: APP, exp: Math.floor(NOW / 1000) + 3600, serviceurl: SURL, extra: 1 })}.${s}`
    expect(await verify(`Bearer ${forged}`, activity())).toEqual({ ok: false, reason: 'bad signature' })
  })

  test('metadata: unavailable, or a jwks_uri off the metadata host, gives no keys', async () => {
    const down = createChannelTokenVerifier(APP, { fetch: (async () => new Response('x', { status: 503 })) as unknown as typeof fetch, now: () => NOW })
    expect(await down(`Bearer ${token()}`, activity())).toEqual({ ok: false, reason: 'signing keys unavailable' })
    const off = verifier(undefined, 'https://evil.example/keys')
    expect(await off.verify(`Bearer ${token()}`, activity())).toEqual({ ok: false, reason: 'signing keys unavailable' })
  })

  test('cache: keys are fetched once; an unknown kid refetches only after REFETCH_MIN_MS', async () => {
    let now = NOW
    const { verify, urls } = verifier(undefined, undefined, () => now)
    expect((await verify(`Bearer ${token()}`, activity())).ok).toBe(true)
    expect((await verify(`Bearer ${token()}`, activity())).ok).toBe(true)
    expect(urls.length).toBe(2)
    await verify(`Bearer ${token({}, { kid: 'k9' })}`, activity())
    expect(urls.length).toBe(2)
    now += REFETCH_MIN_MS + 1
    await verify(`Bearer ${token({ exp: Math.floor(now / 1000) + 3600 }, { kid: 'k9' })}`, activity())
    expect(urls.length).toBe(4)
  })
})

describe('POST /api/messages: the verifier runs before CloudAdapter', () => {
  const body = JSON.stringify(activity())
  function req(auth?: string) {
    return new Request('http://127.0.0.1:3979/api/messages', { method: 'POST', body, headers: { 'content-type': 'application/json', ...(auth ? { authorization: auth } : {}) } })
  }
  test('a refused token is a 401 and the adapter is never called', async () => {
    const process = mock(async () => {})
    const seen: Array<string | undefined> = []
    const inboundAuth = async (h: string | undefined) => { seen.push(h); return { ok: false as const, reason: 'x' } }
    const res = await handleMessagesPost(req('Bearer emu.lated.token'), { adapter: { process } as never, inboundAuth }, async () => {})
    expect(res.status).toBe(401)
    expect(process).not.toHaveBeenCalled()
    expect(seen).toEqual(['Bearer emu.lated.token'])
  })
  test('an accepted token goes on to the adapter', async () => {
    const process = mock(async (_req: unknown, res: { status(n: number): void; end(b?: string): void }) => { res.status(200); res.end() })
    const inboundAuth = async () => ({ ok: true as const, claims: {} })
    const res = await handleMessagesPost(req('Bearer a.b.c'), { adapter: { process } as never, inboundAuth }, async () => {})
    expect(process).toHaveBeenCalledTimes(1)
    expect(res.status).toBe(200)
  })
})

describe('server.ts wiring', () => {
  const server = readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8')
  test('the listener gets the channel-token verifier for this bot app id', () => {
    const start = server.indexOf('const listener = startListener({')
    expect(start).toBeGreaterThan(-1)
    const end = server.indexOf('})', start)
    expect(server.slice(start, end)).toContain('inboundAuth: createChannelTokenVerifier(config.appId),')
  })
})
