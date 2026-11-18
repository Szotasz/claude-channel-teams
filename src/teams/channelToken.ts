/**
 * Channel-token-only check for inbound Bot Framework requests (TEAMSEMUTOKEN1118).
 *
 * Why this exists, ahead of CloudAdapter: botframework-connector (4.23.3, parameterizedBotFrameworkAuthentication.ts
 * JwtTokenValidation_authenticateToken) sends every token that EmulatorValidation.isTokenFromEmulator accepts down the
 * emulator branch, and that branch only requires the token's appid/azp claim to be the bot's own app id. Whoever holds
 * the bot's app password (it sits in the channel .env, readable by every process of the same user) can request such an
 * AAD token and post a forged activity: a message "from" the paired owner, a card action, a permission answer.
 *
 * So the plugin admits ONLY what the Bot Framework channel service signs, the rules of the SDK's own channel branch
 * (ChannelValidation_authenticateChannelToken + governmentChannelValidation_ValidateIdentity + JwtTokenExtractor),
 * narrowed:
 *   - "Bearer <jwt>", alg exactly RS256, a kid;
 *   - iss exactly https://api.botframework.com;
 *   - aud exactly the bot app id;
 *   - exp / nbf with CLOCK_SKEW_S of tolerance (the SDK allows 5 minutes);
 *   - the serviceurl claim equal to the activity's serviceUrl (the activity must carry one);
 *   - the signing key from the Bot Framework OpenID metadata's JWKS, an RSA key whose endorsements include the
 *     channel, and the channel is msteams (the SDK skips the endorsement check for a key with no endorsements; here a
 *     key without "msteams" is refused).
 * Emulator and skill tokens have another issuer and are refused by the iss rule before any key is fetched.
 *
 * Node's crypto only (createPublicKey from the JWK, RSA-SHA256 verify); the metadata fetch and the clock are injectable.
 * The same logic lives in Marveen's dashboard (bot-channel-token.ts) for the approval card; keep the two in step.
 */
import { createPublicKey, verify as cryptoVerify, type KeyObject } from 'node:crypto'

export const CHANNEL_TOKEN_ISSUER = 'https://api.botframework.com'
export const CHANNEL_OPENID_METADATA_URL = 'https://login.botframework.com/v1/.well-known/openidconfiguration'
export const REQUIRED_CHANNEL = 'msteams'
export const CLOCK_SKEW_S = 300
/** Keys are kept this long; an unknown kid, and a fetch after a failed one, refetch at most once per REFETCH_MIN_MS. */
export const KEYS_TTL_MS = 24 * 60 * 60 * 1000
export const REFETCH_MIN_MS = 5 * 60 * 1000

export type ChannelTokenResult =
  | { ok: true; claims: Record<string, unknown> }
  | { ok: false; reason: string }

export interface ChannelTokenDeps {
  fetch?: typeof fetch
  now?: () => number
  metadataUrl?: string
}

interface CachedKey { key: KeyObject; endorsements: string[] }

function b64urlJson(part: string): Record<string, unknown> | null {
  try {
    const v = JSON.parse(Buffer.from(part, 'base64url').toString('utf8')) as unknown
    return v && typeof v === 'object' && !Array.isArray(v) ? (v as Record<string, unknown>) : null
  } catch {
    return null
  }
}

/**
 * A verifier bound to one bot app id, with its own key cache. `verify(authHeader, activity)` never throws: a refusal
 * is a fixed reason, never the token.
 */
export function createChannelTokenVerifier(appId: string, deps: ChannelTokenDeps = {}) {
  const doFetch = deps.fetch ?? fetch
  const now = deps.now ?? Date.now
  const metadataUrl = deps.metadataUrl ?? CHANNEL_OPENID_METADATA_URL
  let keys = new Map<string, CachedKey>()
  let fetchedAt = -Infinity
  // R1 (Dani, #55 review): the time of the last FAILED fetch. Without it, every request that passes the cheap claims
  // during an outage would start a new outbound fetch.
  let failedAt = -Infinity

  async function loadKeys(): Promise<void> {
    const metaRes = await doFetch(metadataUrl)
    if (!metaRes.ok) throw new Error(`metadata HTTP ${metaRes.status}`)
    const meta = (await metaRes.json()) as { jwks_uri?: unknown }
    if (typeof meta.jwks_uri !== 'string') throw new Error('metadata has no jwks_uri')
    const jwksUrl = new URL(meta.jwks_uri)
    if (jwksUrl.protocol !== 'https:' || jwksUrl.host !== new URL(metadataUrl).host) throw new Error('jwks_uri is not on the metadata host')
    const jwksRes = await doFetch(jwksUrl.toString())
    if (!jwksRes.ok) throw new Error(`jwks HTTP ${jwksRes.status}`)
    const jwks = (await jwksRes.json()) as { keys?: unknown }
    const next = new Map<string, CachedKey>()
    for (const k of Array.isArray(jwks.keys) ? jwks.keys : []) {
      const jwk = k as Record<string, unknown>
      if (jwk.kty !== 'RSA' || typeof jwk.kid !== 'string' || typeof jwk.n !== 'string' || typeof jwk.e !== 'string') continue
      const endorsements = Array.isArray(jwk.endorsements) ? jwk.endorsements.filter((x): x is string => typeof x === 'string') : []
      try {
        next.set(jwk.kid, { key: createPublicKey({ key: { kty: 'RSA', n: jwk.n, e: jwk.e }, format: 'jwk' }), endorsements })
      } catch { /* an unusable key is skipped */ }
    }
    keys = next
    fetchedAt = now()
  }

  async function keyFor(kid: string): Promise<CachedKey | undefined> {
    const age = now() - fetchedAt
    const stale = age > KEYS_TTL_MS
    if (stale || (!keys.has(kid) && age > REFETCH_MIN_MS)) {
      if (now() - failedAt < REFETCH_MIN_MS) {
        if (stale) throw new Error('signing keys unavailable (a recent fetch failed)')
        return keys.get(kid)
      }
      try {
        await loadKeys()
      } catch (err) {
        failedAt = now()
        throw err
      }
    }
    return keys.get(kid)
  }

  return async function verifyChannelToken(authHeader: string | undefined, activity: unknown): Promise<ChannelTokenResult> {
    if (typeof authHeader !== 'string' || !authHeader.startsWith('Bearer ')) return { ok: false, reason: 'no bearer token' }
    const parts = authHeader.slice('Bearer '.length).trim().split('.')
    if (parts.length !== 3 || parts.some((p) => !/^[A-Za-z0-9_-]+$/.test(p))) return { ok: false, reason: 'malformed token' }
    const [h, p, s] = parts as [string, string, string]
    const header = b64urlJson(h)
    const claims = b64urlJson(p)
    if (!header || !claims) return { ok: false, reason: 'malformed token' }
    if (header.alg !== 'RS256') return { ok: false, reason: 'algorithm is not RS256' }
    if (typeof header.kid !== 'string' || header.kid === '') return { ok: false, reason: 'no key id' }
    if (claims.iss !== CHANNEL_TOKEN_ISSUER) return { ok: false, reason: 'not a Bot Framework channel token (issuer)' }
    if (typeof claims.aud !== 'string' || claims.aud !== appId) return { ok: false, reason: 'audience is not this bot' }
    const t = Math.floor(now() / 1000)
    if (typeof claims.exp !== 'number' || claims.exp + CLOCK_SKEW_S < t) return { ok: false, reason: 'token expired' }
    if (claims.nbf !== undefined && (typeof claims.nbf !== 'number' || claims.nbf - CLOCK_SKEW_S > t)) return { ok: false, reason: 'token not yet valid' }
    const act = activity && typeof activity === 'object' ? (activity as { serviceUrl?: unknown; channelId?: unknown }) : {}
    if (typeof act.serviceUrl !== 'string' || act.serviceUrl === '' || claims.serviceurl !== act.serviceUrl) {
      return { ok: false, reason: 'serviceUrl claim does not match the activity' }
    }
    if (act.channelId !== REQUIRED_CHANNEL) return { ok: false, reason: 'activity is not from the msteams channel' }
    let key: CachedKey | undefined
    try {
      key = await keyFor(header.kid)
    } catch {
      return { ok: false, reason: 'signing keys unavailable' }
    }
    if (!key) return { ok: false, reason: 'unknown signing key' }
    if (!key.endorsements.includes(REQUIRED_CHANNEL)) return { ok: false, reason: 'signing key is not endorsed for msteams' }
    let good = false
    try {
      good = cryptoVerify('RSA-SHA256', Buffer.from(`${h}.${p}`), key.key, Buffer.from(s, 'base64url'))
    } catch {
      good = false
    }
    if (!good) return { ok: false, reason: 'bad signature' }
    return { ok: true, claims }
  }
}

export type ChannelTokenVerifier = ReturnType<typeof createChannelTokenVerifier>
