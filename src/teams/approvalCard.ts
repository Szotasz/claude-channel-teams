/**
 * Marveen approval cards (TEAMSAPPROVECARD1118): the plugin's part.
 *
 * The Marveen dashboard posts an approval card itself (an email the owner approves or rejects), with Action.Execute
 * buttons whose verb is APPROVAL_CARD_VERB. A click arrives here as an `adaptiveCard/action` invoke. The plugin does NOT
 * decide it, and it does NOT hand it to the Claude session as a "[Card action]": the agent whose email it is must not be
 * on the path. The plugin forwards the click, with its original Authorization header, to the dashboard
 * (MARVEEN_APPROVAL_URL, loopback only). The dashboard checks the Bot Framework channel token itself, then the tenant, the
 * owner, the posted card and its one-time nonce, and decides. The plugin returns the card the dashboard answers with as
 * the invoke response, so Teams shows it in place of the clicked card.
 *
 * Before forwarding, the plugin applies its own gates as well: the bot's tenant on the conversation, and the clicker on
 * the allowlist. Anything that goes wrong (no URL configured, a gate, the dashboard unreachable or refusing, an answer
 * that is not a valid card) is answered with a fixed "decide on the dashboard" card. A model-supplied card cannot carry
 * the verb (cards.ts refuses it), so the agent cannot make a button that would be taken for one of these.
 */
import { validateAdaptiveCard, APPROVAL_CARD_VERB } from './cards.js'

export { APPROVAL_CARD_VERB }

export const FORWARD_TIMEOUT_MS = 8_000

type Card = Record<string, unknown>

/** The fixed card for every case the dashboard did not decide. Both languages: the plugin does not know the install's. */
export const NOT_DECIDED_CARD: Card = {
  type: 'AdaptiveCard',
  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
  version: '1.4',
  body: [
    { type: 'TextBlock', text: 'Itt most nem dönthető el; a dashboardon dönthetsz.', wrap: true },
    { type: 'TextBlock', text: 'This cannot be decided here right now; you can decide on the dashboard.', wrap: true, isSubtle: true },
  ],
}

/** True for an Action.Execute invoke whose verb is the approval verb, the only shape the dashboard's card produces. */
export function isApprovalCardInvoke(activity: unknown): boolean {
  if (!activity || typeof activity !== 'object') return false
  const a = activity as { type?: unknown; name?: unknown; value?: { action?: { verb?: unknown } } }
  return a.type === 'invoke' && a.name === 'adaptiveCard/action' && a.value?.action?.verb === APPROVAL_CARD_VERB
}

/**
 * Dani B1 (#57 review): a client that sends the approval button as an Action.Submit fallback would post a MESSAGE whose
 * value is the button's data ({ approvalId, decision, nonce }), and cardAction.ts hands any message with an object value
 * to the session as a submit, nonce included. Such a message is dropped before CloudAdapter: never forwarded (the
 * dashboard takes only the Action.Execute invoke), never shown to the agent. A message answer cannot carry a card, so the
 * clicker sees nothing here; the owner can still decide on the card's other path or on the dashboard.
 */
export function isApprovalCardSubmit(activity: unknown): boolean {
  if (!activity || typeof activity !== 'object') return false
  const a = activity as { type?: unknown; value?: unknown }
  if (a.type !== 'message' || !a.value || typeof a.value !== 'object' || Array.isArray(a.value)) return false
  const v = a.value as { verb?: unknown; action?: { verb?: unknown }; approvalId?: unknown; nonce?: unknown }
  if (v.verb === APPROVAL_CARD_VERB || v.action?.verb === APPROVAL_CARD_VERB) return true
  return v.approvalId !== undefined && v.nonce !== undefined
}

/**
 * MARVEEN_APPROVAL_URL, only when it is an http(s) URL on the loopback interface (127.0.0.1, ::1, localhost). The click
 * carries a Bot Framework token that is valid for a while; it must not leave the host.
 */
export function parseApprovalUrl(raw: string | undefined): string | undefined {
  const v = (raw ?? '').trim()
  if (!v) return undefined
  let u: URL
  try { u = new URL(v) } catch { return undefined }
  if (u.protocol !== 'http:' && u.protocol !== 'https:') return undefined
  if (!['127.0.0.1', '[::1]', 'localhost'].includes(u.hostname)) return undefined
  if (u.username || u.password) return undefined
  return u.toString()
}

export interface ApprovalForwardDeps {
  /** The dashboard route, already checked by parseApprovalUrl; undefined = not wired. */
  url: string | undefined
  /** The bot's tenant (TEAMS_BOT_TENANT_ID). */
  tenantId: string
  isAllowed: (aadObjectId: string) => boolean
  fetch?: typeof fetch
  log?: (line: string) => void
}

/**
 * One approval-card click -> the card to show in its place. Never throws, never reaches the agent.
 */
export function createApprovalForwarder(deps: ApprovalForwardDeps) {
  const log = deps.log ?? ((line: string) => { process.stderr.write(line) })
  const doFetch = deps.fetch ?? fetch
  return async function forwardApprovalClick(authHeader: string | undefined, activity: unknown): Promise<Card> {
    const a = activity as { from?: { aadObjectId?: unknown }; conversation?: { tenantId?: unknown } }
    if (!deps.url) {
      log('teams channel: approval card click not forwarded (MARVEEN_APPROVAL_URL is not set to a loopback URL)\n')
      return NOT_DECIDED_CARD
    }
    const tenant = typeof a.conversation?.tenantId === 'string' ? a.conversation.tenantId.toLowerCase() : ''
    if (!deps.tenantId || tenant !== deps.tenantId.toLowerCase()) {
      log('teams channel: approval card click not forwarded (tenant mismatch)\n')
      return NOT_DECIDED_CARD
    }
    const from = typeof a.from?.aadObjectId === 'string' ? a.from.aadObjectId : ''
    if (!from || !deps.isAllowed(from)) {
      log('teams channel: approval card click not forwarded (sender not on the allowlist)\n')
      return NOT_DECIDED_CARD
    }
    let res: Response
    try {
      res = await doFetch(deps.url, {
        method: 'POST',
        headers: { 'content-type': 'application/json', ...(authHeader ? { authorization: authHeader } : {}) },
        body: JSON.stringify(activity),
        redirect: 'error',
        signal: AbortSignal.timeout(FORWARD_TIMEOUT_MS),
      })
    } catch {
      log('teams channel: approval card click: the dashboard did not answer\n')
      return NOT_DECIDED_CARD
    }
    let body: unknown
    try { body = await res.json() } catch { body = undefined }
    const card = body && typeof body === 'object' ? (body as { card?: unknown }).card : undefined
    if (card !== undefined && validateAdaptiveCard(card).ok) {
      if (!res.ok) log(`teams channel: approval card click: the dashboard answered ${res.status}\n`)
      return card as Card
    }
    log(`teams channel: approval card click: the dashboard answered ${res.status} without a usable card\n`)
    return NOT_DECIDED_CARD
  }
}

export type ApprovalForwarder = ReturnType<typeof createApprovalForwarder>
