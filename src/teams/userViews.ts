/**
 * User-specific card views: the reply tool can list people who each see their own Adaptive Card instead of the base
 * card. Teams does this with the card's `refresh` property (Universal Actions): the client of each person in
 * `refresh.userIds` sends an `adaptiveCard/action` invoke, and the bot answers it with the card that person should see.
 *
 * This module validates the `user_cards` argument, adds the refresh to the cards (after validation: cards.ts refuses a
 * refresh in a card Claude sends), and keeps the views and the person -> user MRI map.
 * Kept in memory only: after a restart the stored views are gone, a refresh for an unknown view is answered 'Ignored.'
 * and the client keeps showing the base card.
 */

import { validateAdaptiveCard } from './cards.js'

export const USER_VIEW_VERB = 'teams-plugin-user-view'
export const MAX_USER_CARDS = 60
export const MAX_STORED_VIEWS = 500
export const MAX_KNOWN_USERS = 1000

export interface UserCard {
  aadObjectId: string
  card: Record<string, unknown>
}

const AAD_ID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

/** True when any key named `verb`, at any depth, has the reserved verb as its value. */
export function usesReservedVerb(root: unknown): boolean {
  const stack: unknown[] = [root]
  while (stack.length > 0) {
    const v = stack.pop()
    if (Array.isArray(v)) {
      stack.push(...v)
    } else if (isPlainObject(v)) {
      for (const [key, child] of Object.entries(v)) {
        if (key === 'verb' && child === USER_VIEW_VERB) return true
        stack.push(child)
      }
    }
  }
  return false
}

const versionOk = (card: unknown) => {
  const v = (card as { version?: unknown } | null | undefined)?.version
  return v === '1.4' || v === '1.5'
}

export function validateUserCards(raw: unknown, baseCard: unknown): UserCard[] {
  if (!Array.isArray(raw)) throw new Error('user_cards must be an array')
  if (raw.length === 0) throw new Error('user_cards must not be empty')
  if (raw.length > MAX_USER_CARDS) throw new Error(`user_cards can list at most ${MAX_USER_CARDS} people`)
  if (baseCard === undefined) throw new Error('user_cards needs card (the card everyone else sees)')
  if (!versionOk(baseCard)) throw new Error('card.version must be 1.4 or 1.5 when user_cards is set')
  if (usesReservedVerb(baseCard)) throw new Error(`card uses the reserved verb ${USER_VIEW_VERB}`)
  const seen = new Set<string>()
  const out: UserCard[] = []
  raw.forEach((entry, i) => {
    if (!isPlainObject(entry)) throw new Error(`user_cards[${i}] must be an object`)
    const id = entry.aad_object_id
    if (typeof id !== 'string' || !AAD_ID_RE.test(id)) {
      throw new Error(`user_cards[${i}].aad_object_id must be an AAD object id`)
    }
    const lower = id.toLowerCase()
    if (seen.has(lower)) throw new Error(`user_cards[${i}].aad_object_id is listed twice`)
    seen.add(lower)
    const checked = validateAdaptiveCard(entry.card)
    if (!checked.ok) throw new Error(`user_cards[${i}].card is invalid: ${checked.error}`)
    if (!versionOk(entry.card)) throw new Error(`user_cards[${i}].card.version must be 1.4 or 1.5`)
    if (usesReservedVerb(entry.card)) throw new Error(`user_cards[${i}].card uses the reserved verb ${USER_VIEW_VERB}`)
    out.push({ aadObjectId: lower, card: entry.card as Record<string, unknown> })
  })
  return out
}

export function withRefresh(
  card: Record<string, unknown>,
  viewId: string,
  userMris: string[],
): Record<string, unknown> {
  return {
    ...card,
    refresh: {
      action: { type: 'Action.Execute', title: 'Refresh', verb: USER_VIEW_VERB, data: { view: viewId } },
      userIds: [...userMris],
    },
  }
}

export function createKnownUsers(max = MAX_KNOWN_USERS) {
  const map = new Map<string, string>()
  return {
    set(aadObjectId: string, mri: string): void {
      if (!mri) return
      const key = aadObjectId.toLowerCase()
      map.delete(key)
      if (map.size >= max) {
        const oldest = map.keys().next()
        if (!oldest.done) map.delete(oldest.value)
      }
      map.set(key, mri)
    },
    get(aadObjectId: string): string | undefined {
      return map.get(aadObjectId.toLowerCase())
    },
    size(): number {
      return map.size
    },
  }
}

export function createUserViewStore(max = MAX_STORED_VIEWS, newId: () => string = () => crypto.randomUUID()) {
  const views = new Map<string, { base: Record<string, unknown>; cards: Map<string, Record<string, unknown>> }>()
  return {
    create(
      baseCard: Record<string, unknown>,
      userCards: UserCard[],
      userMris: string[],
    ): { id: string; base: Record<string, unknown> } {
      const id = newId()
      const base = withRefresh(baseCard, id, userMris)
      const cards = new Map<string, Record<string, unknown>>()
      for (const u of userCards) cards.set(u.aadObjectId, withRefresh(u.card, id, userMris))
      views.set(id, { base, cards })
      while (views.size > max) {
        const oldest = views.keys().next()
        if (oldest.done) break
        views.delete(oldest.value)
      }
      return { id, base }
    },
    cardFor(viewId: string, aadObjectId: string): Record<string, unknown> | undefined {
      const view = views.get(viewId)
      if (!view) return undefined
      return view.cards.get(aadObjectId.toLowerCase()) ?? view.base
    },
    size(): number {
      return views.size
    },
  }
}

export function userViewInvokeFrom(activity: {
  type?: string
  name?: string
  value?: unknown
}): { viewId: string | undefined } | undefined {
  if (activity.type !== 'invoke' || activity.name !== 'adaptiveCard/action') return undefined
  if (!isPlainObject(activity.value)) return undefined
  const action = activity.value.action
  if (!isPlainObject(action) || action.verb !== USER_VIEW_VERB) return undefined
  const view = isPlainObject(action.data) ? action.data.view : undefined
  return { viewId: typeof view === 'string' && view !== '' ? view : undefined }
}
