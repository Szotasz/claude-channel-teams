/**
 * Outbound mentions for `reply`. People must be members of the conversation (checked at send time); tags must be
 * listed in TEAMS_MENTION_TAGS and work only in a team channel. See docs/mentions.md.
 */

import { TeamsInfo, type TurnContext } from 'botbuilder'

export const MAX_MENTIONS = 10

const AAD_ID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/
const TAG_ID_RE = /^[A-Za-z0-9+/]{1,512}={0,2}$/

function isValidName(v: unknown): v is string {
  return (
    typeof v === 'string' &&
    v.length >= 1 &&
    v.trim() === v &&
    v.length <= 100 &&
    !/[<>,=\u0000-\u001f\u007f]/.test(v)
  )
}

export function parseMentionTags(raw: string | undefined): ReadonlyMap<string, string> | null {
  const tags = new Map<string, string>()
  if (raw === undefined || raw.trim() === '') return tags
  for (const part of raw.split(',')) {
    const entry = part.trim()
    if (entry === '') continue
    const eq = entry.indexOf('=')
    if (eq < 0) return null
    const name = entry.slice(0, eq).trim()
    const id = entry.slice(eq + 1).trim()
    if (!isValidName(name) || !TAG_ID_RE.test(id) || tags.has(name)) return null
    tags.set(name, id)
  }
  return tags
}

export type Mention =
  | { type: 'user'; aadObjectId: string; name: string }
  | { type: 'tag'; name: string; id: string }

export function validateMentions(
  raw: unknown,
  text: string,
  tags: ReadonlyMap<string, string>,
  inChannel: boolean,
): Mention[] {
  if (!Array.isArray(raw)) throw new Error('mentions must be an array')
  if (raw.length === 0) throw new Error('mentions must not be empty')
  if (raw.length > MAX_MENTIONS) throw new Error(`mentions can list at most ${MAX_MENTIONS}`)
  const out: Mention[] = []
  const users = new Set<string>()
  const seenTags = new Set<string>()
  const names = new Set<string>()
  raw.forEach((m: unknown, i) => {
    if (typeof m !== 'object' || m === null || Array.isArray(m)) {
      throw new Error(`mentions[${i}] must be an object`)
    }
    const e = m as Record<string, unknown>
    if (e.type !== 'user' && e.type !== 'tag') {
      throw new Error(`mentions[${i}].type must be "user" or "tag"`)
    }
    let mention: Mention
    if (e.type === 'user') {
      if (typeof e.aad_object_id !== 'string' || !AAD_ID_RE.test(e.aad_object_id)) {
        throw new Error(`mentions[${i}].aad_object_id must be an AAD object id`)
      }
      if (!isValidName(e.name)) {
        throw new Error(`mentions[${i}].name must be 1-100 characters without < > , =`)
      }
      const id = e.aad_object_id.toLowerCase()
      if (users.has(id)) throw new Error(`mentions[${i}] is listed twice`)
      users.add(id)
      mention = { type: 'user', aadObjectId: id, name: e.name }
    } else {
      if (tags.size === 0) throw new Error('tag mentions are not enabled (set TEAMS_MENTION_TAGS)')
      if (typeof e.tag !== 'string' || !tags.has(e.tag)) {
        throw new Error(`mentions[${i}].tag is not in TEAMS_MENTION_TAGS`)
      }
      if (!inChannel) throw new Error('tag mentions work only in a team channel')
      if (seenTags.has(e.tag)) throw new Error(`mentions[${i}] is listed twice`)
      seenTags.add(e.tag)
      mention = { type: 'tag', name: e.tag, id: tags.get(e.tag)! }
    }
    if (names.has(mention.name)) throw new Error(`mentions[${i}] uses the same name as an earlier entry`)
    names.add(mention.name)
    if (!text.includes(`<at>${mention.name}</at>`)) {
      throw new Error(`mentions[${i}]: the text must contain <at>${mention.name}</at>`)
    }
    out.push(mention)
  })
  return out
}

export function mentionEntity(
  name: string,
  id: string,
  tag: boolean,
): { type: 'mention'; text: string; mentioned: { id: string; name: string; type?: 'tag' } } {
  return {
    type: 'mention',
    text: `<at>${name}</at>`,
    mentioned: { id, name, ...(tag ? { type: 'tag' as const } : {}) },
  }
}

export function isChannelReference(
  ref: { conversation?: { id?: string; conversationType?: string } } | undefined,
): boolean {
  const c = ref?.conversation
  if (c?.conversationType === 'channel') return true
  return typeof c?.id === 'string' && /@thread\.tacv2/.test(c.id)
}

/** Resolves to the member's id, or undefined when the person is not a member. */
export type MemberLookup = (ctx: TurnContext, aadObjectId: string) => Promise<string | undefined>

export function isNotMemberError(err: unknown): boolean {
  if (typeof err !== 'object' || err === null) return false
  const e = err as { statusCode?: unknown; code?: unknown; body?: { error?: { code?: unknown } } }
  return (
    e.statusCode === 404 ||
    e.code === 'MemberNotFoundInConversation' ||
    e.body?.error?.code === 'MemberNotFoundInConversation'
  )
}

export function createMemberLookup(
  getMember: (ctx: TurnContext, userId: string) => Promise<{ id?: string } | undefined>,
): MemberLookup {
  return async (ctx, aadObjectId) => {
    let member
    try {
      member = await getMember(ctx, aadObjectId)
    } catch (err) {
      if (isNotMemberError(err)) return undefined
      throw err
    }
    return typeof member?.id === 'string' && member.id.length > 0 ? member.id : undefined
  }
}

export const teamsMemberLookup: MemberLookup = createMemberLookup((ctx, id) => TeamsInfo.getMember(ctx, id))
