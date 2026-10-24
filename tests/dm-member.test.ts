import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import type { ConversationReference } from 'botbuilder'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore, isPersonalReference } from '../src/teams/conversationRefs'
import {
  BotBlockedError,
  NoServiceUrlError,
  PersonalInstallMissingError,
  RecipientNotAllowedError,
  createDirectMessenger,
} from '../src/teams/directMessage'

// Letters in the GUID so a case variant genuinely differs.
const USER_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
const OTHER_ID = '11111111-2222-4333-8444-555555555555'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const SERVICE_URL = 'https://smba.trafficmanager.net/amer/'

/** A 1:1 chat reference as the inbound path would store one. */
function personalRef(convId: string, aadId: string, serviceUrl = SERVICE_URL): Partial<ConversationReference> {
  return {
    channelId: 'msteams',
    serviceUrl,
    conversation: { id: convId, conversationType: 'personal' as const, isGroup: false },
    user: { id: `u:${aadId}`, aadObjectId: aadId },
  } as Partial<ConversationReference>
}

/** A team-channel reference — same person, but not a 1:1 chat. */
function channelRef(convId: string, aadId: string, serviceUrl = SERVICE_URL): Partial<ConversationReference> {
  return {
    channelId: 'msteams',
    serviceUrl,
    conversation: { id: convId, conversationType: 'channel' as const, isGroup: true },
    user: { id: `u:${aadId}`, aadObjectId: aadId },
  } as Partial<ConversationReference>
}

/** Resolve to the rejection (or undefined when the call unexpectedly resolves). */
async function errOf(p: Promise<unknown>): Promise<unknown> {
  try {
    await p
  } catch (e) {
    return e
  }
  return undefined
}

describe('dm_member', () => {
  let dir: string
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cct-dm-'))
  })
  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  /**
   * Real allowlist (file-backed) and real ref store (in memory); only the
   * adapter and the replier are fakes. The adapter records the
   * createConversationAsync arguments and runs the callback with a context
   * whose conversation id is 'a:new-1' and whose sendActivity returns 'm-1'.
   */
  function setup(opts: { failWith?: unknown; sendFailsWith?: unknown } = {}) {
    const allowlist = createAllowlist(join(dir, 'allowlist.json'))
    allowlist.addEntry(USER_ID)
    const refs = createConversationRefStore()
    const calls: any[] = []
    const sends: any[] = []
    const adapter = {
      async createConversationAsync(
        appId: unknown,
        channelId: unknown,
        serviceUrl: unknown,
        audience: unknown,
        conversationParameters: unknown,
        logic: (ctx: unknown) => Promise<void>,
      ) {
        calls.push({ appId, channelId, serviceUrl, audience, conversationParameters })
        if (opts.failWith !== undefined) throw opts.failWith
        await logic({
          activity: { conversation: { id: 'a:new-1' } },
          sendActivity: async (activity: unknown) => {
            sends.push(activity)
            if (opts.sendFailsWith !== undefined) throw opts.sendFailsWith
            return { id: 'm-1' }
          },
        })
      },
    }
    const replierCalls: Array<{ conversationId: string; text: string }> = []
    const dm = createDirectMessenger({
      config: { appId: 'app-id', tenantId: TENANT_ID },
      adapter: adapter as any,
      allowlist,
      refs,
      replier: {
        sendReply: async (conversationId: string, text: string) => {
          replierCalls.push({ conversationId, text })
          return 'm-old'
        },
      },
    })
    return { dm, refs, calls, sends, replierCalls }
  }

  test('not allowlisted: RecipientNotAllowedError, nothing called', async () => {
    const { dm, calls, replierCalls } = setup()
    await expect(dm.dmMember(OTHER_ID, 'hi')).rejects.toBeInstanceOf(RecipientNotAllowedError)
    expect(calls.length).toBe(0)
    expect(replierCalls.length).toBe(0)
  })

  test('an id differing only in case and whitespace from an allowed one IS allowed', async () => {
    const { dm, calls, replierCalls } = setup()
    // It clears the allowlist gate and only then stops at the next step: with
    // an empty ref store there is no known service URL yet. The not-allowlisted
    // error is never thrown.
    await expect(dm.dmMember(`  ${USER_ID.toUpperCase()} \n`, 'hi')).rejects.toBeInstanceOf(
      NoServiceUrlError,
    )
    expect(calls.length).toBe(0)
    expect(replierCalls.length).toBe(0)
  })

  test('an existing personal ref is reused through the replier, no createConversationAsync', async () => {
    const { dm, refs, calls, replierCalls } = setup()
    refs.put('a:old', personalRef('a:old', USER_ID), USER_ID)
    const r = await dm.dmMember(USER_ID, 'hello')
    expect(r).toEqual({ conversationId: 'a:old', messageId: 'm-old', created: false })
    expect(replierCalls).toEqual([{ conversationId: 'a:old', text: 'hello' }])
    expect(calls.length).toBe(0)
  })

  test('a stored channel ref for the same user does not count as an existing 1:1 chat', async () => {
    const { dm, refs, calls } = setup()
    refs.put('19:chan@thread.tacv2', channelRef('19:chan@thread.tacv2', USER_ID), USER_ID)
    const r = await dm.dmMember(USER_ID, 'hello')
    expect(calls.length).toBe(1)
    expect(r.created).toBe(true)
    expect(r.conversationId).toBe('a:new-1')
  })

  test('create path: exact arguments, markdown send, stored personal ref, result', async () => {
    const { dm, refs, calls, sends, replierCalls } = setup()
    // The user is known from a channel conversation, so its serviceUrl is the
    // route, but their chat has to be opened fresh. The id arrives upper-cased
    // and padded: the member id must be the normalized (lower-case) one.
    refs.put('c:seed', channelRef('c:seed', USER_ID), USER_ID)
    const r = await dm.dmMember(`  ${USER_ID.toUpperCase()} `, 'hello')
    expect(calls.length).toBe(1)
    expect(calls[0].appId).toBe('app-id')
    expect(calls[0].channelId).toBe('msteams')
    expect(calls[0].serviceUrl).toBe(SERVICE_URL)
    expect(calls[0].audience).toBeUndefined()
    expect(calls[0].conversationParameters).toEqual({
      isGroup: false,
      bot: { id: 'app-id' },
      members: [{ id: USER_ID }],
      tenantId: TENANT_ID,
      channelData: { tenant: { id: TENANT_ID } },
    })
    expect(sends).toEqual([{ type: 'message', text: 'hello', textFormat: 'markdown' }])
    expect(replierCalls.length).toBe(0)
    const stored = refs.get('a:new-1')
    expect(stored).toBeDefined()
    expect(stored!.aadObjectId).toBe(USER_ID)
    expect(isPersonalReference(stored!.ref)).toBe(true)
    expect(refs.personalConversationIdFor(USER_ID)).toBe('a:new-1')
    expect(r).toEqual({ conversationId: 'a:new-1', messageId: 'm-1', created: true })
  })

  test('a bot-opened chat never becomes the permission-relay target', async () => {
    const { dm, refs, calls } = setup()
    // No 1:1 chat with anyone yet: after dm_member the relay must still find none.
    refs.put('c:seed', channelRef('c:seed', USER_ID), USER_ID)
    await dm.dmMember(USER_ID, 'hello')
    expect(calls.length).toBe(1)
    expect(refs.personalConversationIdFor(USER_ID)).toBe('a:new-1')
    expect(refs.mostRecentPersonalConversationId()).toBeUndefined()
  })

  test('the operator 1:1 chat stays the relay target when the bot opens a chat AFTER it', async () => {
    const { dm, refs, calls } = setup()
    refs.put('a:op', personalRef('a:op', OTHER_ID), OTHER_ID)
    refs.put('c:seed', channelRef('c:seed', USER_ID), USER_ID)
    // The store stamps with Date.now: a later millisecond, so a fresh stamp on the new chat would win.
    await new Promise(resolve => setTimeout(resolve, 5))
    await dm.dmMember(USER_ID, 'hello')
    expect(calls.length).toBe(1)
    expect(refs.personalConversationIdFor(USER_ID)).toBe('a:new-1')
    expect(refs.mostRecentPersonalConversationId()).toBe('a:op')
  })

  test('the new chat is stored even when the first send into it fails', async () => {
    const boom = Object.assign(new Error('send blew up'), { statusCode: 500 })
    const { dm, refs, calls } = setup({ sendFailsWith: boom })
    refs.put('c:seed', channelRef('c:seed', USER_ID), USER_ID)
    // The send error is not a 403 from conversation creation: rethrown as-is.
    expect(await errOf(dm.dmMember(USER_ID, 'hello'))).toBe(boom)
    expect(calls.length).toBe(1)
    expect(refs.get('a:new-1')).toBeDefined()
    expect(refs.personalConversationIdFor(USER_ID)).toBe('a:new-1')
  })

  test('no stored ref at all: NoServiceUrlError, no adapter call', async () => {
    const { dm, calls, replierCalls } = setup()
    await expect(dm.dmMember(USER_ID, 'hi')).rejects.toBeInstanceOf(NoServiceUrlError)
    expect(calls.length).toBe(0)
    expect(replierCalls.length).toBe(0)
  })

  test('403 without MessageWritesBlocked: PersonalInstallMissingError, one attempt, cause kept', async () => {
    const forbidden = Object.assign(new Error('Forbidden'), { statusCode: 403 })
    const { dm, refs, calls } = setup({ failWith: forbidden })
    refs.put('c:seed', channelRef('c:seed', USER_ID), USER_ID)
    const err = (await errOf(dm.dmMember(USER_ID, 'hi'))) as PersonalInstallMissingError
    expect(err).toBeInstanceOf(PersonalInstallMissingError)
    expect(err.message).toContain('not installed for them in personal scope')
    expect(err.cause).toBe(forbidden)
    expect(calls.length).toBe(1) // 403 is never retried
  })

  test('403 mentioning MessageWritesBlocked in a nested body: BotBlockedError', async () => {
    const forbidden = Object.assign(new Error('x'), {
      statusCode: 403,
      details: { error: { message: '{"subCode":"MessageWritesBlocked"}' } },
    })
    const { dm, refs, calls } = setup({ failWith: forbidden })
    refs.put('c:seed', channelRef('c:seed', USER_ID), USER_ID)
    const err = (await errOf(dm.dmMember(USER_ID, 'hi'))) as BotBlockedError
    expect(err).toBeInstanceOf(BotBlockedError)
    expect(err.message).toContain('blocked or removed the bot')
    expect(err.cause).toBe(forbidden)
    expect(calls.length).toBe(1)
  })

  test('any other error from createConversationAsync is rethrown unchanged', async () => {
    const boom = Object.assign(new Error('boom'), { statusCode: 500 })
    const { dm, refs, calls } = setup({ failWith: boom })
    refs.put('c:seed', channelRef('c:seed', USER_ID), USER_ID)
    expect(await errOf(dm.dmMember(USER_ID, 'hi'))).toBe(boom)
    expect(calls.length).toBe(1)
  })

  test('serviceUrlHint picks the most recently seen ref, not the first', () => {
    // A variable clock: two puts in the same millisecond would tie.
    let clock = 1000
    const refs = createConversationRefStore({ now: () => clock })
    refs.put('c:1', personalRef('c:1', 'user-a', 'https://smba.example/one/'), 'user-a')
    expect(refs.serviceUrlHint()).toBe('https://smba.example/one/')
    clock = 2000
    refs.put('c:2', channelRef('c:2', 'user-a', 'https://smba.example/two/'), 'user-a')
    expect(refs.serviceUrlHint()).toBe('https://smba.example/two/')
    // A later put with an older clock stamp does not win over c:2.
    clock = 1500
    refs.put('c:3', personalRef('c:3', 'user-b', 'https://smba.example/three/'), 'user-b')
    expect(refs.serviceUrlHint()).toBe('https://smba.example/two/')
  })

  test('personalConversationIdFor picks the most recent personal chat of that user', () => {
    let clock = 1000
    const refs = createConversationRefStore({ now: () => clock })
    refs.put('p:1', personalRef('p:1', 'user-a'), 'user-a')
    clock = 2000
    refs.put('g:1', channelRef('g:1', 'user-a'), 'user-a')
    expect(refs.personalConversationIdFor('user-a')).toBe('p:1')
    clock = 3000
    refs.put('p:2', personalRef('p:2', 'user-a'), 'user-a')
    expect(refs.personalConversationIdFor('user-a')).toBe('p:2')
    // Case and whitespace are normalized on both sides of the compare.
    expect(refs.personalConversationIdFor('  USER-A ')).toBe('p:2')
    // A newer channel ref does not make a group chat the 1:1 answer.
    clock = 4000
    refs.put('g:2', channelRef('g:2', 'user-a'), 'user-a')
    expect(refs.personalConversationIdFor('user-a')).toBe('p:2')
    expect(refs.personalConversationIdFor('user-b')).toBeUndefined()
    expect(createConversationRefStore().personalConversationIdFor('user-a')).toBeUndefined()
  })

  test('store methods: no crash on an empty store', () => {
    const refs = createConversationRefStore()
    expect(refs.serviceUrlHint()).toBeUndefined()
    expect(refs.personalConversationIdFor(USER_ID)).toBeUndefined()
  })
})
