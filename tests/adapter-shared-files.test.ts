/**
 * Adapter wiring for files shared in a channel or a group chat
 * (deps.sharedFiles, src/teams/adapter.ts).
 *
 * The handler must fire only for a shared conversation whose activity carries an
 * <attachment id="…"> placeholder and no direct file attachment, and its
 * annotations must land in the delivered text the same way the 1:1 attachment
 * path's do. The injected fake records the lookup it was handed, so these tests
 * also pin the exact shape the production fetchSharedFiles call receives.
 *
 * The sender name is a fixture, not a person.
 */

import { describe, test, expect, beforeEach } from 'bun:test'
import { mkdtempSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { BotAdapter, TurnContext, type Activity } from 'botbuilder'

import { makeTurnHandler } from '../src/teams/adapter'
import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import type { SharedFileLookup } from '../src/teams/sharedFiles'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'
const TEAM = '19:team0001@thread.tacv2'
const CHAN = '19:chan0001@thread.tacv2'
const CHAT = '19:chat0001@thread.v2'
const SKYPE_CHAN = '19:skype0001@thread.skype'
const OTHER_CHAT = '19:chat0002@thread.v2'
const SERVICE = 'https://smba.trafficmanager.net/emea/00000000-0000-4000-8000-000000000002/'
const ANNOTATION = '[Attached: /x/report.xlsx (application/xlsx, 10 B)]'

const HTML_WITH_REF = '<p>see this</p><attachment id="ref-1"></attachment>'

class NoopAdapter extends BotAdapter {
  async sendActivities(): Promise<any> { return [] }
  async updateActivity(): Promise<void> {}
  async deleteActivity(): Promise<void> {}
  async continueConversation(): Promise<void> {}
}

function makeActivity(overrides: Partial<Activity> = {}): Activity {
  return {
    type: 'message',
    text: 'see this',
    id: 'MSG',
    timestamp: new Date('2026-07-01T10:30:00Z'),
    from: { id: 'user-1', name: 'Test User', aadObjectId: FIXTURE_ID },
    conversation: { id: `${CHAN};messageid=ROOT`, tenantId: TENANT_ID, conversationType: 'channel' },
    channelData: { team: { id: TEAM }, channel: { id: CHAN } },
    recipient: { id: 'bot-1', name: 'Bot' },
    channelId: 'msteams',
    serviceUrl: SERVICE,
    attachments: [{ contentType: 'text/html', content: HTML_WITH_REF }] as any,
    ...overrides,
  } as unknown as Activity
}

describe('adapter: shared files in channels and group chats', () => {
  let dir: string
  let allowlistFile: string

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cct-adapter-shared-'))
    allowlistFile = join(dir, 'allowlist.json')
  })

  /** A turn handler with an allowlisted sender, shared conversations opted in for
   *  unmentioned delivery, and a recording sharedFiles fake. */
  function setup(
    sharedFiles: (lookup: SharedFileLookup, fromName: string) => Promise<{ annotations: string[] }>,
    unmentioned: Set<string> = new Set([CHAN, CHAT, SKYPE_CHAN, OTHER_CHAT]),
  ) {
    const allowlist = createAllowlist(allowlistFile)
    allowlist.addEntry(FIXTURE_ID)
    const events: any[] = []
    const lookups: SharedFileLookup[] = []
    const names: string[] = []
    const turn = makeTurnHandler({
      config: {
        stateDir: '/tmp/unused',
        allowlistFile: '/tmp/unused',
        appId: 'app-id',
        appPassword: 'app-pw',
        appType: 'SingleTenant' as const,
        tenantId: TENANT_ID,
        bindHost: '127.0.0.1',
        port: 0,
        receivedFilesDir: '/tmp/unused',
        unmentionedConversations: unmentioned,
      } as any,
      adapter: {} as any,
      allowlist,
      refs: createConversationRefStore(),
      onEvent: e => events.push(e),
      sharedFiles: async (lookup, fromName) => {
        lookups.push(lookup)
        names.push(fromName)
        return sharedFiles(lookup, fromName)
      },
    })
    const run = (a: Activity) => turn(new TurnContext(new NoopAdapter() as any, a))
    return { run, events, lookups, names }
  }

  test('A1 a channel reply calls sharedFiles with the channel lookup and appends the annotation', async () => {
    const s = setup(async () => ({ annotations: [ANNOTATION] }))
    await s.run(makeActivity())
    expect(s.lookups.length).toBe(1)
    expect(s.lookups[0]).toEqual({
      kind: 'channel',
      teamThreadId: TEAM,
      channelId: CHAN,
      messageId: 'MSG',
      threadRootId: 'ROOT',
      serviceUrl: SERVICE,
      aadGroupId: undefined,
      refIds: ['ref-1'],
    })
    expect(s.names).toEqual(['Test User'])
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe(`see this\n\n${ANNOTATION}`)
    expect(s.events[0].hasAttachments).toBe(true)
    rmSync(dir, { recursive: true, force: true })
  })

  test('A2 the same html in a personal (1:1) chat does not call sharedFiles', async () => {
    const s = setup(async () => ({ annotations: [ANNOTATION] }))
    await s.run(makeActivity({
      conversation: { id: 'a:personal-1', tenantId: TENANT_ID, conversationType: 'personal' } as any,
    }))
    expect(s.lookups.length).toBe(0)
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('see this')
    rmSync(dir, { recursive: true, force: true })
  })

  test('A3 a channel message that also carries a file.download.info attachment does not call sharedFiles', async () => {
    const s = setup(async () => ({ annotations: [ANNOTATION] }))
    await s.run(makeActivity({
      attachments: [
        { contentType: 'text/html', content: HTML_WITH_REF },
        { contentType: 'application/vnd.microsoft.teams.file.download.info', content: { downloadUrl: 'https://contoso.sharepoint.com/x' } },
      ] as any,
    }))
    expect(s.lookups.length).toBe(0)
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('see this')
    rmSync(dir, { recursive: true, force: true })
  })

  test('A4 a group chat calls sharedFiles with the chat lookup and the sender id', async () => {
    const s = setup(async () => ({ annotations: [ANNOTATION] }))
    await s.run(makeActivity({
      conversation: { id: CHAT, tenantId: TENANT_ID, conversationType: 'groupChat', isGroup: true } as any,
    }))
    expect(s.lookups.length).toBe(1)
    expect(s.lookups[0]).toEqual({
      kind: 'chat',
      chatId: CHAT,
      messageId: 'MSG',
      senderAadObjectId: FIXTURE_ID,
      refIds: ['ref-1'],
    })
    expect(s.events[0].text).toBe(`see this\n\n${ANNOTATION}`)
    rmSync(dir, { recursive: true, force: true })
  })

  test('A5 a throwing sharedFiles still delivers the message with a failure note', async () => {
    const s = setup(async () => { throw new Error('boom') })
    await s.run(makeActivity())
    expect(s.lookups.length).toBe(1)
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toContain('see this')
    expect(s.events[0].text).toContain('[Attachment handling failed: boom]')
    rmSync(dir, { recursive: true, force: true })
  })

  test('A6 an activity without an id, or with an empty one, does not call sharedFiles', async () => {
    for (const id of [undefined, '']) {
      const s = setup(async () => ({ annotations: [ANNOTATION] }))
      await s.run(makeActivity({ id }))
      expect(s.lookups.length).toBe(0)
      expect(s.events.length).toBe(1)
    }
    rmSync(dir, { recursive: true, force: true })
  })

  test('A7 a channel message without team/channel data in channelData yields no lookup', async () => {
    const s = setup(async () => ({ annotations: [ANNOTATION] }))
    await s.run(makeActivity({ channelData: {} as any }))
    expect(s.lookups.length).toBe(0)
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('see this')
    rmSync(dir, { recursive: true, force: true })
  })

  test('X17 a channel typed conversation with a @thread.skype id still gets the channel lookup', async () => {
    const s = setup(async () => ({ annotations: [ANNOTATION] }))
    await s.run(makeActivity({
      conversation: { id: `${SKYPE_CHAN};messageid=ROOT`, tenantId: TENANT_ID, conversationType: 'channel' } as any,
    }))
    expect(s.lookups.length).toBe(1)
    expect(s.lookups[0]).toEqual({
      kind: 'channel',
      teamThreadId: TEAM,
      channelId: CHAN,
      messageId: 'MSG',
      threadRootId: 'ROOT',
      serviceUrl: SERVICE,
      aadGroupId: undefined,
      refIds: ['ref-1'],
    })
    expect(s.events[0].text).toBe(`see this\n\n${ANNOTATION}`)
    rmSync(dir, { recursive: true, force: true })
  })

  test('X17b a chat is recognised from conversationType alone and from isGroup alone', async () => {
    const byType = setup(async () => ({ annotations: [] }))
    await byType.run(makeActivity({
      conversation: { id: CHAT, tenantId: TENANT_ID, conversationType: 'groupChat' } as any,
    }))
    expect(byType.lookups.length).toBe(1)
    expect(byType.lookups[0]).toEqual({
      kind: 'chat',
      chatId: CHAT,
      messageId: 'MSG',
      senderAadObjectId: FIXTURE_ID,
      refIds: ['ref-1'],
    })

    const byFlag = setup(async () => ({ annotations: [] }))
    await byFlag.run(makeActivity({
      conversation: { id: OTHER_CHAT, tenantId: TENANT_ID, isGroup: true } as any,
      channelData: {} as any,
    }))
    expect(byFlag.lookups.length).toBe(1)
    expect(byFlag.lookups[0]).toEqual({
      kind: 'chat',
      chatId: OTHER_CHAT,
      messageId: 'MSG',
      senderAadObjectId: FIXTURE_ID,
      refIds: ['ref-1'],
    })
    rmSync(dir, { recursive: true, force: true })
  })

  test('X17c a sharedFiles answer with no annotations leaves the text untouched', async () => {
    const s = setup(async () => ({ annotations: [] }))
    await s.run(makeActivity())
    expect(s.lookups.length).toBe(1)
    expect(s.events.length).toBe(1)
    expect(s.events[0].text).toBe('see this')
    rmSync(dir, { recursive: true, force: true })
  })
})
