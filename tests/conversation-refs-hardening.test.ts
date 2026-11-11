import { afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync, symlinkSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore, storedReference } from '../src/teams/conversationRefs'
import {
  ConversationNotAllowedError,
  UnknownConversationError,
  createReplySender,
} from '../src/teams/reply'

const USER = '00000000-0000-4000-8000-000000000001'
const SERVICE_URL = 'https://smba.trafficmanager.net/emea/'

function fullRef(): any {
  return {
    activityId: 'a1',
    channelId: 'msteams',
    locale: 'en-US',
    serviceUrl: SERVICE_URL,
    conversation: { id: 'conv-1', conversationType: 'personal', tenantId: 't', isGroup: false, name: 'Secret Room' },
    bot: { id: 'b', name: 'Bot Name' },
    user: { id: 'u', aadObjectId: 'aad', name: 'Test Person', role: 'user' },
    extra: 1,
  }
}

function captureStderr<T>(fn: () => T): { result: T; lines: string[] } {
  const lines: string[] = []
  const orig = process.stderr.write
  process.stderr.write = ((chunk: any) => {
    lines.push(String(chunk))
    return true
  }) as any
  try {
    return { result: fn(), lines }
  } finally {
    process.stderr.write = orig
  }
}

describe('conversation refs hardening', () => {
  let dir: string
  let file: string

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'cct-refs-hard-'))
    file = join(dir, 'refs.json')
  })
  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('a) storedReference keeps only the send fields', () => {
    expect(storedReference(fullRef())).toEqual({
      activityId: 'a1',
      channelId: 'msteams',
      locale: 'en-US',
      serviceUrl: SERVICE_URL,
      conversation: { id: 'conv-1', conversationType: 'personal', tenantId: 't', isGroup: false },
      bot: { id: 'b' },
      user: { id: 'u', aadObjectId: 'aad' },
    })
    const bad = storedReference({
      channelId: 5,
      conversation: { id: 'c', isGroup: 'no', tenantId: 3 },
      bot: { id: 7 },
      user: { id: 5, aadObjectId: 'x' },
    } as any)
    expect(bad).toEqual({ conversation: { id: 'c' } })
    expect('user' in bad).toBe(false)
    expect('bot' in bad).toBe(false)
  })

  test('b) the file holds no display names; memory keeps the full ref', () => {
    const ref = fullRef()
    const a = createConversationRefStore({ file })
    a.put('conv-1', ref, USER)
    const text = readFileSync(file, 'utf8')
    for (const s of ['Secret Room', 'Bot Name', 'Test Person', '"role"', '"extra"']) {
      expect(text).not.toContain(s)
    }
    expect(a.get('conv-1')?.ref).toBe(ref)
    expect(createConversationRefStore({ file }).get('conv-1')?.ref).toEqual(storedReference(ref))
  })

  test('b2) a file written before this version (full references) loads only the send fields', () => {
    writeFileSync(file, JSON.stringify({ version: 1, refs: [{ conversationId: 'conv-1', ref: fullRef(), aadObjectId: USER, lastSeen: 1 }] }), { mode: 0o600 })
    const loaded = createConversationRefStore({ file }).get('conv-1')?.ref as any
    expect(loaded).toEqual(storedReference(fullRef()))
    expect(JSON.stringify(loaded)).not.toContain('Test Person')
    // The file itself is rewritten at load, before any new message.
    expect(readFileSync(file, 'utf8')).not.toContain('Test Person')
  })

  test('b3) no file yet (first start): no stderr line', () => {
    const { result, lines } = captureStderr(() => createConversationRefStore({ file }))
    expect(result.size()).toBe(0)
    expect(lines).toEqual([])
  })

  test.skipIf(!existsSync('/proc/self/fd'))('b4) loading closes the file it opened', () => {
    createConversationRefStore({ file }).put('conv-1', fullRef(), USER)
    const before = readdirSync('/proc/self/fd').length
    for (let i = 0; i < 20; i++) createConversationRefStore({ file })
    // A leak would add 20; a little tolerance for descriptors the runtime opens or closes meanwhile.
    expect(readdirSync('/proc/self/fd').length).toBeLessThanOrEqual(before + 2)
  })

  describe('c) refused files', () => {
    test('symlink', () => {
      const real = join(dir, 'real.json')
      createConversationRefStore({ file: real }).put('conv-1', fullRef(), USER)
      symlinkSync(real, file)
      const { result, lines } = captureStderr(() => createConversationRefStore({ file }))
      expect(result.size()).toBe(0)
      expect(lines.length).toBe(1)
      expect(lines[0]).toContain('cannot be opened (a symlink)')
    })

    test('directory', () => {
      mkdirSync(file)
      const { result, lines } = captureStderr(() => createConversationRefStore({ file }))
      expect(result.size()).toBe(0)
      expect(lines.length).toBe(1)
      expect(lines[0]).toContain('not a regular file')
    })

    test.skipIf(!existsSync('/usr/bin/mkfifo'))('a named pipe does not hang startup', () => {
      Bun.spawnSync(['/usr/bin/mkfifo', file])
      const { result, lines } = captureStderr(() => createConversationRefStore({ file }))
      expect(result.size()).toBe(0)
      expect(lines.length).toBe(1)
      expect(lines[0]).toContain('not a regular file')
    })

    test('an error while checking or reading the file: logged, empty, no crash', () => {
      createConversationRefStore({ file }).put('conv-1', fullRef(), USER)
      const { result, lines } = captureStderr(() => createConversationRefStore({ file, uid: () => { throw new Error('boom') } }))
      expect(result.size()).toBe(0)
      expect(lines.length).toBe(1)
      expect(lines[0]).toContain('cannot be read (boom)')
    })

    test('another user', () => {
      createConversationRefStore({ file }).put('conv-1', fullRef(), USER)
      const mine = process.getuid?.() ?? 0
      const { result, lines } = captureStderr(() => createConversationRefStore({ file, uid: () => mine + 1 }))
      expect(result.size()).toBe(0)
      expect(lines.length).toBe(1)
      expect(lines[0]).toContain('belongs to another user')
    })
  })

  test('d) future lastSeen is clamped to now', () => {
    const now = 1_700_000_000_000
    writeFileSync(
      file,
      JSON.stringify({
        version: 1,
        refs: [
          {
            conversationId: 'conv-1',
            ref: storedReference(fullRef()),
            aadObjectId: USER,
            lastSeen: now + 10 * 86_400_000,
          },
        ],
      }),
    )
    expect(createConversationRefStore({ file, now: () => now }).get('conv-1')?.lastSeen).toBe(now)
    // A past lastSeen is kept as it is.
    writeFileSync(file, JSON.stringify({ version: 1, refs: [{ conversationId: 'conv-1', ref: storedReference(fullRef()), aadObjectId: USER, lastSeen: now - 5000 }] }))
    expect(createConversationRefStore({ file, now: () => now }).get('conv-1')?.lastSeen).toBe(now - 5000)
  })

  test('b5) an old file whose every entry is dropped is rewritten too: no display name stays on disk', () => {
    writeFileSync(file, JSON.stringify({ version: 1, refs: [{ conversationId: 'conv-1', ref: fullRef(), aadObjectId: USER, lastSeen: 1 }] }), { mode: 0o600 })
    const store = createConversationRefStore({ file, isAllowed: () => false })
    expect(store.size()).toBe(0)
    const text = readFileSync(file, 'utf8')
    expect(text).not.toContain('Test Person')
    expect(JSON.parse(text)).toEqual({ version: 1, refs: [] })
  })

  test('b6) a file already in the current form is not rewritten at load', () => {
    createConversationRefStore({ file }).put('conv-1', fullRef(), USER)
    const before = statSync(file).ino
    createConversationRefStore({ file })
    expect(statSync(file).ino).toBe(before)
  })

  describe('e) outbound gate across a restart', () => {
    let allowlist: ReturnType<typeof createAllowlist>
    let handed: any[]
    const adapter = () =>
      ({
        async continueConversationAsync(_appId: string, ref: any, logic: any) {
          handed.push(ref)
          await logic({ sendActivity: async () => ({ id: 'm1' }) })
        },
      }) as any
    const replier = (refs: ReturnType<typeof createConversationRefStore>) =>
      createReplySender({
        config: { appId: 'app' } as any,
        adapter: adapter(),
        allowlist,
        refs,
      }).sendReply

    beforeEach(() => {
      allowlist = createAllowlist(join(dir, 'allowlist.json'))
      allowlist.addEntry(USER)
      handed = []
    })

    test('without a file the reference is lost on restart', async () => {
      const a = createConversationRefStore()
      a.put('conv-1', fullRef(), USER)
      expect(await replier(a)('conv-1', 'hi')).toBe('m1')
      const b = createConversationRefStore()
      await expect(replier(b)('conv-1', 'hi')).rejects.toBeInstanceOf(UnknownConversationError)
    })

    test('with a file the reference survives a restart', async () => {
      const a = createConversationRefStore({ file })
      a.put('conv-1', fullRef(), USER)
      expect(await replier(a)('conv-1', 'hi')).toBe('m1')
      const b = createConversationRefStore({ file, isAllowed: id => allowlist.isAllowed(id) })
      handed = []
      expect(await replier(b)('conv-1', 'hi')).toBe('m1')
      expect(handed[0].conversation.id).toBe('conv-1')
      expect(handed[0].serviceUrl).toBe(SERVICE_URL)
      // After the restart the send uses only the stored fields.
      expect(handed[0]).toEqual(storedReference(fullRef()))
    })

    test('a removed user cannot be replied to, before or after restart', async () => {
      const a = createConversationRefStore({ file })
      a.put('conv-1', fullRef(), USER)
      allowlist.removeEntry(USER)
      await expect(replier(a)('conv-1', 'hi')).rejects.toBeInstanceOf(ConversationNotAllowedError)
      const c = createConversationRefStore({ file, isAllowed: id => allowlist.isAllowed(id) })
      await expect(replier(c)('conv-1', 'hi')).rejects.toBeInstanceOf(UnknownConversationError)
      expect(handed).toEqual([])
    })

    test('an unknown conversation id is rejected before and after restart', async () => {
      const a = createConversationRefStore({ file })
      a.put('conv-1', fullRef(), USER)
      await expect(replier(a)('nope', 'hi')).rejects.toBeInstanceOf(UnknownConversationError)
      const b = createConversationRefStore({ file, isAllowed: id => allowlist.isAllowed(id) })
      await expect(replier(b)('nope', 'hi')).rejects.toBeInstanceOf(UnknownConversationError)
      expect(handed).toEqual([])
    })
  })
})
