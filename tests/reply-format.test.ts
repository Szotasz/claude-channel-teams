/**
 * Reply format tests — textFormat on sendReply / editReply.
 */

import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, rmSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'

import { createAllowlist } from '../src/pairing/allowlist'
import { createConversationRefStore } from '../src/teams/conversationRefs'
import { createReplySender, parseReplyFormat } from '../src/teams/reply'

const FIXTURE_ID = '00000000-0000-4000-8000-000000000001'
const TENANT_ID = '00000000-0000-4000-8000-000000000002'

function setup() {
  const dir = mkdtempSync(join(tmpdir(), 'cct-reply-format-'))
  const allowlist = createAllowlist(join(dir, 'allowlist.json'))
  allowlist.addEntry(FIXTURE_ID)
  const refs = createConversationRefStore()
  refs.put('conv-abc', { conversation: { id: 'conv-abc' } } as any, FIXTURE_ID)
  const sent: any[] = []
  const updated: any[] = []
  const adapter = {
    async continueConversationAsync(_appId: string, _ref: any, logic: any) {
      await logic({
        async sendActivity(act: any) {
          sent.push(act)
        },
        async updateActivity(act: any) {
          updated.push(act)
        },
      })
    },
  } as any
  const config = {
    stateDir: '/tmp/unused',
    allowlistFile: '/tmp/unused',
    appId: 'app-id-from-config',
    appPassword: 'app-pw',
    appType: 'SingleTenant' as const,
    tenantId: TENANT_ID,
    bindHost: '127.0.0.1',
    port: 0,
  }
  const replier = createReplySender({ config, adapter, allowlist, refs })
  return { replier, sent, updated, cleanup: () => rmSync(dir, { recursive: true }) }
}

describe('reply format', () => {
  test('F1 sendReply defaults to markdown, honours extendedmarkdown', async () => {
    const { replier, sent, cleanup } = setup()
    await replier.sendReply('conv-abc', 'hi')
    await replier.sendReply('conv-abc', '| a |\n|---|\n| 1 |', undefined, { format: 'extendedmarkdown' })
    expect(sent.map(a => a.textFormat)).toEqual(['markdown', 'extendedmarkdown'])
    cleanup()
  })

  test('F2 editReply defaults to markdown, honours extendedmarkdown', async () => {
    const { replier, updated, cleanup } = setup()
    await replier.editReply('conv-abc', 'm1', 'hi')
    await replier.editReply('conv-abc', 'm1', '- [ ] todo', 'extendedmarkdown')
    expect(updated.map(a => a.textFormat)).toEqual(['markdown', 'extendedmarkdown'])
    expect(updated[1].id).toBe('m1')
    cleanup()
  })
})

describe('parseReplyFormat', () => {
  test('omitted or null is markdown, the two formats pass, anything else is an error', () => {
    expect(parseReplyFormat(undefined)).toBe('markdown')
    expect(parseReplyFormat(null)).toBe('markdown')
    expect(parseReplyFormat('markdown')).toBe('markdown')
    expect(parseReplyFormat('extendedmarkdown')).toBe('extendedmarkdown')
    for (const bad of ['plain', '', 'xml', 42, true, {}]) {
      expect(() => parseReplyFormat(bad)).toThrow('format must be markdown or extendedmarkdown')
    }
  })
})

describe('server.ts format wiring', () => {
  const server = readFileSync(join(import.meta.dir, '..', 'src', 'server.ts'), 'utf8')
  const caseBody = (name: string) => {
    const start = server.indexOf(`      case '${name}': {`)
    expect(start).toBeGreaterThan(-1)
    return server.slice(start, server.indexOf("\n      case '", start + 1))
  }
  test('reply parses args.format and passes a non-default format in the options', () => {
    const body = caseBody('reply')
    expect(body).toContain('const format = parseReplyFormat(args.format)')
    expect(body).toContain("...(format !== 'markdown' ? { format } : {}),")
  })
  test('edit_message parses args.format and passes it to editReply', () => {
    const body = caseBody('edit_message')
    expect(body).toContain('const format = parseReplyFormat(args.format)')
    expect(body).toContain('await replier.editReply(conversationId, messageId, text, format)')
  })
})
