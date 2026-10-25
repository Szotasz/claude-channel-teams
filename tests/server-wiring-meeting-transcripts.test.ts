/**
 * Wiring checks for the meeting-transcripts option. src/server.ts starts the listener and the MCP transport on
 * import, so like tests/server-wiring.test.ts these read the source, scoped to the block that does the work. The
 * config side is parseFlag, already covered in tests/server-wiring-meeting-events.test.ts.
 */
import { describe, expect, test } from 'bun:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

const read = (...p: string[]) => readFileSync(join(import.meta.dir, '..', 'src', ...p), 'utf8')
const server = read('server.ts')

/** The block that starts at `anchor`, from its first '{' up to the matching closing brace. */
function block(src: string, anchor: string): string {
  const start = src.indexOf(anchor)
  expect(start).toBeGreaterThan(-1)
  expect(src.indexOf(anchor, start + 1)).toBe(-1) // exactly one such block
  let depth = 0
  for (let i = src.indexOf('{', start); i >= 0 && i < src.length; i++) {
    if (src[i] === '{') depth++
    else if (src[i] === '}' && --depth === 0) return src.slice(start, i + 1)
  }
  throw new Error(`unbalanced block: ${anchor}`)
}

describe('server.ts meeting-transcripts wiring', () => {
  test('transcriptsOn needs both flags: the events and the transcripts one', () => {
    expect(server).toContain('const transcriptsOn = config.meetingEvents && config.meetingTranscripts')
  })

  test('the meeting store is a version-1 JSON file in the state directory', () => {
    expect(server).toContain("createMeetingStore(join(config.stateDir, 'meeting-transcripts.json'))")
  })

  test('the tool list carries the transcript tool under transcriptsOn', () => {
    const tools = block(server, 'mcp.setRequestHandler(ListToolsRequestSchema,')
    expect(tools).toContain('...transcriptTools(transcriptsOn),')
  })

  test('the get_meeting_transcript case refuses the call when transcripts are off', () => {
    const c = block(server, "case 'get_meeting_transcript': {")
    expect(c).toContain("if (!transcriptsOn) throw new Error('meeting transcripts are not enabled')")
  })

  test('the transcript instruction appears once, only inside the transcriptsOn branch', () => {
    const sentences = server.match(/'get_meeting_transcript\(meeting_id\)[^\n]*/g) ?? []
    expect(sentences.length).toBe(1)
    expect(server).toMatch(/\.\.\.\(transcriptsOn\n\s*\? \[\n\s*'get_meeting_transcript\(meeting_id\)/)
  })

  test('the listener gets the meeting store only under transcriptsOn', () => {
    const listener = block(server, 'const listener = startListener({')
    expect(listener).toContain('...(transcriptsOn ? { meetingStore } : {}),')
    expect(listener).not.toContain('meetingStore,')
  })
})
