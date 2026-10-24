/**
 * stripInvisible — removal of invisible Unicode characters (tag characters,
 * zero-width spaces / joiners, bidi controls, BOM) from inbound Teams message
 * text, with per-class counts. Defense against "ASCII smuggling": text a
 * person cannot see but the model reads.
 *
 * Every special character here is written as an escape, never as a literal
 * invisible character.
 */
import { describe, expect, test } from 'bun:test'
import {
  invisibleMeta,
  stripInvisible,
  withInvisibleStripped,
} from '../src/teams/invisibleChars'

/** The flag of England: black flag + tag letters g b e n g + CANCEL TAG. */
const ENGLAND = '\u{1F3F4}\u{E0067}\u{E0062}\u{E0065}\u{E006E}\u{E0067}\u{E007F}'
const ZERO_COUNTS = { tag: 0, zwsp: 0, bidi: 0, bom: 0 }

describe('invisible character stripping', () => {
  test('I1 each class removed and counted, visible text unchanged', () => {
    // 3 tag characters not after a black flag, 2 U+200B, 1 U+2060, 1 U+FEFF,
    // 1 U+202E and 1 U+2066.
    const input =
      'hello\u{E0041} world\u{E0042}!\u{E0043}' +
      'a\u{200B}b\u{200B}c' +
      'x\u{2060}y' +
      'p\u{FEFF}q' +
      'l\u{202E}m\u{2066}n'
    const { text, counts } = stripInvisible(input)
    expect(text).toBe('hello world!abcxypqlmn')
    expect(counts).toEqual({ tag: 3, zwsp: 3, bidi: 2, bom: 1 })
  })

  test('I2 the flag of England (emoji tag sequence) is kept whole', () => {
    const input = `cheers ${ENGLAND}!`
    const { text, counts } = stripInvisible(input)
    expect(text).toBe(`cheers ${ENGLAND}!`)
    expect([...text]).toEqual([...input]) // byte-identical, tag sequence intact
    expect(counts).toEqual(ZERO_COUNTS)
  })

  test('I3 a flag-shaped run without the closing CANCEL TAG loses its tags, keeps the flag', () => {
    const { text, counts } = stripInvisible('\u{1F3F4}\u{E0067}\u{E0062}!')
    expect(text).toBe('\u{1F3F4}!')
    expect(counts).toEqual({ ...ZERO_COUNTS, tag: 2 })
  })

  test('I4 a lone CANCEL TAG and 8 tag letters after a black flag are all removed', () => {
    const eight = '\u{E0061}\u{E0062}\u{E0063}\u{E0064}\u{E0065}\u{E0066}\u{E0067}\u{E0068}'
    // Lone \u{E007F}, then black flag + 8 tag letters + \u{E007F} (one letter
    // past the 7-letter valid-sequence limit).
    const input = `a\u{E007F}b\u{1F3F4}${eight}\u{E007F}c`
    const { text, counts } = stripInvisible(input)
    expect(text).toBe('ab\u{1F3F4}c')
    expect(counts).toEqual({ ...ZERO_COUNTS, tag: 10 }) // 1 lone + 8 letters + 1 closing
  })

  test('I4b a black flag with no tag letter before CANCEL TAG is no sequence: the cancel tag goes', () => {
    const r = stripInvisible('x\u{1F3F4}\u{E007F}y')
    expect(r.text).toBe('x\u{1F3F4}y')
    expect(r.counts).toEqual({ tag: 1, zwsp: 0, bidi: 0, bom: 0 })
  })

  test('I4c a tag outside the letter/digit range breaks the sequence: every tag in it goes', () => {
    // U+E007B is a tag character, but not a tag letter or digit, so this is not a valid flag.
    const r = stripInvisible('\u{1F3F4}\u{E0067}\u{E007B}\u{E007F}')
    expect(r.text).toBe('\u{1F3F4}')
    expect(r.counts).toEqual({ tag: 3, zwsp: 0, bidi: 0, bom: 0 })
  })

  test('I4d a flag with tag digits, and one with seven tag characters (the upper bound), are kept whole', () => {
    const withDigits = '\u{1F3F4}\u{E0075}\u{E0073}\u{E0030}\u{E0031}\u{E007F}'
    const seven = '\u{1F3F4}' + '\u{E0061}'.repeat(7) + '\u{E007F}'
    for (const flag of [withDigits, seven]) {
      const r = stripInvisible(flag)
      expect(r.text).toBe(flag)
      expect(r.counts).toEqual({ tag: 0, zwsp: 0, bidi: 0, bom: 0 })
    }
  })

  test('I5 ZWJ sequences, ZWNJ and variation selectors are kept identical', () => {
    const family = '\u{1F468}\u{200D}\u{1F469}\u{200D}\u{1F467}'
    const input = `${family} says na\u{200C}ive \u{2764}\u{FE0F}`
    const { text, counts } = stripInvisible(input)
    expect(text).toBe(`${family} says na\u{200C}ive \u{2764}\u{FE0F}`)
    expect(text === input).toBe(true)
    expect(counts).toEqual(ZERO_COUNTS)
  })

  test('I6 clean text returns the same string value and zero counts', () => {
    const input = 'plain text 123 with emoji \u{1F600}'
    const result = stripInvisible(input)
    expect(result.text === input).toBe(true)
    expect(result.counts).toEqual(ZERO_COUNTS)
    expect(invisibleMeta(result.counts)).toEqual({})
  })

  test('I7 invisibleMeta emits only non-zero classes with exact key names', () => {
    expect(invisibleMeta(ZERO_COUNTS)).toEqual({})
    expect(invisibleMeta({ tag: 2, zwsp: 0, bidi: 1, bom: 0 })).toEqual({
      invisible_tag: '2',
      invisible_bidi: '1',
    })
    const all = invisibleMeta({ tag: 1, zwsp: 2, bidi: 3, bom: 4 })
    expect(all).toEqual({
      invisible_tag: '1',
      invisible_zwsp: '2',
      invisible_bidi: '3',
      invisible_bom: '4',
    })
    for (const value of Object.values(all)) expect(typeof value).toBe('string')
  })

  test('I8 withInvisibleStripped copies only when something was removed', () => {
    const clean = { text: 'hello', conversationId: 'c-1' }
    expect(withInvisibleStripped(clean)).toBe(clean)
    expect('invisible' in withInvisibleStripped(clean)).toBe(false)

    const dirty = { text: 'a\u{200B}b\u{E0041}c', conversationId: 'c-1' }
    const stripped = withInvisibleStripped(dirty)
    expect(stripped).not.toBe(dirty)
    expect(stripped.text).toBe('abc')
    expect(stripped.conversationId).toBe('c-1')
    expect(stripped.invisible).toEqual({ tag: 1, zwsp: 1, bidi: 0, bom: 0 })
    // The original event is not mutated.
    expect(dirty.text).toBe('a\u{200B}b\u{E0041}c')
    expect(dirty.invisible).toBeUndefined()
  })
})
