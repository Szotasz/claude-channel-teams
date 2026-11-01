/**
 * Tests for stripComments, the helper that blanks out comments before a test checks a source file's text, so a
 * commented-out copy of a pinned line can no longer satisfy the check. Every test also pins the two invariants the
 * call sites rely on: the output has the same length and the same line breaks as the input.
 */
import { describe, expect, test } from 'bun:test'
import { stripComments } from './strip-comments'

/** Both invariants: same length, and the same number of '\n' (so offsets and line-anchored regexes still line up). */
function same(out: string, src: string): boolean {
  return out.length === src.length && out.split('\n').length === src.split('\n').length
}

describe('stripComments', () => {
  test('1 a line comment becomes spaces, the newline is kept', () => {
    const src = 'a()\n// b()\nc()'
    const out = stripComments(src)
    expect(out).toBe('a()\n' + ' '.repeat(6) + '\nc()')
    expect(same(out, src)).toBe(true)
  })

  test('2 a block comment over three lines keeps its line breaks', () => {
    const src = 'x /* one\ntwo\nthree */ y'
    const out = stripComments(src)
    expect(out).not.toContain('one')
    expect(out).not.toContain('two')
    expect(out).not.toContain('three')
    expect(out.startsWith('x ')).toBe(true)
    expect(out.endsWith(' y')).toBe(true)
    expect(out.match(/\n/g)).toHaveLength(2)
    expect(same(out, src)).toBe(true)
  })

  test('3 slashes inside a string literal are not a comment, the tail after them is', () => {
    const src = "const u = 'https://example.test/a' // tail"
    const out = stripComments(src)
    expect(out).toContain('https://example.test/a')
    expect(out).not.toContain('tail')
    expect(same(out, src)).toBe(true)
  })

  test('4 a double-quoted string holding comment markers is unchanged', () => {
    const src = 'const d = "a // b /* c */"'
    const out = stripComments(src)
    expect(out).toBe(src)
    expect(same(out, src)).toBe(true)
  })

  test('5 a template literal holding comment markers is unchanged', () => {
    const src = 'const t = `x // y /* z */`'
    const out = stripComments(src)
    expect(out).toBe(src)
    expect(same(out, src)).toBe(true)
  })

  test('6 a comment inside a template substitution is blanked, the template text after it is kept', () => {
    const src = 'const t = `a ${f(/* c */ 1)} b // kept`'
    const out = stripComments(src)
    expect(out).toBe('const t = `a ${f(' + ' '.repeat(7) + ' 1)} b // kept`')
    expect(same(out, src)).toBe(true)
  })

  test('7 a template nested inside a substitution of another template', () => {
    const src = 'const t = `a ${`b ${"c"} // in`} d` // out'
    const out = stripComments(src)
    expect(out).toContain('`a ${`b ${"c"} // in`} d`')
    expect(out).not.toContain('out')
    expect(same(out, src)).toBe(true)
  })

  test('8 a regex whose body looks like comments', () => {
    const src = 'const r = /\\/\\/[^\\n]*/g // gone'
    const out = stripComments(src)
    expect(out).toContain('/\\/\\/[^\\n]*/g')
    expect(out).not.toContain('gone')
    expect(same(out, src)).toBe(true)
  })

  test('9 a quote inside a regex does not open a string', () => {
    const src = "const q = /'/\n// hidden()\nshown()"
    const out = stripComments(src)
    expect(out).toContain('shown()')
    expect(out).not.toContain('hidden()')
    expect(out).toContain("/'/")
    expect(same(out, src)).toBe(true)
  })

  test('10 a slash inside a regex character class does not end the regex', () => {
    const src = 'const s = /[/]x/ // c'
    const out = stripComments(src)
    expect(out).toContain('/[/]x/')
    expect(out).not.toContain('// c')
    expect(same(out, src)).toBe(true)
  })

  test('11 division after an identifier, a value or ")" is not a regex', () => {
    const src = 'const v = a / b // c\nconst w = (x) / 2 /* d */'
    const out = stripComments(src)
    expect(out).toContain('a / b')
    expect(out).toContain('(x) / 2')
    expect(out).not.toContain('// c')
    expect(out).not.toContain('/* d */')
    expect(same(out, src)).toBe(true)
  })

  test('12 a keyword before a slash still starts a regex', () => {
    const src = "return /'/.test(s) // c"
    const out = stripComments(src)
    expect(out).toContain("/'/.test(s)")
    expect(out).not.toContain('// c')
    expect(same(out, src)).toBe(true)
  })

  test('13 an escaped quote does not end the string', () => {
    const src = "const e = 'it\\'s // not a comment'"
    const out = stripComments(src)
    expect(out).toBe(src)
    expect(same(out, src)).toBe(true)
  })

  test('14 an unterminated block comment is blanked to the end of input', () => {
    const src = 'a /* b\nc'
    const out = stripComments(src)
    expect(out).toBe('a ' + '    ' + '\n' + ' ')
    expect(same(out, src)).toBe(true)
  })

  test('15 the review case: a commented-out guard no longer reads as present', () => {
    const src = '  // if (!allowlist.isAllowed(id)) return undefined\n  return stored'
    const out = stripComments(src)
    expect(out).not.toContain('isAllowed')
    expect(out).toContain('  return stored')
    expect(same(out, src)).toBe(true)
  })

  test('16 a guard inside a block comment is invisible to a line regex, the active line still matches', () => {
    const src = '/*\n      if (acks && ok) acks.arm(id)\n*/\nif (acks) acks.arm(id)'
    const out = stripComments(src)
    expect(out).not.toMatch(/^\s*(if \(acks && .*\) acks\.arm\(id\))$/m)
    expect(out).toMatch(/^if \(acks\) acks\.arm\(id\)$/m)
    expect(same(out, src)).toBe(true)
  })

  test('17 whitespace ends a word: a keyword after another word still starts a regex', () => {
    const src = "} else return /'/.test(s) // c"
    const out = stripComments(src)
    expect(out).toContain("/'/.test(s)")
    expect(out).not.toContain('// c')
    expect(same(out, src)).toBe(true)
  })

  test('18 a regex at the start of a substitution, whatever came before the template', () => {
    const src = "const t = f(x)`${/'/.test(s)}` // c"
    const out = stripComments(src)
    expect(out).toContain("`${/'/.test(s)}`")
    expect(out).not.toContain('// c')
    expect(same(out, src)).toBe(true)
  })
})
