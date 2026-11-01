/**
 * The source with every line and block comment blanked out (each comment character becomes a space, line breaks
 * kept), so a source-text check cannot be satisfied by a commented-out copy of the code it pins. String, template
 * and regex literals are kept as written. The result has the same length and the same line breaks as the input, so
 * offsets and line-anchored regexes still line up.
 */

/** Keywords after which a '/' still starts a regex literal rather than division. */
const REGEX_KEYWORDS = new Set([
  'return', 'typeof', 'case', 'do', 'else', 'in', 'of', 'void', 'yield', 'await', 'delete', 'instanceof', 'new',
  'throw',
])

const isWordChar = (c: string): boolean => /[A-Za-z0-9_$]/.test(c)

export function stripComments(src: string): string {
  // Comments are blanked in place; every other character is copied unchanged, so length and line breaks survive.
  const out = src.split('')
  let i = 0
  // The previous significant token, which decides whether a lone '/' starts a regex or is division. Whitespace and
  // comments do not count as significant.
  let lastCh = '' // '' means start of input, where a '/' always starts a regex
  let lastWord = '' // the identifier ending at lastCh, when lastCh is a word character
  let afterLiteral = false // a string, template or regex literal just closed, so its value makes '/' division
  let wordEnded = false // whitespace, a comment or a literal ended the identifier in lastWord

  function regexAllowed(): boolean {
    if (lastCh === '') return true
    if (afterLiteral) return false
    if (lastCh === ')' || lastCh === ']') return false
    return !isWordChar(lastCh) || REGEX_KEYWORDS.has(lastWord)
  }

  /** One ordinary code character: advance and remember it as the previous significant token. */
  function take(): void {
    const c = src[i]
    if (/\s/.test(c)) wordEnded = true
    else {
      lastWord = isWordChar(c) ? (wordEnded ? c : lastWord + c) : ''
      lastCh = c
      afterLiteral = false
      wordEnded = false
    }
    i++
  }

  /** A line comment: blank through to, not including, the next '\n' (or the end of input). */
  function blankLineComment(): void {
    wordEnded = true
    while (i < src.length && src[i] !== '\n') out[i++] = ' '
  }

  /** A block comment through its terminator or the end of input; '\n' and '\r' are kept so lines still line up. */
  function blankBlockComment(): void {
    wordEnded = true
    out[i] = ' '
    out[i + 1] = ' '
    i += 2
    while (i < src.length) {
      if (src[i] === '*' && src[i + 1] === '/') {
        out[i] = ' '
        out[i + 1] = ' '
        i += 2
        return
      }
      if (src[i] !== '\n' && src[i] !== '\r') out[i] = ' '
      i++
    }
  }

  /** A '...' or "..." literal, copied unchanged; a backslash escapes the next character. */
  function scanString(quote: string): void {
    i++ // the opening quote
    while (i < src.length) {
      const c = src[i]
      if (c === '\\') { i += 2; continue }
      if (c === quote) { i++; break }
      if (c === '\n') return // malformed: an unescaped line break must not swallow the rest of the file
      i++
    }
    afterLiteral = true
    wordEnded = true
  }

  /** A template literal, copied unchanged except that ${...} substitutions are code and scanned as such. */
  function scanTemplate(): void {
    i++ // the opening backtick
    while (i < src.length) {
      const c = src[i]
      if (c === '\\') { i += 2; continue }
      if (c === '`') { i++; break }
      if (c === '$' && src[i + 1] === '{') {
        i += 2
        // the substitution starts like code after a '{': a '/' at its start is a regex
        lastCh = '{'
        lastWord = ''
        afterLiteral = false
        scanCode(true)
        if (i < src.length && src[i] === '}') i++ // the '}' scanCode stopped at
        continue
      }
      i++
    }
    afterLiteral = true
    wordEnded = true
  }

  /** A regex literal, copied unchanged: '\' escapes the next character, '[...]' is a class where '/' does not end it. */
  function scanRegex(): void {
    i++ // the opening '/'
    let inClass = false
    while (i < src.length) {
      const c = src[i]
      if (c === '\\') { i += 2; continue }
      if (c === '\n') break // malformed: treat the regex as ended at the line break
      if (inClass) {
        if (c === ']') inClass = false
      } else if (c === '[') inClass = true
      else if (c === '/') { i++; break }
      i++
    }
    afterLiteral = true
    wordEnded = true
  }

  /** Code: comments, literals and everything else. In a substitution it stops at the '}' closing it (depth 0). */
  function scanCode(inSubstitution: boolean): void {
    let depth = 0
    while (i < src.length) {
      const c = src[i]
      if (c === '/') {
        const next = src[i + 1]
        if (next === '/') { blankLineComment(); continue }
        if (next === '*') { blankBlockComment(); continue }
        if (!regexAllowed()) { take(); continue } // division
        scanRegex()
        continue
      }
      if (c === "'" || c === '"') { scanString(c); continue }
      if (c === '`') { scanTemplate(); continue }
      if (inSubstitution && c === '}') {
        if (depth === 0) return
        depth--
      } else if (inSubstitution && c === '{') depth++
      take()
    }
  }

  scanCode(false)
  return out.join('')
}
