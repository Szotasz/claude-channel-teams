/**
 * Invisible-character stripping for inbound message text.
 *
 * A forwarded or pasted Teams message can carry characters a person cannot see
 * but the model reads anyway ("ASCII smuggling": Unicode tag characters
 * U+E0000..U+E007F encode hidden ASCII payloads), or characters that reorder
 * what the reader sees (bidi controls). stripInvisible removes those before
 * the text is handed to Claude Code, counting what it removed per class so the
 * channel meta can surface it (see invisibleMeta).
 *
 * Kept intact, on purpose: emoji ZWJ sequences (U+200D), ZWNJ (U+200C),
 * variation selectors (U+FE00..U+FE0F) and all visible text — removing those
 * would corrupt legitimate rendering. The one exception inside the tag block
 * is the emoji tag sequence (regional flags such as Scotland or England),
 * which is valid visible text and is kept whole.
 */

export interface InvisibleCounts {
  tag: number
  zwsp: number
  bidi: number
  bom: number
}

/** U+1F3F4 WAVING BLACK FLAG — the base of every emoji tag sequence. */
const BLACK_FLAG = 0x1f3f4
/** U+E007F CANCEL TAG — terminates an emoji tag sequence. */
const CANCEL_TAG = 0xe007f

/** Tag characters valid inside an emoji tag sequence: tag digits U+E0030..U+E0039 and tag lowercase U+E0061..U+E007A. */
function isTagSequenceChar(cp: number): boolean {
  return (cp >= 0xe0030 && cp <= 0xe0039) || (cp >= 0xe0061 && cp <= 0xe007a)
}

/** Any code point of the Supplemental tag block: U+E0000..U+E007F. */
function isTagChar(cp: number): boolean {
  return cp >= 0xe0000 && cp <= 0xe007f
}

/** Zero-width spaces and word joiners: U+200B, U+2060..U+2064. */
function isZeroWidth(cp: number): boolean {
  return cp === 0x200b || (cp >= 0x2060 && cp <= 0x2064)
}

/** Bidirectional embedding / isolate controls: U+202A..U+202E and U+2066..U+2069. */
function isBidi(cp: number): boolean {
  return (cp >= 0x202a && cp <= 0x202e) || (cp >= 0x2066 && cp <= 0x2069)
}

export function stripInvisible(text: string): { text: string; counts: InvisibleCounts } {
  const counts: InvisibleCounts = { tag: 0, zwsp: 0, bidi: 0, bom: 0 }
  // for..of iterates by code point, so an astral character (the black flag,
  // other emoji) is never split into surrogate halves.
  const points: string[] = []
  for (const ch of text) points.push(ch)

  const kept: string[] = []
  let i = 0
  while (i < points.length) {
    const ch = points[i]!
    const cp = ch.codePointAt(0)!
    if (cp === BLACK_FLAG) {
      // A valid emoji tag sequence: black flag, 1..7 tag sequence characters,
      // CANCEL TAG. Kept whole (e.g. the flag of England). A run that does not
      // complete that shape falls through: the flag itself always stays and the
      // tag characters after it are removed by the generic rule below.
      let j = i + 1
      while (j < points.length && isTagSequenceChar(points[j]!.codePointAt(0)!)) j++
      const tagChars = j - i - 1
      if (tagChars >= 1 && tagChars <= 7 && points[j]?.codePointAt(0) === CANCEL_TAG) {
        for (let k = i; k <= j; k++) kept.push(points[k]!)
        i = j + 1
        continue
      }
      kept.push(ch)
      i++
      continue
    }
    if (isTagChar(cp)) {
      counts.tag++
    } else if (isZeroWidth(cp)) {
      counts.zwsp++
    } else if (cp === 0xfeff) {
      counts.bom++
    } else if (isBidi(cp)) {
      counts.bidi++
    } else {
      kept.push(ch)
    }
    i++
  }

  if (counts.tag === 0 && counts.zwsp === 0 && counts.bidi === 0 && counts.bom === 0) {
    return { text, counts }
  }
  return { text: kept.join(''), counts }
}

/** The channel-meta form of the counts: only non-zero classes, decimal strings. */
export function invisibleMeta(counts: InvisibleCounts): Record<string, string> {
  const meta: Record<string, string> = {}
  if (counts.tag > 0) meta.invisible_tag = String(counts.tag)
  if (counts.zwsp > 0) meta.invisible_zwsp = String(counts.zwsp)
  if (counts.bidi > 0) meta.invisible_bidi = String(counts.bidi)
  if (counts.bom > 0) meta.invisible_bom = String(counts.bom)
  return meta
}

/**
 * Strip an inbound event's text in place of the pipeline: returns the SAME
 * object when nothing was removed, otherwise a shallow copy carrying the
 * stripped text and the removal counts under `invisible`.
 */
export function withInvisibleStripped<E extends { text: string }>(event: E): E & { invisible?: InvisibleCounts } {
  const { text, counts } = stripInvisible(event.text)
  if (counts.tag === 0 && counts.zwsp === 0 && counts.bidi === 0 && counts.bom === 0) {
    // The cast is the identity at runtime: `invisible` is optional, so E already
    // satisfies the intersection — TS just cannot prove it for a generic.
    return event as E & { invisible?: InvisibleCounts }
  }
  return { ...event, text, invisible: counts }
}
