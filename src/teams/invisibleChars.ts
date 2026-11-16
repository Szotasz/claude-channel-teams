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
 * is the three RGI subdivision flags (England, Scotland, Wales), which are
 * valid visible text and are kept whole. Any other tag character is removed
 * and counted, so a run of flag-shaped tags cannot carry a hidden payload.
 */

export interface InvisibleCounts {
  tag: number
  zwsp: number
  bidi: number
  bom: number
}

/** U+1F3F4 WAVING BLACK FLAG — the base of every emoji tag sequence. */
const BLACK_FLAG = 0x1f3f4
/** The tag characters after the black flag in the three RGI flags: England, Scotland, Wales (each ends with CANCEL TAG U+E007F). */
const RGI_FLAG_TAGS: number[][] = [
  [0xe0067, 0xe0062, 0xe0065, 0xe006e, 0xe0067, 0xe007f],
  [0xe0067, 0xe0062, 0xe0073, 0xe0063, 0xe0074, 0xe007f],
  [0xe0067, 0xe0062, 0xe0077, 0xe006c, 0xe0073, 0xe007f],
]

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
      // Only the three RGI subdivision flags (England, Scotland, Wales) are kept
      // whole. In any other case the flag itself stays and the tag characters
      // after it are removed and counted by the generic rule below.
      const rgi = RGI_FLAG_TAGS.some((tags) => tags.every((t, k) => points[i + 1 + k]?.codePointAt(0) === t))
      if (rgi) {
        for (let k = i; k <= i + 6; k++) kept.push(points[k]!)
        i += 7
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
