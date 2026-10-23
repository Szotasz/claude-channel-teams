/**
 * Adaptive Cards that the reply tool sends. The card comes from the model, so it is checked before it
 * goes out: the shape Teams renders, a size Teams accepts, and https-only links and images.
 */

export const ADAPTIVE_CARD_CONTENT_TYPE = 'application/vnd.microsoft.card.adaptive'
export const MAX_CARD_CHARS = 28000

const MAX_DEPTH = 32

type Result = { ok: true } | { ok: false; error: string }

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

function walk(value: unknown, path: string, depth: number): string | undefined {
  if (typeof value !== 'object' || value === null) return undefined
  if (depth > MAX_DEPTH) return `card${path ? '.' + path : ''} is nested too deeply (more than ${MAX_DEPTH} levels)`
  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) {
      const err = walk(value[i], `${path}[${i}]`, depth + 1)
      if (err) return err
    }
    return undefined
  }
  for (const [key, child] of Object.entries(value)) {
    const childPath = path ? `${path}.${key}` : key
    if (
      (/url$/i.test(key) || key === 'backgroundImage') &&
      typeof child === 'string' &&
      !child.startsWith('https://')
    ) {
      return `card.${childPath} must be an https:// URL`
    }
    const err = walk(child, childPath, depth + 1)
    if (err) return err
  }
  return undefined
}

export function validateAdaptiveCard(card: unknown): Result {
  if (!isPlainObject(card)) return { ok: false, error: 'card must be a JSON object' }
  if (card.type !== 'AdaptiveCard') {
    return { ok: false, error: 'card.type must be "AdaptiveCard"' }
  }
  if (typeof card.version !== 'string' || !/^1\.[0-5]$/.test(card.version)) {
    return { ok: false, error: 'card.version must be 1.0 to 1.5' }
  }
  const size = JSON.stringify(card).length
  if (size > MAX_CARD_CHARS) {
    return { ok: false, error: `card is too large (${size} characters, at most ${MAX_CARD_CHARS})` }
  }
  const err = walk(card, '', 0)
  if (err) return { ok: false, error: err }
  return { ok: true }
}
