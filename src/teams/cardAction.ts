/**
 * A click on a card button this bot sent. It is NOT text the user typed: the title, verb and data come from the
 * card (which the model wrote), so the event is marked as a card action and never goes through the permission-reply match.
 */

export interface CardAction {
  /** 'automatic': an Action.Execute the client fired by itself (a card refresh), not a click. */
  kind: 'execute' | 'submit' | 'automatic'
  title?: string
  verb?: string
  dataJson: string
}

export const MAX_CARD_DATA_CHARS = 2000

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

export function cleanLabel(value: unknown, max = 100): string | undefined {
  if (typeof value !== 'string') return undefined
  const cleaned = value
    .replace(/[\u0000-\u001f\u007f-\u009f"<>\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, max)
    .trim()
  return cleaned === '' ? undefined : cleaned
}

function json(x: unknown): string {
  let s: string
  try {
    s = JSON.stringify(x === undefined ? null : x)
  } catch {
    return '"[unreadable data]"'
  }
  if (typeof s !== 'string') return '"[unreadable data]"'
  if (s.length <= MAX_CARD_DATA_CHARS) return s
  // Still JSON and still within the cap: the start of the data as a string, marked as truncated.
  let keep = MAX_CARD_DATA_CHARS
  let out = JSON.stringify({ truncated: true, start: s.slice(0, keep) })
  while (out.length > MAX_CARD_DATA_CHARS && keep > 0) {
    keep = Math.max(0, keep - (out.length - MAX_CARD_DATA_CHARS))
    out = JSON.stringify({ truncated: true, start: s.slice(0, keep) })
  }
  return out
}

/** undefined when the activity is not a card click. */
export function cardActionFrom(activity: { type?: string; name?: string; value?: unknown; text?: string }): CardAction | undefined {
  if (activity.type === 'invoke' && activity.name === 'adaptiveCard/action') {
    const v = activity.value
    if (isPlainObject(v) && v.trigger === 'automatic') return { kind: 'automatic', dataJson: 'null' }
    const raw = isPlainObject(v) ? v.action : undefined
    const action = isPlainObject(raw) ? raw : {}
    return { kind: 'execute', title: cleanLabel(action.title), verb: cleanLabel(action.verb), dataJson: json(action.data) }
  }
  if (activity.type === 'message' && isPlainObject(activity.value)) {
    return { kind: 'submit', dataJson: json(activity.value) }
  }
  return undefined
}

export function cardActionText(action: CardAction): string {
  const head = `[Card action]${action.title ? ` title="${action.title}"` : ''}${action.verb ? ` verb="${action.verb}"` : ''}`
  return `${head}\ndata: ${action.dataJson}`
}
