// Turns rich message parts (HTML, cards) into plain text for the agent. It never
// fetches anything: an `<img src>` or a card image URL is never requested. The
// image itself arrives as its own image attachment, and fetching attacker-chosen
// URLs with the bot's token is an exfiltration surface. Only http(s) and mailto
// link targets are printed.

export const MAX_INLINE_CHARS = 64 * 1024

export const CARD_CONTENT_TYPES: ReadonlySet<string> = new Set([
  'application/vnd.microsoft.card.adaptive',
  'application/vnd.microsoft.card.hero',
  'application/vnd.microsoft.card.thumbnail',
  'application/vnd.microsoft.teams.card.o365connector',
])

const SAFE_URL_RE = /^(https?:\/\/|mailto:)/i

function isSafeUrl(u: string): boolean {
  return SAFE_URL_RE.test(u)
}

function decodeCodePoint(n: number): string {
  if (!Number.isInteger(n) || n < 0 || n > 0x10ffff) return ''
  if (n >= 0xd800 && n <= 0xdfff) return ''
  if (n < 0x20 && n !== 9 && n !== 10) return ''
  if (n >= 0x7f && n <= 0x9f) return ''
  return String.fromCodePoint(n)
}

function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos|nbsp);/gi, (m, e: string) => {
    const l = e.toLowerCase()
    if (l[0] === '#') {
      const n = l[1] === 'x' ? parseInt(l.slice(2), 16) : parseInt(l.slice(1), 10)
      return decodeCodePoint(n)
    }
    switch (l) {
      case 'amp': return '&'
      case 'lt': return '<'
      case 'gt': return '>'
      case 'quot': return '"'
      case 'apos': return "'"
      case 'nbsp': return ' '
    }
    return m
  })
}

// The scanner is linear, so the input bound can be wider than the output cap (capInline, 64K): the HTML is
// cut to this many characters before conversion.
export const MAX_HTML_INPUT_CHARS = 256 * 1024

// One pass, and every search moves forward (indexOf from the current position, never back), so the
// cost is linear in the input; unclosed tags cannot make it rescan.
function render(html: string, opts: { dropImages: boolean; dropMentions: boolean }): string {
  // ASCII-only case fold: String#toLowerCase can change the length (U+0130 becomes two code units), and
  // offsets found in `lower` are used on `html`.
  const lower = html.replace(/[A-Z]+/g, (u) => u.toLowerCase())
  // The next real closing tag of `name` at or after `from`: `</name` followed by whitespace, '/' or '>'
  // (so </scripture> does not close <script>). Forward-only, so the scan stays linear.
  const closingTag = (name: string, from: number): number => {
    const needle = `</${name}`
    for (let at = lower.indexOf(needle, from); at >= 0; at = lower.indexOf(needle, at + 1)) {
      const next = lower.charAt(at + needle.length)
      if (next === '' || next === '>' || next === '/' || /\s/.test(next)) return at
    }
    return -1
  }
  const out: string[] = []
  let anchor: { href: string; start: number } | null = null
  let i = 0
  const n = html.length
  while (i < n) {
    const lt = html.indexOf('<', i)
    if (lt < 0) {
      out.push(html.slice(i))
      break
    }
    if (lt > i) out.push(html.slice(i, lt))
    if (html.startsWith('<!--', lt)) {
      const end = html.indexOf('-->', lt + 4)
      if (end < 0) break
      i = end + 3
      continue
    }
    const gt = html.indexOf('>', lt + 1)
    if (gt < 0) {
      out.push(html.slice(lt))
      break
    }
    const chunk = html.slice(lt + 1, gt)
    i = gt + 1
    const m = /^(\/?)\s*([a-zA-Z][a-zA-Z0-9:-]*)/.exec(chunk)
    if (!m) continue
    const closing = m[1] === '/'
    const name = m[2]!.toLowerCase()
    if (!closing && (name === 'script' || name === 'style')) {
      const close = closingTag(name, i)
      if (close < 0) break
      const end = html.indexOf('>', close)
      if (end < 0) break
      i = end + 1
      continue
    }
    if (!closing && opts.dropMentions && (name === 'at' || (name === 'span' && /schema\.skype\.com\/mention/i.test(chunk)))) {
      const close = closingTag(name, i)
      if (close < 0) break
      const end = html.indexOf('>', close)
      if (end < 0) break
      i = end + 1
      continue
    }
    switch (name) {
      case 'a':
        if (!closing) {
          if (anchor) break
          const h = /(?:^|\s)href\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+))/i.exec(chunk)
          if (h) anchor = { href: (h[1] ?? h[2] ?? h[3] ?? '').trim(), start: out.length }
        } else if (anchor) {
          const { href, start } = anchor
          anchor = null
          // The scheme is checked on the DECODED target (https&#58;// is https://); the raw text is printed and
          // decoded with the rest of the text below, so nothing is decoded twice.
          if (isSafeUrl(decodeEntities(href).trim())) {
            const inner = out.slice(start).join('')
            if (inner.trim() === href) out.length = start, out.push(href)
            else out.push(` (${href})`)
          }
        }
        break
      case 'img':
        if (!closing && !opts.dropImages) {
          const a = /(?:^|\s)alt\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(chunk)
          const alt = (a?.[1] ?? a?.[2] ?? '').trim()
          out.push(alt ? `[image: ${alt}]` : '[image]')
        }
        break
      case 'br':
        out.push('\n')
        break
      case 'blockquote':
        out.push(closing ? ' [/quote]\n' : '\n[quote] ')
        break
      case 'li':
        out.push(closing ? '\n' : '- ')
        break
      case 'p': case 'div': case 'tr': case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6':
      case 'table': case 'ul': case 'ol':
        if (closing) out.push('\n')
        break
      case 'td': case 'th':
        if (closing) out.push(' | ')
        break
    }
  }
  let s = decodeEntities(out.join(''))
  s = s.replace(/\r/g, '')
  s = s.replace(/[ \t]+/g, ' ')
  s = s.split('\n').map((l) => l.trim()).join('\n')
  s = s.replace(/\n{3,}/g, '\n\n')
  return s.trim()
}

export function htmlToText(html: string): string {
  const cut = html.length > MAX_HTML_INPUT_CHARS
  const s = render(cut ? html.slice(0, MAX_HTML_INPUT_CHARS) : html, { dropImages: false, dropMentions: false })
  return cut ? `${s}\n[... HTML cut at ${MAX_HTML_INPUT_CHARS} characters before conversion]` : s
}

const TRUNCATION_NOTE = '\n[... truncated at 65536 characters]'
export function capInline(text: string): string {
  if (text.length <= MAX_INLINE_CHARS) return text
  return text.slice(0, MAX_INLINE_CHARS - TRUNCATION_NOTE.length) + TRUNCATION_NOTE
}

// For deciding whether an HTML part only repeats the message: the HTML without its images and
// without any @mention markup, as text, whitespace-collapsed. The message side is the same with
// its <at>...</at> mentions removed (a channel message arrives with the bot's own mention cut out,
// and other mentions as raw <at> tags), so a mention never makes the two differ.
export function comparableHtml(html: string): string {
  return render(html.slice(0, MAX_HTML_INPUT_CHARS), { dropImages: true, dropMentions: true }).replace(/\s+/g, ' ').trim()
}
export function comparableMessage(text: string): string {
  return render(text.slice(0, MAX_HTML_INPUT_CHARS), { dropImages: false, dropMentions: true }).replace(/\s+/g, ' ').trim()
}

// ── cards ───────────────────────────────────────────────────────────────────

type Obj = Record<string, unknown>

function isObj(v: unknown): v is Obj {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

const MAX_CARD_INPUT_CHARS = 4 * MAX_INLINE_CHARS

interface Budget { used: number; over: boolean }

/** String field value, passed through htmlToText when it carries a real tag. Counts against the budget. */
function readStr(v: unknown, budget: Budget): string {
  if (typeof v !== 'string' || budget.over) return ''
  budget.used += v.length
  if (budget.used > MAX_CARD_INPUT_CHARS) {
    budget.over = true
    return ''
  }
  return (/<\/?[a-zA-Z][^>]*>/.test(v) ? htmlToText(v) : v).trim()
}

function withUrl(title: string, url: string): string {
  return url && isSafeUrl(url) ? (title ? `${title} (${url})` : url) : title
}

const MAX_DEPTH = 20
const MAX_NODES = 2000

function adaptiveToLines(root: Obj, budget: Budget): string[] {
  const str = (v: unknown): string => readStr(v, budget)
  const out: string[] = []
  let nodes = 0
  const walkList = (v: unknown, depth: number): void => {
    if (!Array.isArray(v)) return
    for (const n of v) walk(n, depth)
  }
  const walk = (node: unknown, depth: number): void => {
    if (depth > MAX_DEPTH || nodes >= MAX_NODES || !isObj(node)) return
    nodes++
    const type = typeof node.type === 'string' ? node.type : ''
    switch (type) {
      case 'TextBlock':
        out.push(str(node.text))
        return
      case 'RichTextBlock':
        if (Array.isArray(node.inlines)) {
          out.push(
            node.inlines
              .map((i) => (isObj(i) ? str(i.text) : typeof i === 'string' ? str(i) : ''))
              .join(''),
          )
        }
        return
      case 'FactSet':
        if (Array.isArray(node.facts)) {
          for (const f of node.facts) {
            if (!isObj(f)) continue
            const t = str(f.title)
            const v = str(f.value)
            out.push(t && v ? `${t}: ${v}` : t || v)
          }
        }
        return
      case 'Container':
      case 'Column':
        walkList(node.items, depth + 1)
        return
      case 'ColumnSet':
        walkList(node.columns, depth + 1)
        return
      case 'ActionSet':
        walkList(node.actions, depth + 1)
        return
      case 'Action.OpenUrl':
        out.push(withUrl(str(node.title), typeof node.url === 'string' ? node.url.trim() : ''))
        return
      case 'Action.Submit':
      case 'Action.Execute':
        out.push(`[button] ${str(node.title)}`)
        return
      case 'Action.ShowCard':
        out.push(`[button] ${str(node.title)}`)
        if (isObj(node.card)) {
          walkList(node.card.body, depth + 1)
          walkList(node.card.actions, depth + 1)
        }
        return
      default:
        walkList(node.items, depth + 1)
        walkList(node.columns, depth + 1)
        walkList(node.body, depth + 1)
        walkList(node.actions, depth + 1)
        // Adaptive 1.5 Table: rows[] of TableRow, each with cells[] of TableCell holding items[].
        walkList(node.rows, depth + 1)
        walkList(node.cells, depth + 1)
    }
  }
  walkList(root.body, 1)
  walkList(root.actions, 1)
  return out
}

function heroToLines(c: Obj, budget: Budget): string[] {
  const str = (v: unknown): string => readStr(v, budget)
  const out = [str(c.title), str(c.subtitle), str(c.text)]
  if (Array.isArray(c.buttons)) {
    for (const b of c.buttons) {
      if (!isObj(b)) continue
      const title = str(b.title)
      const value = typeof b.value === 'string' ? b.value.trim() : ''
      if (b.type === 'openUrl' && isSafeUrl(value)) out.push(title ? `${title} (${value})` : value)
      else out.push(`[button] ${title}`)
    }
  }
  return out
}

function connectorToLines(c: Obj, budget: Budget): string[] {
  const str = (v: unknown): string => readStr(v, budget)
  const out = [str(c.title), str(c.summary), str(c.text)]
  if (Array.isArray(c.sections)) {
    for (const s of c.sections) {
      if (!isObj(s)) continue
      out.push(str(s.activityTitle), str(s.activitySubtitle), str(s.activityText), str(s.title), str(s.text))
      if (Array.isArray(s.facts)) {
        for (const f of s.facts) {
          if (!isObj(f)) continue
          const n = str(f.name)
          const v = str(f.value)
          out.push(n && v ? `${n}: ${v}` : n || v)
        }
      }
    }
  }
  if (Array.isArray(c.potentialAction)) {
    for (const a of c.potentialAction) {
      if (!isObj(a)) continue
      const name = str(a.name)
      const uris: string[] = []
      if (Array.isArray(a.targets)) {
        for (const t of a.targets) if (isObj(t) && typeof t.uri === 'string') uris.push(t.uri.trim())
      }
      if (Array.isArray(a.target)) {
        for (const t of a.target) if (typeof t === 'string') uris.push(t.trim())
      }
      const uri = uris.find(isSafeUrl)
      out.push(uri ? `${name} (${uri})` : `[button] ${name}`)
    }
  }
  return out
}

export function cardToText(contentType: string, content: unknown): string {
  if (!isObj(content)) return ''
  const budget: Budget = { used: 0, over: false }
  let lines: string[]
  switch (contentType) {
    case 'application/vnd.microsoft.card.adaptive':
      lines = adaptiveToLines(content, budget)
      break
    case 'application/vnd.microsoft.card.hero':
    case 'application/vnd.microsoft.card.thumbnail':
      lines = heroToLines(content, budget)
      break
    case 'application/vnd.microsoft.teams.card.o365connector':
      lines = connectorToLines(content, budget)
      break
    default:
      return ''
  }
  const kept = lines.filter((l) => l.length > 0)
  if (budget.over) kept.push('[... card cut: too large]')
  return kept.join('\n')
}
