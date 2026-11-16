import { describe, expect, test } from 'bun:test'
import { capInline, cardToText, comparableHtml, comparableMessage, htmlToText, MAX_INLINE_CHARS, neutralizeMarkup } from '../src/teams/richContent'

describe('htmlToText', () => {
  test('forwarded message fixture', () => {
    const html = '<div><p>Tender: <b>Napelem park</b></p><p>Deadline&nbsp;2026-11-05</p><a href="https://portal.example.test/t/123">open</a><br><table><tr><td>A</td><td>B</td></tr></table></div>'
    expect(htmlToText(html)).toBe('Tender: Napelem park\nDeadline 2026-11-05\nopen (https://portal.example.test/t/123)\nA | B |')
  })
  test('script removed', () => expect(htmlToText('<script>alert(1)</script>x')).toBe('x'))
  test('escaped tags stay literal', () => expect(htmlToText('&lt;b&gt;')).toBe('<b>'))
  test('javascript link dropped', () => expect(htmlToText('<a href="javascript:alert(1)">click</a>')).toBe('click'))
  test('link equal to text', () => expect(htmlToText('<a href="https://x.test">https://x.test</a>')).toBe('https://x.test'))
  test('img is never fetched, alt shown', () => expect(htmlToText('<img src="https://evil.test/p.png" alt="logo">')).toBe('[image: logo]'))
  test('blockquote', () => {
    const t = htmlToText('<blockquote itemtype="http://schema.skype.com/Reply"><strong>Anna</strong><p>see the PDF</p></blockquote>reply')
    for (const s of ['[quote]', 'see the PDF', '[/quote]', 'reply']) expect(t).toContain(s)
  })
  test('numeric entities', () => expect(htmlToText('&#0;a&#x41;')).toBe('aA'))
})

describe('capInline boundary', () => {
  test('exactly the cap is kept, one more is cut to the cap', () => {
    const at = 'a'.repeat(MAX_INLINE_CHARS)
    expect(capInline(at)).toBe(at)
    const over = capInline('a'.repeat(MAX_INLINE_CHARS + 1))
    expect(over.length).toBe(MAX_INLINE_CHARS)
    expect(over.endsWith('\n[... truncated at 65536 characters]')).toBe(true)
  })
  test('neutralizeMarkup keeps < and > as entities', () => {
    expect(neutralizeMarkup('a<b>c</b>')).toBe('a&lt;b&gt;c&lt;/b&gt;')
  })
})

describe('capInline', () => {
  test('truncates', () => {
    const t = capInline('a'.repeat(70000))
    const note = '\n[... truncated at 65536 characters]'
    expect(t.startsWith('a'.repeat(MAX_INLINE_CHARS - note.length))).toBe(true)
    expect(t.length).toBe(MAX_INLINE_CHARS)
    expect(t.endsWith('\n[... truncated at 65536 characters]')).toBe(true)
    expect(capInline('abc')).toBe('abc')
  })
})

describe('cardToText', () => {
  test('adaptive', () => {
    const t = cardToText('application/vnd.microsoft.card.adaptive', {
      type: 'AdaptiveCard',
      body: [
        { type: 'TextBlock', text: 'Hello' },
        { type: 'FactSet', facts: [{ title: 'a', value: '1' }, { title: 'b', value: '2' }] },
      ],
      actions: [
        { type: 'Action.OpenUrl', title: 'Open', url: 'https://x.test' },
        { type: 'Action.Submit', title: 'Go' },
      ],
    })
    expect(t).toBe('Hello\na: 1\nb: 2\nOpen (https://x.test)\n[button] Go')
  })
  test('hero', () => {
    const t = cardToText('application/vnd.microsoft.card.hero', {
      title: 'T', text: 'body', buttons: [{ type: 'openUrl', title: 'Go', value: 'https://x.test' }],
    })
    expect(t).toBe('T\nbody\nGo (https://x.test)')
  })
  test('o365connector', () => {
    const t = cardToText('application/vnd.microsoft.teams.card.o365connector', {
      title: 'Build', sections: [{ activityTitle: 'Act', facts: [{ name: 'k', value: 'v' }] }],
      potentialAction: [{ '@type': 'OpenUri', name: 'View', targets: [{ os: 'default', uri: 'https://x.test/b' }] }],
    })
    expect(t).toBe('Build\nAct\nk: v\nView (https://x.test/b)')
  })
  test('javascript url not printed', () => {
    expect(cardToText('application/vnd.microsoft.card.adaptive', {
      body: [{ type: 'Action.OpenUrl', title: 'Bad', url: 'javascript:alert(1)' }],
    })).toBe('Bad')
  })
  test('non-object', () => expect(cardToText('application/vnd.microsoft.card.hero', 'x')).toBe(''))
})

describe('htmlToText input bound', () => {
  test('a huge HTML full of unclosed anchors is cut before conversion and converts in bounded time', () => {
    const html = 'a'.repeat(300_000)
    const t0 = Date.now()
    const out = htmlToText(html)
    expect(Date.now() - t0).toBeLessThan(3000)
    expect(out).toContain('[... HTML cut at 262144 characters before conversion]')
  })

  test('an HTML at the bound is not cut', () => {
    expect(htmlToText('a'.repeat(262_144))).not.toContain('HTML cut')
  })
})

describe('round-1 fixes', () => {
  test('capInline stays within the cap including the note', () => {
    const t = capInline('a'.repeat(70000))
    expect(t.length).toBe(65536)
    expect(t.endsWith('\n[... truncated at 65536 characters]')).toBe(true)
  })
  test('unquoted href', () => {
    expect(htmlToText('<a href=https://example.test/document>Open document</a>')).toBe('Open document (https://example.test/document)')
  })
  test('comparable helpers', () => {
    expect(comparableHtml('<p><span itemscope="" itemtype="http://schema.skype.com/Mention" itemid="0">Bot</span> hello <img src="x"></p>')).toBe('hello')
    expect(comparableMessage('hello <at>Anna</at>')).toBe('hello')
    expect(comparableHtml('<p>[image: invoice number]</p>')).toBe('[image: invoice number]')
  })
  test('card text with < and > is not treated as HTML', () => {
    const t = cardToText('application/vnd.microsoft.card.adaptive', { body: [{ type: 'TextBlock', text: 'a < b and c > d' }] })
    expect(t).toBe('a < b and c > d')
  })
  test('card input budget', () => {
    const body = Array.from({ length: 300 }, () => ({ type: 'TextBlock', text: 'x'.repeat(1000) }))
    const t = cardToText('application/vnd.microsoft.card.adaptive', { body })
    expect(t.endsWith('[... card cut: too large]')).toBe(true)
  })
})

describe('round-2 fixes', () => {
  test('a whole input at the bound of bare < characters stays fast (each < is looked at once)', () => {
    const big = '<'.repeat(256 * 1024)
    const t0 = Date.now()
    htmlToText(big); comparableHtml(big); comparableMessage(big)
    expect(Date.now() - t0).toBeLessThan(300)
  })

  test('the dedupe helpers are bounded on unclosed img and mention tags', () => {
    const inputs = [
      '<'.repeat(65536),
      '<img'.repeat(16384),
      '<at '.repeat(16384),
      '<at>x'.repeat(13107),
      '<a href="https://x.test/">a '.repeat(2340),
      '<script>'.repeat(8192),
    ]
    for (const s of inputs) {
      const t0 = Date.now()
      htmlToText(s)
      comparableHtml(s)
      comparableMessage(s)
      expect(Date.now() - t0).toBeLessThan(200)
    }
  })

  test('an entity-encoded https target is kept', () => {
    expect(htmlToText('<a href="https&#58;//example.test/document">Open document</a>')).toBe('Open document (https://example.test/document)')
    expect(htmlToText('<a href="javascript&#58;alert(1)">x</a>')).toBe('x')
  })

  test('an escaped entity is decoded once', () => {
    expect(htmlToText('&amp;#39;')).toBe('&#39;')
  })

  test('an adaptive Table is read', () => {
    const card = { type: 'AdaptiveCard', body: [{ type: 'Table', rows: [{ type: 'TableRow', cells: [{ type: 'TableCell', items: [{ type: 'TextBlock', text: 'Cell A' }] }] }] }] }
    expect(cardToText('application/vnd.microsoft.card.adaptive', card)).toContain('Cell A')
  })
})

describe('single-pass scanner', () => {
  test('a < with no > after it stays', () => expect(htmlToText('a < b and c')).toBe('a < b and c'))
  test('comments are dropped', () => expect(htmlToText('x<!-- c -->y')).toBe('xy'))
  test('script and style are dropped', () => expect(htmlToText('<script>a</script>b<style>c</style>d')).toBe('bd'))
  test('nested anchors use the outer target', () => {
    expect(htmlToText('<a href="https://a.test">one <a href="https://b.test">two</a></a>')).toContain('one two (https://a.test)')
  })
})

describe('round-3 fixes', () => {
  test('a length-changing lowercase character before a script does not shift the scan', () => {
    expect(htmlToText('\u0130'.repeat(20) + '<script>x</script><p>ok</p>')).toContain('ok')
  })

  test('a longer tag name does not close script, style or a mention', () => {
    expect(htmlToText('<script>hidden</scripture>still hidden</script><p>safe</p>')).toBe('safe')
    expect(comparableHtml('<at>Bot</atx>name</at> hello')).toBe('hello')
  })

  test('only a real href attribute names the link target', () => {
    expect(htmlToText('<a data-href="https://other.test" href="https://real.test">Open</a>')).toBe('Open (https://real.test)')
    expect(htmlToText('<img data-alt="no" alt="yes">')).toBe('[image: yes]')
  })
})

