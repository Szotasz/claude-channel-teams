/**
 * The source with every line and block comment blanked out (each comment character becomes a space, line breaks
 * kept), so a source-text check cannot be satisfied by a commented-out copy of the code it pins. String, template
 * and regex literals are kept as written. The result has the same length and the same line breaks as the input, so
 * offsets and line-anchored regexes still line up.
 *
 * The comment ranges come from the TypeScript parser, not from a scanner of our own: whether a '/' starts a regex,
 * is a division or opens a comment depends on the grammar (a postfix '++' or '!', an `if (...)` header, a property
 * named like a keyword), and only the parser knows it.
 */
import ts from 'typescript'

export function stripComments(src: string): string {
  const file = ts.createSourceFile('source.ts', src, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS)
  const out = src.split('')
  const blank = (ranges: ts.CommentRange[] | undefined): void => {
    for (const r of ranges ?? []) for (let i = r.pos; i < r.end; i++) if (out[i] !== '\n') out[i] = ' '
  }
  // Every comment is trivia before or after some token, and getChildren reaches every token, punctuation and the
  // end-of-file token included.
  const visit = (node: ts.Node): void => {
    blank(ts.getLeadingCommentRanges(src, node.pos))
    blank(ts.getTrailingCommentRanges(src, node.end))
    for (const child of node.getChildren(file)) visit(child)
  }
  visit(file)
  return out.join('')
}
