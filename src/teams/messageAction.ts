/**
 * Teams message actions (the "..." / "More actions" menu of a message).
 *
 * When the user submits one, Teams sends an invoke activity named
 * 'composeExtension/submitAction'. This turns it into the text Claude gets.
 * Everything from the quoted message and the form is untrusted data.
 */

import type { Activity } from 'botbuilder'

export type MessageAction = { command: string; text: string }

const MAX_MESSAGE = 4000
const MAX_NOTE = 2000
const MAX_FROM = 100
const MAX_COMMAND = 64
const TEAMS_LINK_PREFIX = 'https://teams.microsoft.com/'

function cut(s: string, max: number): string {
  return s.length > max ? `${s.slice(0, max)} [...]` : s
}

function htmlToText(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div)\s*>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

export function messageActionFrom(activity: Activity): MessageAction | undefined {
  if (activity.type !== 'invoke' || activity.name !== 'composeExtension/submitAction') return undefined
  const value = activity.value as Record<string, any> | null | undefined
  if (!value || typeof value !== 'object' || value.commandContext !== 'message') return undefined

  const rawCommand = typeof value.commandId === 'string' ? value.commandId : ''
  const command = rawCommand.replace(/[^A-Za-z0-9_-]/g, '').slice(0, MAX_COMMAND) || 'unknown'

  const payload = (value.messagePayload ?? {}) as Record<string, any>
  const displayName = payload.from?.user?.displayName
  const from = typeof displayName === 'string'
    ? displayName
        .replace(/[\u0000-\u001f\u007f-\u009f"<>]/g, '')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, MAX_FROM)
    : ''

  const lines = [`[Message action] command="${command}"`, `from: ${from || '-'}`]
  const link = payload.linkToMessage
  // No whitespace or control character: a link must stay on its own line.
  if (typeof link === 'string' && link.startsWith(TEAMS_LINK_PREFIX) && !/[\s\u0000-\u001f\u007f-\u009f]/.test(link)) {
    lines.push(`link: ${link}`)
  }

  // The note (what the user typed) comes before the quoted message, and every line of the quoted message is prefixed
  // with '> ': the message was written by someone else, so it must not be able to forge a 'note:' or 'link:' line.
  const rawNote = value.data?.note
  const note = typeof rawNote === 'string' ? cut(rawNote.trim(), MAX_NOTE) : ''
  lines.push(`note: ${note || '-'}`)

  const content = payload.body?.content
  let message = typeof content === 'string' ? content : ''
  if (payload.body?.contentType === 'html') message = htmlToText(message)
  message = cut(message, MAX_MESSAGE)
  if (message) lines.push('message:', ...message.split('\n').map(l => `> ${l}`))
  else lines.push('message: -')

  return { command, text: lines.join('\n') }
}
