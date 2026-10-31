/**
 * Feedback buttons on replies (opt-in, TEAMS_FEEDBACK). A reply sent with `feedbackChannelData()` shows Teams' thumbs
 * up / down buttons. A click arrives as an invoke activity (`message/submitAction`, actionName 'feedback'); the
 * adapter answers it and, for an allowlisted same-tenant person, delivers it to Claude as a channel event whose text is
 * `[Feedback] like|dislike` plus the optional comment, with message_id = the rated message.
 */

export const MAX_FEEDBACK_CHARS = 2000
/** A new object for every reply, so nothing that touches one message's channelData can reach another's. */
export function feedbackChannelData(): { feedbackLoop: { type: 'default' } } {
  return { feedbackLoop: { type: 'default' } }
}

const isObject = (v: unknown): v is Record<string, unknown> =>
  typeof v === 'object' && v !== null && !Array.isArray(v)

export function isFeedbackInvoke(activity: { type?: string; name?: string; value?: unknown }): boolean {
  return (
    activity.type === 'invoke' &&
    activity.name === 'message/submitAction' &&
    isObject(activity.value) &&
    activity.value.actionName === 'feedback'
  )
}

export function parseFeedback(value: unknown): { reaction: 'like' | 'dislike'; comment: string } | undefined {
  if (!isObject(value) || !isObject(value.actionValue)) return undefined
  const { reaction, feedback } = value.actionValue
  if (reaction !== 'like' && reaction !== 'dislike') return undefined
  let comment = ''
  if (typeof feedback === 'string') {
    try {
      const parsed: unknown = JSON.parse(feedback)
      if (isObject(parsed) && typeof parsed.feedbackText === 'string') {
        // Control, line- and paragraph-separator and bidi characters become spaces; newline and tab stay.
        const cleaned = parsed.feedbackText
          .replace(/[\u0000-\u0008\u000b-\u001f\u007f-\u009f\u2028\u2029\u202a-\u202e\u2066-\u2069]/g, ' ')
          .trim()
        // Cut by code point, so a surrogate pair is never split.
        comment = Array.from(cleaned).slice(0, MAX_FEEDBACK_CHARS).join('')
      }
    } catch {}
  }
  return { reaction, comment }
}

export function feedbackEventText(f: { reaction: 'like' | 'dislike'; comment: string }): string {
  return `[Feedback] ${f.reaction}` + (f.comment ? `\n${f.comment}` : '')
}
