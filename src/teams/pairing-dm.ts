/**
 * Private delivery path for the gated pairing hint (owner decision, Teams #21).
 *
 * When a not-yet-allowlisted member triggers the bot from a shared
 * conversation, the hint text ("This bot is gated. To pair, send me a direct
 * (1:1) message.") is delivered to that member as a 1:1 message and NEVER
 * posted into the group/channel conversation — not even as a fallback when
 * the private delivery fails. The channel is readable by everyone in it, so a
 * public hint would both nag the whole room and invite the member into a
 * conversation the allowlist gate cannot see.
 *
 * Exactly two delivery routes, and no third:
 *
 *   - A personal reference answers in place via `sendActivityRaw`: the member
 *     is already talking to the bot 1:1, so the reply stays private.
 *   - Anything else (channel, group chat, unknown type — fail closed) asks the
 *     adapter to open a bot-initiated 1:1 conversation with the member and
 *     sends the hint through that conversation alone.
 *
 * Errors from either route propagate unchanged: the caller in
 * src/teams/adapter.ts catches them and logs to stderr. Swallowing here would
 * drop the hint silently, and a channel fallback would publish it — both
 * wrong. That is why `sendActivityRaw` is never called for a non-personal
 * reference, under any circumstance.
 */

import type { ConversationReference } from 'botbuilder'

export interface PairingDmDeps {
  /** Reply into an existing 1:1 reference. Ungated on purpose: the pairing DM is bot-initiated. */
  sendActivityRaw: (ref: Partial<ConversationReference>, text: string) => Promise<void>
  /** True only for a 1:1 chat with the bot — see conversationRefs.isPersonalReference. */
  isPersonal: (ref: Partial<ConversationReference>) => boolean
  /**
   * Open a bot-initiated 1:1 conversation with the member the ref belongs to
   * and run `logic` inside it. Production wiring is
   * CloudAdapter.createConversationAsync, following src/teams/directMessage.ts.
   */
  createConversation: (
    ref: Partial<ConversationReference>,
    logic: (ctx: { sendActivity: (text: string) => Promise<unknown> }) => Promise<void>,
  ) => Promise<void>
}

/**
 * Build the `sendPairingDm` hook: the returned sender delivers `text` to the
 * member behind `ref` as a 1:1 message only. There is no public fallback.
 */
export function makePairingDmSender(
  opts: PairingDmDeps,
): (ref: Partial<ConversationReference>, text: string) => Promise<void> {
  return async (ref, text) => {
    if (opts.isPersonal(ref)) {
      await opts.sendActivityRaw(ref, text)
      return
    }
    // Group/channel reference (or a type we do not recognize): the hint goes
    // through the bot-initiated 1:1 conversation or not at all. Any error —
    // opening the conversation or sending inside it — rethrows unchanged.
    await opts.createConversation(ref, async ctx => {
      await ctx.sendActivity(text)
    })
  }
}
