# Marveen approval cards (`MARVEEN_APPROVAL_URL`)

On a Marveen install whose main channel is Teams, the Marveen dashboard can ask its owner to approve an outgoing email with a card in the owner's 1:1 chat with the bot. The card has **Approve** and **Reject** buttons (`Action.Execute`, verb `marveen-approval`). The dashboard sends that card itself, not through this plugin, and the dashboard decides the click.

This plugin's part is to make sure the click reaches the dashboard and **never reaches the Claude session**. The agent whose email it is must not be on the path.

1. A click is an `adaptiveCard/action` invoke on `POST /api/messages`. After the channel-token check (`src/teams/channelToken.ts`, see docs/security.md), an invoke with the verb `marveen-approval` is taken in `handleMessagesPost`. CloudAdapter, the turn handler and the session never see it.
2. The plugin forwards the click only if all of these hold:
   - `MARVEEN_APPROVAL_URL` is set to a loopback URL (`127.0.0.1`, `::1` or `localhost`, http or https, no credentials);
   - the conversation's tenant is the bot's tenant;
   - the clicker is on the allowlist.

   It forwards the activity, with its original `Authorization` header, to that URL. Redirects are not followed, and the timeout is 8 s.
3. The dashboard checks the Bot Framework channel token itself, then the tenant, the owner, the card it posted and the card's one-time nonce, and decides. It answers with a card (Approved / Rejected / Expired / Already resolved / Not accepted). The plugin returns that card as the invoke response, and Teams shows it in place of the clicked card.
4. In every other case the clicker sees a fixed "decide on the dashboard" card. Those cases are: the URL is not set, a gate fails, the dashboard is unreachable, or the dashboard answers without a valid card.

A model-supplied card may not use the verb `marveen-approval`: `validateAdaptiveCard` refuses it, as it refuses `refresh`. Otherwise an agent could make a button that would be taken for an approval click.

`MARVEEN_APPROVAL_URL` is set by Marveen's channel launcher for the main agent's Teams plugin (for example `http://127.0.0.1:3420/api/approvals/teams-card-action`). Unset, approval-card clicks are answered with the fixed card and go nowhere.
