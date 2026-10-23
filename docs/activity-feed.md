# Activity feed notifications (`send_activity_notification`)

The `send_activity_notification` tool sends a Microsoft Teams activity feed
notification to one person about a team channel or a channel thread, through
Microsoft Graph (`POST /teams/{teamId}/sendActivityNotification`) with the bot's
app-only token.

- **Off by default.** The tool, and the line that tells Claude about it,
  exist only when `TEAMS_ACTIVITY_FEED=1`.
- **Channels and channel threads only.** Group chats and 1:1 chats are not
  supported in this version. The conversation must be allowed, as for `reply`.
- **Recipient.** Only people on the allowlist can be notified. The sender of an
  inbound message always is (inbound messages are only delivered from
  allowlisted senders), so "notify the person who asked" works; anyone else is
  refused before any Graph call.
- **Rate limit.** At most 3 notifications per recipient in 10 minutes, kept in
  memory (a restart resets it). A send that fails does not count. At most
  1000 recipients are tracked; when that is full and none of them has expired,
  a new recipient is refused rather than resetting someone else's count.
- **Manifest.** Generate the app package with `scripts/make-teams-manifest.sh
  --activity-feed`. It adds the resource-specific consent (RSC) permission
  `TeamsActivity.Send.Group` and the `webApplicationInfo` block (combinable with
  `--rsc-all-messages`: one block, one permission list). A team that already
  has the app must upgrade it and consent to the new permission.
- **Activity type.** The reserved `systemDefault` type is used, so the manifest
  declares no activity types.
- **Text.** At most 150 characters (Teams shows the first 150 of the preview).
- **Team id.** Learned from an inbound message in the channel, as for
  `get_thread_history`; before that the tool answers that the team id is unknown.
- **Requirements.** The app must be installed in the team and the recipient
  (an AAD object id) must be a member.

The topic is the thread's root message when the conversation id carries
`;messageid=`, otherwise the channel. Errors report only the HTTP status and a
sanitised Graph error code, never the Graph message.
