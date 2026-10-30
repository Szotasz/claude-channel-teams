# Acknowledgement message (`TEAMS_ACK_TEXT`)

After a message is delivered to Claude, if no reply came within
`TEAMS_ACK_DELAY_MS` (default 3000 ms, 0 = right away), the plugin sends the
text in `TEAMS_ACK_TEXT` (for example "Got it, working on it..."). The next
reply overwrites that message in place, so there is no second message.

- One acknowledgement per conversation at a time.
- Not sent for edit / delete events, meeting start / end events, or for channel messages not addressed to the bot.
- A file sent with `send_file` cancels a pending acknowledgement (one already sent is left alone).
- An acknowledgement older than 10 minutes is not overwritten; the reply is sent as a new message.
- If overwriting fails, the reply is sent as a new message. (If Teams applied the overwrite but the call still
  failed, for example on a timeout, the reply then shows twice.)
- The overwritten message stays where the acknowledgement was, so if the person wrote more in the meantime, or a
  file was sent after the acknowledgement, the reply appears above those messages.
- Only a message that reached Claude is acknowledged, and nothing is armed while the plugin shuts down.
- With `TEAMS_STATUS_UPDATES` on, a 1:1 reply that closes an open status stream is that stream's final message, so
  it does not overwrite the acknowledgement; the acknowledgement stays as sent.

## Limits

`TEAMS_ACK_TEXT` is 1-200 characters without control characters; a bad value
turns the acknowledgement off, with a line on stderr. `TEAMS_ACK_DELAY_MS` is an
integer from 0 to 10000; a bad value falls back to 3000, with a line on stderr.

Off by default (no `TEAMS_ACK_TEXT`).

Not measured against a live Teams: overwriting uses updateActivity, as
edit_message does.
