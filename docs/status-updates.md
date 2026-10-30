# Status updates in 1:1 chats (`status_update`)

The `status_update` tool shows a short line about what Claude is doing (for example "Searching the tender list...") using Teams streaming "informative updates". The user sees the line in Teams' progress indicator; Claude's next reply replaces it and closes the stream.

## Enabling

Set `TEAMS_STATUS_UPDATES` to `1` or `true`. It is off by default; without it the tool is not offered.

## Limits

- 1:1 chats only (other conversations are refused).
- At most 1000 characters of plain text per update.
- At most one request per second, counting the reply that closes the stream; a call waits for its slot. Updates and
  the reply for one chat go out one after another, so no update starts while the reply is being sent.
- A stream lasts 2 minutes, counted from the request that opened it. After that, and after Teams has closed the stream
  (for example the user stopped it), no more updates are shown for this reply; just send the reply as usual.
- The typing indicator stops while a stream is open (Teams shows no typing animation then).

## When the stream is already closed

If Teams answers 403 `ContentStreamNotAllowed` (for example the user stopped the stream, or it timed out), the reply is sent again as a new message so it is not lost.

## Not measured

Not measured against a live Teams: the plugin sends through the stored conversation reference, so the updates go out as replies to the user's message (the documentation shows a plain post to the conversation); whether the final reply replaces the status line in the same bubble, and the 403 fallback, were not measured.
