# Channel events (`TEAMS_CHANNEL_EVENTS`)

## What is delivered

When a channel is created, renamed, deleted or restored in a team where the bot is installed, the plugin delivers a
channel event to Claude:

- text: `[Channel created] Project X (team Example Team)` (kind is `created`, `renamed`, `deleted` or `restored`; the
  channel name and the `(team ...)` part appear only when Teams sent the names);
- meta: `channel_event` (the kind), `channel_id`, `channel_name` and `team_id` (the last two only when known), plus
  the usual `conversation_id`, `tenant_id` and `ts`. `aad_object_id` and `from_name` are empty: whoever changed the
  channel is not an allowlisted sender, so the event names nobody.

Other conversation updates (members added or removed, team renamed) are not delivered.

## Off by default, and only for listed teams

Set `TEAMS_CHANNEL_EVENTS=1` (or `true`) to turn it on, and list the teams whose channel events are delivered in
`TEAMS_CHANNEL_EVENT_TEAMS`: comma-separated team ids, as Teams sends them in `channelData.team.id` (they look like
`19:...@thread.tacv2`). An empty or missing list delivers no channel event, and there is no wildcard: `*` matches no
team. An event that carries no team id is not delivered. With the flag on, the plugin writes one warning at startup
when the list is empty or holds `*`. The bot gets these events in the teams where it is installed; no extra
permission is needed.

## Names are user content

Anyone who can create a channel in the team sets its name, and the event carries no allowlisted sender. The names are
therefore cleaned of control, line-separator and bidi characters and capped at 100 Unicode code points. The instructions tell
Claude that such an event is not a request.

Not measured against a live Teams: the shapes follow the Microsoft Teams documentation on conversation events.
