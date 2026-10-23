# Mentions in replies (`mentions`)

## The reply argument

`reply` takes an optional `mentions` array, at most 10 entries. Each entry is one of:

- a person: `{ "type": "user", "aad_object_id": "<AAD object id>", "name": "Ann" }`
- a team tag: `{ "type": "tag", "tag": "Finance" }`

The message text must contain `<at>name</at>` for each entry (the tag's name for a tag). Each person and each tag may be
listed once, and two entries may not use the same name. Nothing is sent if any entry is invalid. Keep names plain: the
text is sent as markdown, and a name with markdown characters (`*`, `_`, a backtick) may no longer match its entity
(not measured).

## People

A person must be a member of the conversation or team. This is checked with the Bot Framework member lookup
(`TeamsInfo.getMember`) just before sending. If someone is not a member, the reply fails, names them, and nothing is
sent. A lookup that fails for another reason also fails the reply.

## Tags

Tags must be listed in `TEAMS_MENTION_TAGS="Name=id,Name2=id2"`. One bad entry (no `=`, an invalid name, an id that is
not base64-like, a repeated name) turns tag mentions off; a line is written to stderr at start-up.

Tag mentions work only in standard team channels: not private or shared channels, not group chats or 1:1 chats. The
plugin refuses a tag mention outside a team channel; it cannot tell a private or shared channel from a standard one, so
there it is Teams that refuses the message, and the reply fails (not measured). Teams allows at most 10 tags per message
and 2 messages with tag mentions per 5 seconds in a thread, and refuses more.

### Getting a tag id

Once per tag: `GET https://graph.microsoft.com/v1.0/teams/{team-id}/tags` with an app token that has
`TeamworkTag.Read.All` (or signed in, e.g. in Graph Explorer, with `TeamworkTag.Read`). Take the `id` of the tag and put
it into the list as returned.

## Not measured against a live Teams

The wire shape follows the Microsoft Teams documentation on channel and group conversations. Not measured: that the
member lookup accepts an AAD object id in a channel conversation, what it returns for a non-member, and whether the tag
id from Graph is used as returned.
