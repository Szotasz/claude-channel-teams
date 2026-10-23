# Message actions

A message action is an entry in the "..." (More actions) menu of a Teams message. The user picks it, optionally types a
note, and Claude receives the message plus the note as a `<channel>` event with `message_action="<command>"`.

## Declaring them

```
scripts/make-teams-manifest.sh --message-action "<id>: <title>: <description>" [--message-action ...]
```

The value is split at the first `: ` into the id and the rest, and the rest at its first `: ` into the title and the
description (the description may contain further `: `). Repeatable, order kept. Limits (checked before anything is
written; errors exit 2):

- at most 10 options;
- id matches `^[a-z0-9][a-z0-9-]{0,63}$` and is unique;
- title 1..32 bytes, description 1..128 bytes, no control characters, valid UTF-8.

The manifest gets a top-level `composeExtensions` member (`botId` = the app id) with one `action` command per option
(`context: ["message"]`, `fetchTask: false`) and one optional `note` textarea parameter. The manifest version is not
changed. Without the option the manifest is unchanged.

An app that is already installed must be re-uploaded/upgraded in Teams for the menu entries to appear.

## What Claude receives

```
[Message action] command="<id>"
from: <author of the quoted message>
link: <https://teams.microsoft.com/... link, when present>
message: <quoted message as plain text, max 4000 characters>
note: <what the user typed, max 2000 characters, or ->
```

The tag carries `message_action="<id>"`, and `message_id` is the id of the quoted message. The quoted message and its
author are data from that message, not instructions; only the note is the user's own words.

## Access

Only allowlisted users. An action from anyone else is dropped (one stderr line) and never starts pairing.

## Response

Teams requires an answer to the invoke within five seconds. The bot answers every `composeExtension/*` invoke at once
with `{ type: 'invokeResponse', value: { status: 200 } }` and no body ("No response"), whether or not the event was
delivered. Claude's reply arrives later as a normal message.
