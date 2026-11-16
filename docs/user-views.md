# User-specific card views (`user_cards`)

The `reply` tool can show different people different versions of one Adaptive Card. Each person listed in `user_cards`
sees their own card; everyone else sees the base card (`card`).

## The reply argument

`user_cards` is an optional array of up to 60 entries `{ aad_object_id, card }`. It needs `card` (the base card), and the
base card and every user card must have version `1.4` or `1.5`. Each `aad_object_id` must be an AAD object id and may be
listed only once (case-insensitive). Every card is validated like any other card.

## How it works

Teams does this with the card's `refresh` property (Universal Actions). The plugin adds the refresh itself, after
validation, to the base card and to every user card: an `Action.Execute` with the reserved verb `teams-plugin-user-view`,
`data.view` = a random id, and `userIds` = the user MRIs of the listed people. A card from Claude may not contain a
`refresh` or the reserved verb; it is refused.

The client of each listed person sends an `adaptiveCard/action` invoke, and the bot answers it with the card stored for
that person. The invoke is never an event to Claude.

## Who sees what

- A listed person sees their own card, as long as they are on the allowlist; once removed, they get the base card.
- Everyone else sees the base card (also if their client asks for a refresh).

## Who can be listed

A person is known only after they wrote to the bot since it started (the plugin learns the user MRI from their
activity). Otherwise the reply fails, names them, and nothing is sent.

## Limits

- 60 people per card.
- 500 stored views in memory; the oldest is dropped.
- 1000 known people; the oldest is dropped.

Views are kept in memory only. After a restart a refresh is answered with 'Ignored.' and everyone sees the base card
until Claude sends the card again.

Not measured against a live Teams: the wire shape follows the Microsoft Teams Universal Actions documentation, whose samples use version 1.4 while its text says 1.5; both are accepted.
