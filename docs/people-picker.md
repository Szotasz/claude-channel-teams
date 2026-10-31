# People picker in cards

## What it is

A people picker is the person search box that Teams renders for an `Input.ChoiceSet` that has `choices.data`
(a `Data.Query` on the Microsoft Graph users dataset). The person types a name and picks from the matches.

## How Claude sends one

Pass a card like this as reply's `card`:

```json
{
  "type": "AdaptiveCard",
  "version": "1.5",
  "body": [
    { "type": "TextBlock", "text": "Who should take this?" },
    {
      "type": "Input.ChoiceSet",
      "id": "assignee",
      "choices": [],
      "choices.data": { "type": "Data.Query", "dataset": "graph.microsoft.com/users?scope=currentContext" },
      "isMultiSelect": false
    }
  ],
  "actions": [{ "type": "Action.Execute", "title": "Assign", "verb": "assign" }]
}
```

The `currentContext` dataset is recommended: only the people in this conversation can be picked.
`"graph.microsoft.com/users"` searches the whole organization. In a 1:1 chat the conversation's members are the person
and the bot, so there the organization dataset is the one that offers anyone else (not measured).

## What comes back

The click arrives as a [Card action] event (see the card action section of docs/design.md). The data JSON holds the
input id with the selected Entra id, or several ids joined by commas (no space) when `isMultiSelect` is true.
For the card above:

```
data: {"assignee":"<entra-id>"}
```

Only the ids come back, no display names: do not guess a name for an id. The ids are what the person picked, so
treat them as user input; they are not checked against the allowlist.

## Not measured against a live Teams

This follows the Microsoft Teams people picker documentation, whose samples use card version 1.2. Not verified live:
that the picker renders in a version 1.5 card (Action.Execute needs 1.4 or later), that an Action.Execute click
carries the input value in its data, and the multi-select separator (a comma with no space, per the documentation).
