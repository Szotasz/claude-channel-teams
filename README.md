# claude-channel-teams

Microsoft Teams channel plugin for [Claude Code][cc] — bridge a Teams bot
into a running Claude Code session, in the same style as the official
[Telegram][tg], [Discord][dc], and [iMessage][im] channels.

> **Status: feature-complete v1.** The plugin is in active production
> use. It boots, authenticates Bot Framework inbound, gates by Entra ID
> Object ID, replies via the `reply` tool, drives the full pairing flow
> through the `/teams:access` operator skill, and relays tool-approval
> prompts from Claude Code to Teams as text-reply verdicts. Not yet
> submitted to the official `claude-plugins-official` marketplace —
> community plugin until then.

## What it does

Send a 1:1 chat to your Teams bot from any device signed into Teams. The
message arrives in your Claude Code session as a `<channel>` event.
Claude reads it, does the work against your local files, and replies
through the same Teams thread.

- **Inbound** — Bot Framework activities are JWT-validated by the
  framework, gated against an Entra ID Object ID allowlist, and pushed
  into the session as `notifications/claude/channel`.
- **Outbound** — Claude calls the `reply` tool; the plugin posts back
  through Bot Framework. The outbound side enforces
  `assertAllowedConversation` so a prompt-injected Claude can't be
  tricked into talking to a non-allowlisted conversation.
- **Pairing** — an unknown DM triggers a one-time pairing code. The
  operator reviews and approves it from their Claude Code terminal with
  `/teams:access pair <pair_id> <code>`. Two halves are required (one
  from terminal listing, one from the user's DM) so no single
  prompt-injected message can supply both.
- **Permission relay** — tool-approval prompts from Claude Code are
  relayed to the primary allowlisted Teams conversation. The operator
  replies `yes <id>` or `no <id>` from Teams and the verdict is fed
  back to Claude.

## Prerequisites

- An Entra ID tenant where you can register an application
- An Azure subscription with permission to create an Azure Bot resource
- A reachable HTTPS endpoint (your own reverse proxy or tunnel)
- [Bun][bun] on the machine running Claude Code
- A persistent session manager if you want self-healing (the included
  `examples/systemd/` shows a supervisor + tmux pattern)

See [`docs/installation.md`](docs/installation.md) for the end-to-end
walkthrough and [`docs/azure-setup.md`](docs/azure-setup.md) for the
Azure portal / `az` CLI steps.

## Install

```sh
# 1. Register the plugin's marketplace (a local directory containing this
#    repo as a plugin entry — see docs/installation.md for the layout).
/plugin marketplace add <your-marketplace-name> <local-marketplace-dir>

# 2. Install and reload.
/plugin install teams@<your-marketplace-name>
/reload-plugins
```

Launch Claude Code with the channel attached:

```sh
claude --channels plugin:teams@<your-marketplace-name>
```

Until this plugin is on the official Anthropic allowlist (or you
whitelist it yourself in managed settings under
`allowedChannelPlugins`), it must be loaded with
`--dangerously-load-development-channels` — see the Claude Code
[channels reference][cref] and [`docs/installation.md`](docs/installation.md)
for the headless-friendly managed-settings pattern.

## Security

The threat model and hardening guidance live in
[`docs/security.md`](docs/security.md). Short version:

- Never run an ungated Teams bot — anyone with a routable identity in an
  addressable tenant could DM it.
- The plugin defaults to pairing (closed-door) policy and refuses
  outbound replies to conversations not in the allowlist.
- The `/teams:access` operator skill is the only path that mutates the
  allowlist; it explicitly refuses to act on requests that arrived via
  channel notifications (prompt-injection fence in the SKILL.md
  front-matter).

## Backfill of undelivered channel messages (opt-in)

Teams sometimes does not deliver a channel message to the bot at all. With
`TEAMS_BACKFILL=1` the plugin polls Microsoft Graph (with the app-only token it
already uses for thread history) and feeds every message it never received
through the **same** inbound handler as a live one: same tenant check, same
allowlist, same @mention gate, and the same seen-id store — so a message is
delivered exactly once, whether the live path or the poll got it, also across a
restart. Recovered messages carry `backfilled="true"` in their `<channel>` tag.

It needs the `ChannelMessage.Read.Group` RSC grant in every team it should cover
(`scripts/make-teams-manifest.sh --rsc-all-messages`); a team without it is
skipped for six hours at a time with one stderr line saying so.

| Env var | Default | Meaning |
| --- | --- | --- |
| `TEAMS_BACKFILL` | off | `1` / `true` turns the poll on |
| `TEAMS_BACKFILL_INTERVAL_S` | 180 (60..3600) | seconds between two poll cycles |
| `TEAMS_BACKFILL_GRACE_S` | 90 (30..900) | a message younger than this is left to the live path |
| `TEAMS_BACKFILL_MAX_AGE_S` | 7200 (600..86400) | a message older than this is never backfilled; raised to at least twice the interval plus the grace |
| `TEAMS_DECISION_LOG` | off | `1` / `true` turns the decision log on |
| `TEAMS_DECISION_LOG_MAX_BYTES` | 1048576 (65536..16777216) | size at which `decisions.log` is rotated |

State: `<stateDir>/backfill-state.json` (per-channel watermarks, team ids, 403
back-off) and `<stateDir>/inbound-ids.json` (the seen message ids, shared with
the live path). Both are mode 0600 and only exist with `TEAMS_BACKFILL` on.

Limits, on purpose:

- Only channels the bot already has a stored thread reference for are polled —
  a channel nobody wrote in since the last restart is not.
- Graph returns at most 50 roots per page with `$expand=replies`, and at most
  **200 inline replies per root**; the poll reads at most 5 pages per channel
  per cycle and stops as soon as a page's oldest reply chain is out of range.
  When more reply chains than that are in range, the older ones are skipped
  for that window and a `backfill_skip` / `page_limit` line is logged.
- `TEAMS_BACKFILL_MAX_AGE_S` is a hard cut: a message older than that is never
  backfilled. The floor (twice the interval plus the grace) leaves room for a
  cycle that takes up to one interval. A cycle slower than that (a slow Graph,
  many candidates) can let a missed message age past a max age set close to the
  floor, so keep it well above it (the defaults: 7200 against a floor of 450).
- The seen-id store keeps the last 500 message ids per channel (and 1000
  conversations): a channel with more than 500 messages inside one backfill
  window (normally interval + grace; up to `TEAMS_BACKFILL_MAX_AGE_S` after a
  Graph outage) can see a message delivered twice.
- A message whose handling is cut off by the process exiting is not recovered by
  the poll: its id is stored when handling starts.
- Edits and deletes of an already-backfilled message are not re-delivered
  (they arrive live or not at all), and attachments are not reconstructed —
  the text gets a `[backfilled: N attachment(s) not included]` note instead.

### Decision log (opt-in)

`TEAMS_DECISION_LOG=1` writes one JSON object per line to
`<stateDir>/decisions.log` (mode 0600, rotated to `decisions.log.1` at
`TEAMS_DECISION_LOG_MAX_BYTES`): what happened to every inbound message, so a
lost one can be traced without keeping message content on disk. A line has
exactly these keys — `ts`, `decision`, `reason`, `source`
(`live`/`backfill`), and `message_id` / `channel` (the conversation id without
`;messageid=…`) when they are known. Never message text, display names, AAD
object ids or tokens.

Decisions: `deliver`, `drop`, `ignore`, `held`, `notify_ok`, `notify_failed`,
`backfill_submit`, `backfill_skip`, `backfill_error`. Reason codes:

| decision | reasons |
| --- | --- |
| `deliver` | `ok` — the message cleared every gate |
| `drop` | `duplicate`, `no_aad_object_id`, `tenant_mismatch`, `no_conversation_id`, `sender_not_allowed` |
| `ignore` | `not_mentioned`, `unmentioned_sender_not_allowed` |
| `held` | `filter_held` — the unmentioned filter judged it not for the bot |
| `notify_ok` / `notify_failed` | `sent` / `send_failed` — the MCP notification to the session |
| `backfill_submit` | `submit` — a Graph message was handed to the live handler |
| `backfill_skip` | `no_team_id`, `submit_failed_3x`, `page_limit` |
| `backfill_error` | `forbidden`, `network`, `http_5xx`, `http_error`, `bad_response`, `throttled` |

`http_error` (another 4xx) and `bad_response` (a body that did not parse) stop a
cycle like a 5xx does. Only message activities are logged; reactions, card
clicks, meeting and channel events are not message turns.

## Project layout

```
src/                  # MCP server + Bot Framework wiring + permission relay
skills/               # /teams:access operator skill
tests/                # Bun test suite (39 tests)
docs/                 # Design, security, install, architecture, pairing
examples/             # settings.json snippet + systemd unit + supervisor
scripts/              # Optional helper scripts
.claude-plugin/       # Anthropic plugin manifest
.mcp.json             # MCP server entry consumed by Claude Code
```

## Contributing

PRs welcome on bug fixes, docs, threat-model gaps, dependency hardening,
group / channel-scope support (currently `personal` scope only),
Adaptive Card variants of the permission prompt, or marketplace
submission paperwork.

If you find a corner the canonical Telegram / Discord / iMessage plugins
handle that we don't — flag it. Cross-pollination between channels
makes the whole feature better.

## Licence

MIT. See [`LICENSE`](LICENSE).

[cc]: https://code.claude.com/docs/en/overview
[tg]: https://github.com/anthropics/claude-plugins-official/tree/main/external_plugins/telegram
[dc]: https://github.com/anthropics/claude-plugins-official/tree/main/external_plugins/discord
[im]: https://github.com/anthropics/claude-plugins-official/tree/main/external_plugins/imessage
[cref]: https://code.claude.com/docs/en/channels-reference
[bun]: https://bun.sh
