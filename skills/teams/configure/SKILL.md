---
name: configure
description: Guided wizard to connect the Microsoft Teams channel for a non-technical operator (tunnel, one-command Azure bot registration, Teams app sideload, pairing). Use when the operator wants to set up Teams, asks how to set this up or who can reach me, or wants to check channel status.
user-invocable: true
allowed-tools:
  - Read
  - Write
  - Bash(ls *)
  - Bash(mkdir *)
  - Bash(cat *)
  - Bash(bash *)
  - Bash(command -v *)
  - Bash(az account show)
  - Bash(tailscale status)
---

# Microsoft Teams Channel Setup (/teams:configure)

A guided wizard that connects a **Microsoft Teams bot** to Claude Code. Unlike
Telegram (long-poll) or Slack (Socket Mode), Teams has **no outbound-only
receive path**: the Azure Bot Service delivers each message as an HTTP POST to a
public `messagingEndpoint`, so a stable tunnel in front of the plugin's local
listener is required (see `docs/architecture.md`).

The heavy lifting lives in three helper scripts shipped in the plugin's
`scripts/` directory. This skill runs them **in order**, checks the
prerequisites first, and walks the operator through the one manual step. The
goal: a non-technical operator (for example a clinic admin) can wire Teams in
about **15 minutes** by following the steps, copying almost nothing by hand.

State lives in `~/.claude/channels/teams/`. The credential is the Azure client
secret, which the Azure script writes straight into `.env`; the operator never
copies a secret.

Reference the helper scripts as `${CLAUDE_PLUGIN_ROOT}/scripts/<name>.sh`. This
skill complements `/teams:access` (which manages who is allowed) and the
reference docs under `docs/` (`installation.md`, `azure-setup.md`, `pairing.md`,
`security.md`).

Arguments passed: `$ARGUMENTS`

- no args, or `status`: run the **status** check (below), then offer the wizard.
- `wizard`: run the **guided setup** (steps 0 to 4 below).
- `set` / `clear`: manual credential edit (advanced fallback, below).

---

## The setup wizard (steps 0 to 4, run in order)

Tell the operator up front what to expect: about 15 minutes, four steps; you
(the assistant) run the scripts and report each result; the only things they do
by hand are enabling a couple of prerequisites, uploading one file to Teams, and
sending one message.

### Step 0: prerequisites (check before anything else)

Run the probes below and, for each gap, give the operator the **exact** fix. Do
not move past a hard-missing prerequisite.

1. **Work/school M365 tenant.** Teams bot development and sideload require a
   Microsoft 365 **work or school** account; a personal Microsoft account does
   **not** work. If the operator only has a personal account, point them to the
   free **Microsoft 365 Developer Program**
   (https://developer.microsoft.com/microsoft-365/dev-program), which gives an
   instant test tenant.
2. **Admin rights for custom-app upload.** Sideloading needs the tenant to allow
   custom-app upload (Teams admin center, Setup policies, turn on "Upload custom
   apps"). If the operator is the tenant admin they can flip it; otherwise their
   IT admin must.
3. **Azure CLI.** Probe `command -v az`. Missing means `brew install azure-cli`,
   then `az login` (opens a browser, sign in with the work/school account). The
   Azure script also needs the `botservice` extension, which it installs on its
   own.
4. **Tailscale + Funnel** (for the default domain-free tunnel). Probe
   `command -v tailscale` and `tailscale status`. Missing means
   `brew install tailscale` then `tailscale up`. Funnel must be enabled for the
   tailnet (admin console, Access controls, add a funnel nodeAttr and enable
   HTTPS certificates). If the operator has their own domain and prefers it,
   they can skip Tailscale and use the cloudflared path in step 1 instead.

Report a green/blocked line per item. Continue only when items 1 to 4 (or the
cloudflared alternative for item 4) are satisfied.

### Step 1: public endpoint (tunnel)

Run the tunnel script (Tailscale Funnel is the default, domain-free path). The
tunnel must forward to the plugin's listener port, which is `TEAMS_PLUGIN_PORT`
(default `3979`, read by `src/config.ts`); pass `--port` so the forward target
matches:

```
bash "${CLAUDE_PLUGIN_ROOT}/scripts/setup-tunnel.sh" --port 3979
```

It starts Funnel in the background and prints the public messaging endpoint:

```
https://<host>.<tailnet>.ts.net/api/messages
```

Capture that URL; step 2 needs it. Own-domain (advanced):
`setup-tunnel.sh --mode cloudflared --hostname teams-bot.<domain>` prints the
cloudflared outline and the `https://teams-bot.<domain>/api/messages` endpoint.

### Step 2: register the Azure bot (one command, auto-writes `.env`)

Preview first with `--dry-run` so the operator sees exactly what will run, no
changes made:

```
bash "${CLAUDE_PLUGIN_ROOT}/scripts/setup-azure-bot.sh" --endpoint <step-1-url> --dry-run
```

Then run it for real (drop `--dry-run`):

```
bash "${CLAUDE_PLUGIN_ROOT}/scripts/setup-azure-bot.sh" --endpoint <step-1-url>
```

This creates the Entra app, a client secret, the Azure Bot (F0 free tier), and
the Teams channel, then writes `TEAMS_BOT_APP_ID`, `TEAMS_BOT_APP_PASSWORD`,
`TEAMS_BOT_TENANT_ID`, `TEAMS_BOT_APP_TYPE`, and `TEAMS_BOT_ENDPOINT_URL`
straight into `~/.claude/channels/teams/.env`. The operator copies nothing and
the secret is never printed. Defaults: bot name `marveen-teams-bot`, resource
group `marveen-teams-rg`, region `westeurope`, app type `SingleTenant` (override
with `--bot-name` / `--resource-group` / `--location` / `--app-type`). It is
idempotent, so re-running reuses the existing app.

### Step 3: Teams app package and sideload (the one manual step)

Generate the app package:

```
bash "${CLAUDE_PLUGIN_ROOT}/scripts/make-teams-manifest.sh"
```

It reads `TEAMS_BOT_APP_ID` from `.env` and writes `./teams-app.zip` (optional
`--name "<assistant name>"`, `--out <path>`). Then the operator uploads it in
Teams by hand, which is the one thing `az` cannot do:

> Teams, then **Apps**, then **Manage your apps**, then **Upload a custom app**,
> then pick `teams-app.zip`.

Walk them through it click by click. This needs the work/school tenant with
custom-app upload allowed (step 0, item 2). The placeholder icons are fine for
now; swap in real ones before any wider publish.

### Step 4: pair and lock down

Pairing is two-factor by design (see `docs/pairing.md`): the operator must
supply both a `pair_id` they read from their own terminal and a `code` the user
reports from their DM. This blocks an attacker who DMs the bot from getting
approved via a prompt-injected "approve the pending one".

1. From the operator's **own** Teams, open a 1:1 chat with the bot and send any
   message. The bot replies with a 6-char `code` and tells them to ask the
   operator to approve it.
2. In the terminal, run `/teams:access` to see the pending row with its
   `pair_id` and the same `code`.
3. Approve with **both** halves: `/teams:access pair <pair_id> <code>`. On
   success the user is added to the allowlist and the bot DMs them "Paired".

The allowlist **is** the lockdown: only approved AAD object ids get through, so
there is no separate "allowlist mode" to switch on. To add more people, repeat
the pairing; to remove someone, `/teams:access revoke <aad_object_id>`. To
freeze the allowlist so the running process can never change it, set
`TEAMS_ACCESS_MODE=static` in `.env` (pairing is then disabled).

---

## No args: status

Read the state and give the operator a complete picture, then offer to run the
wizard:

1. **Credential status.** Read `~/.claude/channels/teams/.env`. Report which of
   `TEAMS_BOT_APP_ID`, `TEAMS_BOT_APP_PASSWORD`, `TEAMS_BOT_TENANT_ID` are set
   (present or absent only, never print the secret value). Any missing means the
   plugin cannot boot (these three are required by `src/config.ts`).
2. **Endpoint and listener.** Report `TEAMS_BOT_ENDPOINT_URL` (the tunnel
   endpoint registered in Azure; diagnostic only) and the local listener port
   (`TEAMS_PLUGIN_PORT`, default `3979`).
3. **Access.** Defer to `/teams:access` for the live allowlist and pending
   pairings (it drives the plugin's MCP tools). Mention `TEAMS_ACCESS_MODE` if
   set to `static`.
4. **What next.** The concrete next step based on state. If nothing is
   configured, offer `/teams:configure wizard`.

---

## Manual / fallback commands

For operators who registered the bot in the Azure portal themselves, or are
moving an existing bot, the credentials can be set by hand instead of via
step 2. The manual portal/CLI flow is documented in `docs/azure-setup.md`.

### `set <app-id> <tenant-id>`: record the bot identity by hand

1. `mkdir -p ~/.claude/channels/teams`
2. Read existing `.env`; set `TEAMS_BOT_APP_ID=<app-id>` and
   `TEAMS_BOT_TENANT_ID=<tenant-id>`, preserving other keys. Write back, no
   quotes. `chmod 600` the file.
3. Prompt the operator to paste the client **secret** so it can be written to
   `TEAMS_BOT_APP_PASSWORD` (write it to `.env`, never echo it back, never to
   chat).
4. Optionally set `TEAMS_BOT_ENDPOINT_URL=<https tunnel url>`.
5. Restart the channel session (or `/reload-plugins`).

### `clear`: remove the credentials

Tell the operator this takes the channel offline until re-registered. With their
confirmation: blank `TEAMS_BOT_APP_ID`, `TEAMS_BOT_APP_PASSWORD`,
`TEAMS_BOT_TENANT_ID` in `.env` (or delete the file). Keep `allowlist.json` and
`pending.json`. The Azure resource itself is untouched; delete it in the portal
(or with `az`) if you want it gone.

---

## Implementation notes

- The helper scripts live in the plugin's `scripts/` directory; reference them
  via `${CLAUDE_PLUGIN_ROOT}/scripts/`. Do not reimplement their logic here; this
  skill orchestrates and explains them.
- The plugin reads `.env` once at boot, so credential or endpoint changes need a
  session restart or `/reload-plugins`. Access state (`allowlist.json` /
  `pending.json`) is managed live via `/teams:access`.
- `TEAMS_BOT_APP_PASSWORD` is the Azure client secret. Never paste it into chat;
  a silent `401` from Azure usually means the secret expired (24-month cap);
  re-run `setup-azure-bot.sh` or re-issue the secret in Entra.
- The tunnel must stay up for inbound to work. Run it as a supervised service
  (launchd agent) so it restarts on its own; if it drops, Azure cannot reach the
  endpoint and inbound stops.
- A missing channels dir means not configured, not an error.
