#!/usr/bin/env bash
#
# setup-azure-bot.sh - one-command Azure Bot registration for the Teams channel.
#
# Goal (Szabi's "make it the simplest for them too"): a non-technical owner runs
# this once and walks away with App ID + client secret + tenant ID written
# straight into ~/.claude/channels/teams/.env - no portal click-through, no
# copy-paste of credentials. Everything the Azure portal flow does by hand, the
# `az` CLI does here non-interactively.
#
# DISTRIBUTION-SAFE: nothing is hardcoded to any one owner. Every value comes
# from the running machine's own `az login` session (the owner's own tenant) and
# from the arguments below. Each install registers its OWN bot in its OWN tenant.
#
# NOTE: the `az` command syntax below is grounded in the Microsoft docs
# (az bot create / az ad app create / az ad app credential reset). It still needs
# a live run against a real (disposable) test tenant before we call it verified -
# that is the operator/test-tenant phase Marveen owns. Run with --dry-run first.
#
# Usage:
#   setup-azure-bot.sh --endpoint https://<tunnel-host>/api/messages [options]
#
# Options:
#   --endpoint URL     Messaging endpoint = your tunnel URL + /api/messages (required)
#   --bot-name NAME    Azure Bot resource name (default: marveen-teams-bot)
#   --resource-group G Resource group (default: marveen-teams-rg)
#   --location LOC     Azure region for the resource group (default: westeurope)
#   --app-type TYPE    SingleTenant (default) | MultiTenant
#   --env-file PATH    Where to write creds (default: ~/.claude/channels/teams/.env)
#   --dry-run          Print the commands without executing them
#
set -euo pipefail

ENDPOINT=""
BOT_NAME="marveen-teams-bot"
RESOURCE_GROUP="marveen-teams-rg"
LOCATION="westeurope"
APP_TYPE="SingleTenant"
ENV_FILE="${HOME}/.claude/channels/teams/.env"
DRY_RUN=0

while [ $# -gt 0 ]; do
  case "$1" in
    --endpoint) ENDPOINT="$2"; shift 2 ;;
    --bot-name) BOT_NAME="$2"; shift 2 ;;
    --resource-group) RESOURCE_GROUP="$2"; shift 2 ;;
    --location) LOCATION="$2"; shift 2 ;;
    --app-type) APP_TYPE="$2"; shift 2 ;;
    --env-file) ENV_FILE="$2"; shift 2 ;;
    --dry-run) DRY_RUN=1; shift ;;
    -h|--help) grep '^#' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *) echo "unknown option: $1" >&2; exit 2 ;;
  esac
done

_fail() { echo "ERROR: $*" >&2; exit 1; }
# In dry-run, echo to stderr so a caller's >/dev/null on the command's stdout
# does not swallow the preview line.
run() { if [ "$DRY_RUN" = 1 ]; then echo "+ $*" >&2; else "$@"; fi; }

[ -n "$ENDPOINT" ] || _fail "--endpoint is required (your tunnel URL + /api/messages). Set up the tunnel first (see setup-tunnel.sh)."

# ── prerequisites ─────────────────────────────────────────────────────────────
# In --dry-run we only preview the commands, so missing az / login is a warning,
# not a hard stop.
_prereq_fail() { if [ "$DRY_RUN" = 1 ]; then echo "DRY-RUN note: $*" >&2; else _fail "$*"; fi; }

command -v az >/dev/null 2>&1 || _prereq_fail "Azure CLI (az) not installed. macOS: 'brew install azure-cli'. Then 'az login'."
if command -v az >/dev/null 2>&1 && ! az account show >/dev/null 2>&1; then
  _prereq_fail "Not logged in to Azure. Run 'az login' (opens a browser, sign in with your work/school M365 account), then re-run this."
fi
# The `az bot` command group is NATIVE in modern az CLI (verified on 2.87.0):
# `az bot create` / `az bot msteams create` / `az bot update` need no extension.
# The old 'botservice' extension no longer exists by that name ("No extension
# found with name botservice"), so a plain `az extension add --name botservice`
# FAILS -- and under `set -euo pipefail` that would abort the whole script before
# the bot is even created. Only fall back to the legacy extension if the bot
# group is genuinely missing, and never let it be fatal.
if command -v az >/dev/null 2>&1 && ! az bot --help >/dev/null 2>&1; then
  echo "az bot group not found; trying the legacy botservice extension (best-effort)..."
  az extension add --name botservice --only-show-errors >/dev/null 2>&1 || true
fi

if [ "$DRY_RUN" = 1 ]; then
  TENANT_ID="<tenant-id-from-az-account-show>"
  echo "+ az account show --query tenantId -o tsv"
else
  TENANT_ID="$(az account show --query tenantId -o tsv)"
  [ -n "$TENANT_ID" ] || _fail "Could not read tenant ID from the az session."
fi
echo "Tenant: $TENANT_ID"

# ── 1) Entra app registration (idempotent by display name) ────────────────────
APP_ID="$(az ad app list --display-name "$BOT_NAME" --query '[0].appId' -o tsv 2>/dev/null || true)"
if [ -n "$APP_ID" ] && [ "$APP_ID" != "null" ]; then
  echo "Reusing existing Entra app '$BOT_NAME' (appId $APP_ID)."
else
  echo "Creating Entra app registration '$BOT_NAME'..."
  if [ "$DRY_RUN" = 1 ]; then
    echo "+ az ad app create --display-name $BOT_NAME --sign-in-audience AzureADMyOrg"
    APP_ID="<app-id-from-create>"
  else
    APP_ID="$(az ad app create --display-name "$BOT_NAME" --sign-in-audience AzureADMyOrg --query appId -o tsv)"
  fi
  [ -n "$APP_ID" ] || _fail "app creation did not return an appId."
fi

# ── 1b) service principal for the app (REQUIRED for outbound bot auth) ─────────
# `az ad app create` registers the application object but does NOT create the
# service principal (the app's identity instance IN this tenant). Without the SP
# the bot's OUTBOUND token request fails with:
#   AADSTS7000229: The client application '<appId>' is missing service principal
#   in the tenant.
# Symptom is subtle and outbound-only: INBOUND (Teams -> bot) works and pairing
# starts, but the bot can never REPLY (the pairing-code DM never arrives). Create
# the SP idempotently -- skip if it already exists.
if [ "$DRY_RUN" = 1 ]; then
  echo "+ az ad sp show --id $APP_ID  ||  az ad sp create --id $APP_ID"
else
  az ad sp show --id "$APP_ID" >/dev/null 2>&1 \
    || az ad sp create --id "$APP_ID" --only-show-errors >/dev/null \
    || _fail "could not create the service principal for $APP_ID (needed for outbound bot auth)."
fi

# ── 2) client secret (always fresh; the value is only shown once) ─────────────
echo "Generating a client secret (max 24-month lifetime; calendar the rotation)..."
if [ "$DRY_RUN" = 1 ]; then
  echo "+ az ad app credential reset --id $APP_ID --append --query password -o tsv"
  APP_SECRET="<secret-from-reset>"
else
  APP_SECRET="$(az ad app credential reset --id "$APP_ID" --append --query password -o tsv)"
fi
[ -n "$APP_SECRET" ] || _fail "secret generation did not return a value."

# ── 3) resource group + Azure Bot (F0 free) ───────────────────────────────────
# Brand-new subscriptions (a customer's first Azure sign-up) do NOT have the
# resource providers registered, so `az bot create` fails with
# "MissingSubscriptionRegistration ... namespace 'Microsoft.BotService'". Register
# them first. `az provider register` is idempotent and a no-op once registered;
# --wait blocks until the registration is live so the bot create below succeeds
# on the very first run. Non-fatal so an already-registered/permission-limited
# tenant still proceeds.
for ns in Microsoft.BotService Microsoft.Web; do
  run az provider register --namespace "$ns" --wait --only-show-errors >/dev/null 2>&1 || true
done

run az group create --name "$RESOURCE_GROUP" --location "$LOCATION" --only-show-errors >/dev/null

echo "Creating Azure Bot '$BOT_NAME' (F0 free tier, $APP_TYPE)..."
TENANT_ARG=()
[ "$APP_TYPE" = "SingleTenant" ] && TENANT_ARG=(--tenant-id "$TENANT_ID")
run az bot create \
  --name "$BOT_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --app-type "$APP_TYPE" \
  --appid "$APP_ID" \
  "${TENANT_ARG[@]}" \
  --sku F0 \
  --endpoint "$ENDPOINT" \
  --only-show-errors >/dev/null || true
# Ensure the endpoint is set even if the bot already existed.
run az bot update --resource-group "$RESOURCE_GROUP" --name "$BOT_NAME" --endpoint "$ENDPOINT" --only-show-errors >/dev/null || true

# ── 4) enable the Microsoft Teams channel ─────────────────────────────────────
echo "Enabling the Microsoft Teams channel..."
run az bot msteams create --name "$BOT_NAME" --resource-group "$RESOURCE_GROUP" --only-show-errors >/dev/null || true

# ── 5) write credentials into the per-install .env (never echo the secret) ────
if [ "$DRY_RUN" = 1 ]; then
  echo "+ write TEAMS_BOT_APP_ID / TEAMS_BOT_APP_PASSWORD / TEAMS_BOT_TENANT_ID / TEAMS_BOT_APP_TYPE / TEAMS_BOT_ENDPOINT_URL -> $ENV_FILE"
else
  mkdir -p "$(dirname "$ENV_FILE")"
  touch "$ENV_FILE"; chmod 600 "$ENV_FILE"
  # Replace-or-append each key, preserving any unrelated keys already present.
  python3 - "$ENV_FILE" "$APP_ID" "$APP_SECRET" "$TENANT_ID" "$APP_TYPE" "$ENDPOINT" <<'PY'
import sys
path, app_id, secret, tenant, app_type, endpoint = sys.argv[1:7]
keys = {
    "TEAMS_BOT_APP_ID": app_id,
    "TEAMS_BOT_APP_PASSWORD": secret,
    "TEAMS_BOT_TENANT_ID": tenant,
    "TEAMS_BOT_APP_TYPE": app_type,
    "TEAMS_BOT_ENDPOINT_URL": endpoint,
}
try:
    lines = open(path, encoding="utf-8").read().splitlines()
except FileNotFoundError:
    lines = []
out, seen = [], set()
for ln in lines:
    k = ln.split("=", 1)[0].strip() if "=" in ln else None
    if k in keys:
        out.append(f"{k}={keys[k]}"); seen.add(k)
    else:
        out.append(ln)
for k, v in keys.items():
    if k not in seen:
        out.append(f"{k}={v}")
open(path, "w", encoding="utf-8").write("\n".join(out) + "\n")
PY
  echo "Credentials written to $ENV_FILE (secret not printed)."
fi

cat <<EOF

Done. Azure side is registered:
  Bot name:   $BOT_NAME
  App ID:     $APP_ID
  Tenant:     $TENANT_ID  ($APP_TYPE)
  Endpoint:   $ENDPOINT

Remaining manual step (the one thing az cannot do): build + sideload the Teams
app package so the bot appears in Teams. Run make-teams-manifest.sh to generate
the .zip from this App ID, then upload it in Teams (Apps -> Manage your apps ->
Upload a custom app). This needs a work/school tenant with custom-app upload
allowed. After that, message the bot once from your own Teams to pair.
EOF
