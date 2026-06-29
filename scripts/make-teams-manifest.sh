#!/usr/bin/env bash
#
# make-teams-manifest.sh - generate the Teams app package (.zip) for the bot.
#
# This is the one step `az` cannot do: Teams needs an app manifest that
# references the bot's App ID, packaged with two icons, then sideloaded into
# Teams (Apps -> Manage your apps -> Upload a custom app). We generate the .zip
# so the owner only has to upload it.
#
# Reads TEAMS_BOT_APP_ID from the env file (or --app-id). Writes teams-app.zip.
#
# Usage:
#   make-teams-manifest.sh [--app-id <guid>] [--name "Marveen"] \
#       [--env-file ~/.claude/channels/teams/.env] [--out ./teams-app.zip]
set -euo pipefail

APP_ID=""
# Resolved below: explicit --name > TEAMS_BOT_DISPLAY_NAME in the .env > generic
# fallback. NOT hardcoded to any one owner/agent -- the bot name shown in Teams
# must match THIS install's agent (distribution rule: no hardcoded owner names).
APP_NAME=""
ENV_FILE="${HOME}/.claude/channels/teams/.env"
OUT="./teams-app.zip"

while [ $# -gt 0 ]; do
  case "$1" in
    --app-id) APP_ID="$2"; shift 2 ;;
    --name) APP_NAME="$2"; shift 2 ;;
    --env-file) ENV_FILE="$2"; shift 2 ;;
    --out) OUT="$2"; shift 2 ;;
    -h|--help) grep '^#' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *) echo "unknown option: $1" >&2; exit 2 ;;
  esac
done

_fail() { echo "ERROR: $*" >&2; exit 1; }

if [ -z "$APP_ID" ] && [ -f "$ENV_FILE" ]; then
  APP_ID="$(grep -E '^TEAMS_BOT_APP_ID=' "$ENV_FILE" | head -1 | cut -d= -f2- || true)"
fi
[ -n "$APP_ID" ] || _fail "no App ID (pass --app-id or set TEAMS_BOT_APP_ID in $ENV_FILE). Run setup-azure-bot.sh first."

# Name-sync: if --name was not given, take TEAMS_BOT_DISPLAY_NAME from the .env
# (the launcher writes the agent's displayName there) so the bot name in Teams
# matches the agent. Generic fallback only if nothing is configured.
if [ -z "$APP_NAME" ] && [ -f "$ENV_FILE" ]; then
  APP_NAME="$(grep -E '^TEAMS_BOT_DISPLAY_NAME=' "$ENV_FILE" | head -1 | cut -d= -f2- || true)"
fi
[ -n "$APP_NAME" ] || APP_NAME="Assistant"

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

# Minimal valid Teams manifest. `id` is the app's own GUID; reusing the bot App
# ID is allowed and keeps it one-value-simple for the owner.
cat > "$WORK/manifest.json" <<JSON
{
  "\$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.16/MicrosoftTeams.schema.json",
  "manifestVersion": "1.16",
  "version": "1.0.0",
  "id": "${APP_ID}",
  "developer": {
    "name": "${APP_NAME}",
    "websiteUrl": "https://example.com",
    "privacyUrl": "https://example.com/privacy",
    "termsOfUseUrl": "https://example.com/terms"
  },
  "name": { "short": "${APP_NAME}", "full": "${APP_NAME} assistant" },
  "description": { "short": "${APP_NAME} assistant bot", "full": "${APP_NAME} - personal assistant bot bridged to Claude Code." },
  "icons": { "color": "color.png", "outline": "outline.png" },
  "accentColor": "#2A6DF4",
  "bots": [
    {
      "botId": "${APP_ID}",
      "scopes": ["personal"],
      "supportsFiles": true,
      "isNotificationOnly": false
    }
  ],
  "permissions": ["identity", "messageTeamMembers"],
  "validDomains": []
}
JSON

# Placeholder icons (192x192 color, 32x32 outline). If ImageMagick/sips is around
# we make solid PNGs; otherwise we emit a 1x1 and warn (Teams accepts it for dev).
make_png() { # size outfile
  local size="$1" out="$2"
  if command -v sips >/dev/null 2>&1; then
    # Build from a tiny base via sips is awkward; fall back to python if present.
    :
  fi
  if command -v python3 >/dev/null 2>&1 && python3 -c 'import PIL' >/dev/null 2>&1; then
    python3 - "$size" "$out" <<'PY'
import sys
from PIL import Image
size=int(sys.argv[1]); out=sys.argv[2]
Image.new("RGBA",(size,size),(42,109,244,255)).save(out)
PY
  else
    # 1x1 transparent PNG (base64) - valid placeholder; replace with real icons later.
    base64 -d > "$out" <<'B64'
iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==
B64
    echo "WARN: Pillow not available; wrote a 1x1 placeholder $out. Replace with a real ${size}px icon before publishing." >&2
  fi
}

make_png 192 "$WORK/color.png"
make_png 32  "$WORK/outline.png"

( cd "$WORK" && zip -q -r "teams-app.zip" manifest.json color.png outline.png )
mkdir -p "$(dirname "$OUT")"
cp "$WORK/teams-app.zip" "$OUT"

cat <<EOF
Wrote $OUT (App ID ${APP_ID}).
Upload it in Teams: Apps -> Manage your apps -> Upload a custom app -> pick this .zip.
(Needs a work/school tenant with custom-app upload allowed.)
EOF
