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
#       [--env-file ~/.claude/channels/teams/.env] [--out ./teams-app.zip] \
#       [--rsc-all-messages] [--command "<title>: <description>"]... \
#       [--message-action "<id>: <title>: <description>"]...
#
# --rsc-all-messages adds the RSC permissions ChannelMessage.Read.Group and ChatMessage.Read.Chat (application) and the
# webApplicationInfo block they need (manifest 1.12 or later; this one is 1.16). Teams then delivers every message in the
# team channels and group chats where the app is installed, not only @mentions, so all of their content reaches the bot
# host: enable it only when you need it. The team or chat owner must consent when the app is installed. An app that is
# already installed in a team must be upgraded there with the new permissions consented (Graph: upgrade the installation
# with a consentedPermissionSet), otherwise the permission is not granted. On its own the plugin still acts only on
# messages that @mention the bot (its mention gate); the permission is what a separate unmentioned-delivery option needs.
#
# --command "<title>: <description>" declares a Teams slash command (repeatable, order kept, max 12). The value is split at
# the first ": "; one leading "/" on the title is dropped, the title must match [A-Za-z0-9][A-Za-z0-9_-]* (max 128 bytes) and
# be unique (case-insensitive); the description is required, max 4000 bytes, without control characters. It adds
# supportsTargetedMessages and one commandLists entry (personal scope only for now; triggers mention and slash) to the bot,
# which raises the manifest to version 1.29 (only when at least one --command is given). The commands arrive as targeted
# messages that are visible only to the sender. Nothing in the server answers a targeted message privately yet, so in a
# channel or group chat the answer would be visible to everyone; that is why commands are limited to 1:1 chats. An app
# that is already installed must be re-uploaded/upgraded in Teams for the change to take effect.
#
# --message-action "<id>: <title>: <description>" adds an entry to the "..." (More actions) menu of a Teams message
# (repeatable, order kept, max 10). The value is split at the first ": " into the id and the rest, the rest at its first
# ": " into the title and the description. The id must match ^[a-z0-9][a-z0-9-]{0,63}$ and be unique; the title is 1..32
# bytes, the description 1..128 bytes, both without control characters and valid UTF-8. It adds a top-level
# composeExtensions member (botId = the App ID, one action command per option, each with an optional "note" textarea);
# the manifest version is not changed. The submitted action reaches Claude as a message_action event. An app that is
# already installed must be re-uploaded/upgraded in Teams for the change to take effect.
set -euo pipefail

APP_ID=""
# Resolved below: explicit --name > TEAMS_BOT_DISPLAY_NAME in the .env > generic
# fallback. NOT hardcoded to any one owner/agent -- the bot name shown in Teams
# must match THIS install's agent (distribution rule: no hardcoded owner names).
APP_NAME=""
ENV_FILE="${HOME}/.claude/channels/teams/.env"
OUT="./teams-app.zip"
RSC_ALL=0
CMD_RAW=()
MA_RAW=()

while [ $# -gt 0 ]; do
  case "$1" in
    --app-id) APP_ID="$2"; shift 2 ;;
    --name) APP_NAME="$2"; shift 2 ;;
    --env-file) ENV_FILE="$2"; shift 2 ;;
    --out) OUT="$2"; shift 2 ;;
    --rsc-all-messages) RSC_ALL=1; shift ;;
    --command)
      [ $# -ge 2 ] || { echo "--command needs a value: '<title>: <description>'" >&2; exit 2; }
      CMD_RAW+=("$2"); shift 2 ;;
    --message-action)
      [ $# -ge 2 ] || { echo "--message-action needs a value: '<id>: <title>: <description>'" >&2; exit 2; }
      MA_RAW+=("$2"); shift 2 ;;
    -h|--help) grep '^#' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *) echo "unknown option: $1" >&2; exit 2 ;;
  esac
done

# Validate every --command before anything is written. Byte semantics: LC_ALL=C.
CMD_TITLES=()
CMD_DESCS=()
CMD_SEEN=" "
validate_commands() {
  local LC_ALL=C
  local v t d key
  if [ "${#CMD_RAW[@]}" -gt 0 ] && ! command -v iconv >/dev/null 2>&1; then
    echo "--command needs iconv to check that each description is valid UTF-8" >&2; exit 2
  fi
  if [ "${#CMD_RAW[@]}" -gt 12 ]; then
    echo "too many --command options (max 12)" >&2; exit 2
  fi
  for v in "${CMD_RAW[@]+"${CMD_RAW[@]}"}"; do
    case "$v" in
      *": "*) ;;
      *) echo "invalid --command (expected '<title>: <description>'): $v" >&2; exit 2 ;;
    esac
    t="${v%%: *}"
    d="${v#*: }"
    t="${t#/}"
    if ! [[ "$t" =~ ^[A-Za-z0-9][A-Za-z0-9_-]*$ ]] || [ "${#t}" -gt 128 ]; then
      echo "invalid --command title: $v" >&2; exit 2
    fi
    if [ -z "$d" ] || [ "${#d}" -gt 4000 ] || [[ "$d" == *[[:cntrl:]]* ]]; then
      echo "invalid --command description (control character or too long)" >&2; exit 2
    fi
    # The byte checks above run under LC_ALL=C, so they let any byte >= 0x80 through: the JSON must be UTF-8.
    if ! printf '%s' "$d" | iconv -f UTF-8 -t UTF-8 >/dev/null 2>&1; then
      echo "invalid --command description (not valid UTF-8)" >&2; exit 2
    fi
    # Not ${t,,}: that needs bash 4, and macOS ships bash 3.2. The title is ASCII here.
    key="$(printf '%s' "$t" | tr ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz)"
    case "$CMD_SEEN" in
      *" $key "*) echo "duplicate --command title: $t" >&2; exit 2 ;;
    esac
    CMD_SEEN="$CMD_SEEN$key "
    d="${d//\\/\\\\}"
    d="${d//\"/\\\"}"
    CMD_TITLES+=("$t")
    CMD_DESCS+=("$d")
  done
}
validate_commands

# Validate every --message-action before anything is written. Byte semantics: LC_ALL=C.
MA_IDS=()
MA_TITLES=()
MA_DESCS=()
MA_SEEN=" "
validate_message_actions() {
  local LC_ALL=C
  local v id rest t d f
  local c1=$'\xc2'"["$'\x80'"-"$'\x9f'"]"
  if [ "${#MA_RAW[@]}" -gt 0 ] && ! command -v iconv >/dev/null 2>&1; then
    echo "--message-action needs iconv to check that each value is valid UTF-8" >&2; exit 2
  fi
  if [ "${#MA_RAW[@]}" -gt 10 ]; then
    echo "too many --message-action options (max 10)" >&2; exit 2
  fi
  for v in "${MA_RAW[@]+"${MA_RAW[@]}"}"; do
    case "$v" in
      *": "*) ;;
      *) echo "invalid --message-action (expected '<id>: <title>: <description>'): $v" >&2; exit 2 ;;
    esac
    id="${v%%: *}"
    rest="${v#*: }"
    case "$rest" in
      *": "*) ;;
      *) echo "invalid --message-action (expected '<id>: <title>: <description>'): $v" >&2; exit 2 ;;
    esac
    t="${rest%%: *}"
    d="${rest#*: }"
    if ! [[ "$id" =~ ^[a-z0-9][a-z0-9-]{0,63}$ ]]; then
      echo "invalid --message-action id (expected ^[a-z0-9][a-z0-9-]{0,63}\$): $id" >&2; exit 2
    fi
    case "$MA_SEEN" in
      *" $id "*) echo "duplicate --message-action id: $id" >&2; exit 2 ;;
    esac
    MA_SEEN="$MA_SEEN$id "
    if [ -z "$t" ] || [ "${#t}" -gt 32 ]; then
      echo "invalid --message-action title (1..32 bytes): $id" >&2; exit 2
    fi
    if [ -z "$d" ] || [ "${#d}" -gt 128 ]; then
      echo "invalid --message-action description (1..128 bytes): $id" >&2; exit 2
    fi
    # [[:cntrl:]] under LC_ALL=C sees only the ASCII controls; the C1 controls U+0080..U+009F are the UTF-8 bytes
    # C2 80..C2 9F.
    if [[ "$t" == *[[:cntrl:]]* ]] || [[ "$d" == *[[:cntrl:]]* ]] || [[ "$t" == *$c1* ]] || [[ "$d" == *$c1* ]]; then
      echo "invalid --message-action (control character in title or description): $id" >&2; exit 2
    fi
    for f in "$t" "$d"; do
      if ! printf '%s' "$f" | iconv -f UTF-8 -t UTF-8 >/dev/null 2>&1; then
        echo "invalid --message-action (not valid UTF-8): $id" >&2; exit 2
      fi
    done
    t="${t//\\/\\\\}"
    t="${t//\"/\\\"}"
    d="${d//\\/\\\\}"
    d="${d//\"/\\\"}"
    MA_IDS+=("$id")
    MA_TITLES+=("$t")
    MA_DESCS+=("$d")
  done
}
validate_message_actions

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

RSC_BLOCK=""
if [ "$RSC_ALL" = 1 ]; then
  RSC_BLOCK=",
  \"webApplicationInfo\": {
    \"id\": \"${APP_ID}\",
    \"resource\": \"https://RscBasedStoreApp\"
  },
  \"authorization\": {
    \"permissions\": {
      \"resourceSpecific\": [
        { \"name\": \"ChannelMessage.Read.Group\", \"type\": \"Application\" },
        { \"name\": \"ChatMessage.Read.Chat\", \"type\": \"Application\" }
      ]
    }
  }"
fi

MA_BLOCK=""
if [ "${#MA_IDS[@]}" -gt 0 ]; then
  MA_ITEMS=""
  for i in "${!MA_IDS[@]}"; do
    [ -z "$MA_ITEMS" ] || MA_ITEMS="$MA_ITEMS,"
    MA_ITEMS="$MA_ITEMS
        {
          \"id\": \"${MA_IDS[$i]}\",
          \"type\": \"action\",
          \"context\": [\"message\"],
          \"title\": \"${MA_TITLES[$i]}\",
          \"description\": \"${MA_DESCS[$i]}\",
          \"fetchTask\": false,
          \"parameters\": [
            { \"name\": \"note\", \"title\": \"Note\", \"description\": \"Optional note for the assistant\", \"inputType\": \"textarea\" }
          ]
        }"
  done
  MA_BLOCK=",
  \"composeExtensions\": [
    {
      \"botId\": \"${APP_ID}\",
      \"commands\": [${MA_ITEMS}
      ]
    }
  ]"
fi

SCHEMA_VER="1.16"
CMD_BLOCK=""
if [ "${#CMD_TITLES[@]}" -gt 0 ]; then
  SCHEMA_VER="1.29"
  CMD_ITEMS=""
  for i in "${!CMD_TITLES[@]}"; do
    [ -z "$CMD_ITEMS" ] || CMD_ITEMS="$CMD_ITEMS,"
    CMD_ITEMS="$CMD_ITEMS
            { \"title\": \"${CMD_TITLES[$i]}\", \"description\": \"${CMD_DESCS[$i]}\" }"
  done
  CMD_BLOCK=",
      \"supportsTargetedMessages\": true,
      \"commandLists\": [
        {
          \"scopes\": [\"personal\"],
          \"triggers\": [\"mention\", \"slash\"],
          \"commands\": [${CMD_ITEMS}
          ]
        }
      ]"
fi

# Minimal valid Teams manifest. `id` is the app's own GUID; reusing the bot App
# ID is allowed and keeps it one-value-simple for the owner.
cat > "$WORK/manifest.json" <<JSON
{
  "\$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v${SCHEMA_VER}/MicrosoftTeams.schema.json",
  "manifestVersion": "${SCHEMA_VER}",
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
      "scopes": ["personal", "team", "groupChat"],
      "supportsFiles": true,
      "isNotificationOnly": false${CMD_BLOCK}
    }
  ],
  "permissions": ["identity", "messageTeamMembers"],
  "validDomains": []${RSC_BLOCK}${MA_BLOCK}
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
if [ "$RSC_ALL" = 1 ]; then
  echo "RSC: requests ChannelMessage.Read.Group and ChatMessage.Read.Chat; the team or chat owner consents at install."
fi
if [ "${#CMD_TITLES[@]}" -gt 0 ]; then
  echo "COMMANDS: declares ${#CMD_TITLES[@]} slash command(s), for 1:1 chats only for now; the manifest is version 1.29. Re-upload or upgrade an already installed app."
fi
if [ "${#MA_IDS[@]}" -gt 0 ]; then
  echo "MESSAGE ACTIONS: declares ${#MA_IDS[@]} message action(s) in the \"...\" menu. Re-upload or upgrade an already installed app, and set TEAMS_MESSAGE_ACTIONS=1 in the plugin .env so the server delivers them."
fi
