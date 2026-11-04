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
#       [--rsc-all-messages] [--rsc-meeting-events] [--rsc-meeting-transcripts] [--rsc-channel-create] [--command "<title>: <description>"]... \
#       [--prompt "<title>: <prompt>"]...
#
# --rsc-all-messages adds the RSC permissions ChannelMessage.Read.Group and ChatMessage.Read.Chat (application) and the
# webApplicationInfo block they need (manifest 1.12 or later; this one is 1.16). Teams then delivers every message in the
# team channels and group chats where the app is installed, not only @mentions, so all of their content reaches the bot
# host: enable it only when you need it. The team or chat owner must consent when the app is installed. An app that is
# already installed in a team must be upgraded there with the new permissions consented (Graph: upgrade the installation
# with a consentedPermissionSet), otherwise the permission is not granted. On its own the plugin still acts only on
# messages that @mention the bot (its mention gate); the permission is what a separate unmentioned-delivery option needs.
#
# --rsc-meeting-events adds the RSC permissions OnlineMeeting.ReadBasic.Chat and ChannelMeeting.ReadBasic.Group
# (application) and the webApplicationInfo block. Teams then sends the bot meeting start and end events for
# scheduled and channel meetings in the chats and teams where the app is installed; the plugin delivers them only
# with TEAMS_MEETING_EVENTS (see docs/meeting-events.md). A chat member or meeting organizer consents at install; an
# installed app must be upgraded with the new permissions. Not tested in a Teams tenant.
#
# --rsc-meeting-transcripts adds the RSC permission OnlineMeetingTranscript.Read.Chat (application) and the
# webApplicationInfo block, so the bot can read the transcripts of scheduled private chat meetings in the chats where
# the app is installed (not channel meetings); the plugin uses it only with TEAMS_MEETING_TRANSCRIPTS (see
# docs/meeting-transcripts.md). Given without --rsc-meeting-events the script warns on stderr and still writes the
# manifest. Not tested in a Teams tenant.
#
# --rsc-channel-create adds the RSC permission Channel.Create.Group (application), which the create_channel tool needs
# (see docs/create-channel.md). The team owner consents at install; an installed app must be upgraded with the new
# permission. Not tested in a Teams tenant.
#
# --command "<title>: <description>" declares a Teams slash command (repeatable, order kept, max 12). The value is split at
# the first ": "; one leading "/" on the title is dropped, the title must match [A-Za-z0-9][A-Za-z0-9_-]* (max 128 bytes) and
# be unique (case-insensitive); the description is required, max 4000 bytes, without control characters. It adds
# supportsTargetedMessages and one commandLists entry (personal, team, groupChat; triggers mention and slash) to the bot,
# which raises the manifest to version 1.29 (only when at least one --command is given). The commands arrive as targeted
# messages that are visible only to the sender. An app that is already installed must be re-uploaded/upgraded in Teams for
# the change to take effect.
#
# --prompt "<title>: <prompt>" declares a Teams prompt starter (repeatable, order kept). The value is split at the
# first ": "; the title must be 1-128 bytes of valid UTF-8 without control characters or leading/trailing spaces, and
# unique among all --command and --prompt titles (ASCII letters compared case-insensitively); the prompt is 1-4000
# bytes of valid UTF-8 without control characters. --command and --prompt together allow at most 12 entries. Each
# becomes a command { "title", "type": "prompt", "prompt", "description" } (description = the prompt text) in the
# same commandLists entry as the --command entries; with --prompt alone that entry has the same scopes but no
# triggers and no supportsTargetedMessages. Either option raises the manifest to version 1.29. Not tested in a
# Teams tenant.
set -euo pipefail

APP_ID=""
# Resolved below: explicit --name > TEAMS_BOT_DISPLAY_NAME in the .env > generic
# fallback. NOT hardcoded to any one owner/agent -- the bot name shown in Teams
# must match THIS install's agent (distribution rule: no hardcoded owner names).
APP_NAME=""
ENV_FILE="${HOME}/.claude/channels/teams/.env"
OUT="./teams-app.zip"
RSC_ALL=0
RSC_MEETING=0
RSC_TRANSCRIPTS=0
CMD_RAW=()
PROMPT_RAW=()
RSC_CHANNEL_CREATE=0

while [ $# -gt 0 ]; do
  case "$1" in
    --app-id) APP_ID="$2"; shift 2 ;;
    --name) APP_NAME="$2"; shift 2 ;;
    --env-file) ENV_FILE="$2"; shift 2 ;;
    --out) OUT="$2"; shift 2 ;;
    --rsc-all-messages) RSC_ALL=1; shift ;;
    --rsc-meeting-events) RSC_MEETING=1; shift ;;
    --rsc-meeting-transcripts) RSC_TRANSCRIPTS=1; shift ;;
    --rsc-channel-create) RSC_CHANNEL_CREATE=1; shift ;;
    --command)
      [ $# -ge 2 ] || { echo "--command needs a value: '<title>: <description>'" >&2; exit 2; }
      CMD_RAW+=("$2"); shift 2 ;;
    --prompt)
      [ $# -ge 2 ] || { echo "--prompt needs a value: '<title>: <prompt>'" >&2; exit 2; }
      PROMPT_RAW+=("$2"); shift 2 ;;
    -h|--help) grep '^#' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *) echo "unknown option: $1" >&2; exit 2 ;;
  esac
done

# The transcript tool reads only meetings the plugin saw a start or end event for, and those events come with
# --rsc-meeting-events: warn, but still write the manifest (the owner may add the other option later).
if [ "$RSC_TRANSCRIPTS" = 1 ] && [ "$RSC_MEETING" != 1 ]; then
  echo "warning: --rsc-meeting-transcripts without --rsc-meeting-events: the plugin keeps a meeting only from its start or end events, so the transcript tool will find none" >&2
fi

# Validate every --command before anything is written. Byte semantics: LC_ALL=C.
CMD_TITLES=()
CMD_DESCS=()
CMD_SEEN=$'\n'
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
      *$'\n'"$key"$'\n'*) echo "duplicate --command title: $t" >&2; exit 2 ;;
    esac
    CMD_SEEN="$CMD_SEEN$key"$'\n'
    d="${d//\\/\\\\}"
    d="${d//\"/\\\"}"
    CMD_TITLES+=("$t")
    CMD_DESCS+=("$d")
  done
}
validate_commands

# Validate every --prompt before anything is written. Byte semantics: LC_ALL=C.
# A C1 control character (U+0080..U+009F) is the byte pair C2 80..C2 9F in UTF-8; [[:cntrl:]] sees only C0 and DEL.
has_c1() { local LC_ALL=C; [[ "$1" == *$'\xc2'[$'\x80'-$'\x9f']* ]]; }
PROMPT_TITLES=()
PROMPT_TEXTS=()
validate_prompts() {
  local LC_ALL=C
  local v t p key
  if [ "${#PROMPT_RAW[@]}" -gt 0 ] && ! command -v iconv >/dev/null 2>&1; then
    echo "--prompt needs iconv to check that each title and prompt is valid UTF-8" >&2; exit 2
  fi
  if [ $(( ${#CMD_RAW[@]} + ${#PROMPT_RAW[@]} )) -gt 12 ]; then
    echo "too many --command and --prompt options (max 12 together)" >&2; exit 2
  fi
  for v in "${PROMPT_RAW[@]+"${PROMPT_RAW[@]}"}"; do
    case "$v" in
      *": "*) ;;
      *) echo "invalid --prompt (expected '<title>: <prompt>'): $v" >&2; exit 2 ;;
    esac
    t="${v%%: *}"
    p="${v#*: }"
    if [ -z "$t" ] || [ "${#t}" -gt 128 ] || [[ "$t" == *[[:cntrl:]]* ]] || has_c1 "$t" || [[ "$t" == " "* ]] || [[ "$t" == *" " ]]; then
      echo "invalid --prompt title: $v" >&2; exit 2
    fi
    if ! printf '%s' "$t" | iconv -f UTF-8 -t UTF-8 >/dev/null 2>&1; then
      echo "invalid --prompt title (not valid UTF-8): $v" >&2; exit 2
    fi
    if [ -z "$p" ] || [ "${#p}" -gt 4000 ] || [[ "$p" == *[[:cntrl:]]* ]] || has_c1 "$p"; then
      echo "invalid --prompt prompt (control character or too long)" >&2; exit 2
    fi
    if ! printf '%s' "$p" | iconv -f UTF-8 -t UTF-8 >/dev/null 2>&1; then
      echo "invalid --prompt prompt (not valid UTF-8)" >&2; exit 2
    fi
    key="$(printf '%s' "$t" | tr ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz)"
    case "$CMD_SEEN" in
      *$'\n'"$key"$'\n'*) echo "duplicate --prompt title: $t" >&2; exit 2 ;;
    esac
    CMD_SEEN="$CMD_SEEN$key"$'\n'
    t="${t//\\/\\\\}"
    t="${t//\"/\\\"}"
    p="${p//\\/\\\\}"
    p="${p//\"/\\\"}"
    PROMPT_TITLES+=("$t")
    PROMPT_TEXTS+=("$p")
  done
}
validate_prompts

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
if [ "$RSC_ALL" = 1 ] || [ "$RSC_MEETING" = 1 ] || [ "$RSC_TRANSCRIPTS" = 1 ] || [ "$RSC_CHANNEL_CREATE" = 1 ]; then
  RSC_PERMS=""
  if [ "$RSC_ALL" = 1 ]; then
    RSC_PERMS="        { \"name\": \"ChannelMessage.Read.Group\", \"type\": \"Application\" },
        { \"name\": \"ChatMessage.Read.Chat\", \"type\": \"Application\" }"
  fi
  if [ "$RSC_MEETING" = 1 ]; then
    [ -z "$RSC_PERMS" ] || RSC_PERMS="${RSC_PERMS},
"
    RSC_PERMS="${RSC_PERMS}        { \"name\": \"OnlineMeeting.ReadBasic.Chat\", \"type\": \"Application\" },
        { \"name\": \"ChannelMeeting.ReadBasic.Group\", \"type\": \"Application\" }"
  fi
  if [ "$RSC_TRANSCRIPTS" = 1 ]; then
    [ -z "$RSC_PERMS" ] || RSC_PERMS="${RSC_PERMS},
"
    RSC_PERMS="${RSC_PERMS}        { \"name\": \"OnlineMeetingTranscript.Read.Chat\", \"type\": \"Application\" }"
  fi
  if [ "$RSC_CHANNEL_CREATE" = 1 ]; then
    [ -z "$RSC_PERMS" ] || RSC_PERMS="${RSC_PERMS},
"
    RSC_PERMS="${RSC_PERMS}        { \"name\": \"Channel.Create.Group\", \"type\": \"Application\" }"
  fi
  RSC_BLOCK=",
  \"webApplicationInfo\": {
    \"id\": \"${APP_ID}\",
    \"resource\": \"https://RscBasedStoreApp\"
  },
  \"authorization\": {
    \"permissions\": {
      \"resourceSpecific\": [
${RSC_PERMS}
      ]
    }
  }"
fi

SCHEMA_VER="1.16"
CMD_BLOCK=""
if [ "${#CMD_TITLES[@]}" -gt 0 ] || [ "${#PROMPT_TITLES[@]}" -gt 0 ]; then
  SCHEMA_VER="1.29"
  CMD_ITEMS=""
  # Only when non-empty: under set -u, bash before 4.4 treats an empty array as unbound.
  [ "${#CMD_TITLES[@]}" -eq 0 ] || for i in "${!CMD_TITLES[@]}"; do
    [ -z "$CMD_ITEMS" ] || CMD_ITEMS="$CMD_ITEMS,"
    CMD_ITEMS="$CMD_ITEMS
            { \"title\": \"${CMD_TITLES[$i]}\", \"description\": \"${CMD_DESCS[$i]}\" }"
  done
  # Only when non-empty: under set -u, bash before 4.4 treats an empty array as unbound.
  [ "${#PROMPT_TITLES[@]}" -eq 0 ] || for i in "${!PROMPT_TITLES[@]}"; do
    [ -z "$CMD_ITEMS" ] || CMD_ITEMS="$CMD_ITEMS,"
    CMD_ITEMS="$CMD_ITEMS
            { \"title\": \"${PROMPT_TITLES[$i]}\", \"type\": \"prompt\", \"prompt\": \"${PROMPT_TEXTS[$i]}\", \"description\": \"${PROMPT_TEXTS[$i]}\" }"
  done
  if [ "${#CMD_TITLES[@]}" -gt 0 ]; then
    CMD_BLOCK=",
      \"supportsTargetedMessages\": true,
      \"commandLists\": [
        {
          \"scopes\": [\"personal\", \"team\", \"groupChat\"],
          \"triggers\": [\"mention\", \"slash\"],
          \"commands\": [${CMD_ITEMS}
          ]
        }
      ]"
  else
    CMD_BLOCK=",
      \"commandLists\": [
        {
          \"scopes\": [\"personal\", \"team\", \"groupChat\"],
          \"commands\": [${CMD_ITEMS}
          ]
        }
      ]"
  fi
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
  "validDomains": []${RSC_BLOCK}
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
if [ "$RSC_MEETING" = 1 ]; then
  echo "RSC: requests OnlineMeeting.ReadBasic.Chat and ChannelMeeting.ReadBasic.Group for meeting events; a chat member or the meeting organizer consents at install."
fi
if [ "$RSC_TRANSCRIPTS" = 1 ]; then
  echo "RSC: requests OnlineMeetingTranscript.Read.Chat for meeting transcripts; a chat member or the meeting organizer consents at install."
fi
if [ "$RSC_CHANNEL_CREATE" = 1 ]; then
  echo "RSC: requests Channel.Create.Group for create_channel; the team owner consents at install."
fi
if [ "${#CMD_TITLES[@]}" -gt 0 ]; then
  echo "COMMANDS: declares ${#CMD_TITLES[@]} slash command(s); the manifest is version 1.29. Re-upload or upgrade an already installed app."
fi
if [ "${#PROMPT_TITLES[@]}" -gt 0 ]; then
  echo "PROMPTS: declares ${#PROMPT_TITLES[@]} prompt starter(s); the manifest is version 1.29. Re-upload or upgrade an already installed app."
fi
