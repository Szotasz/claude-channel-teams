#!/usr/bin/env bash
#
# setup-tunnel.sh - stand up a STABLE public HTTPS endpoint in front of the local
# Teams listener (default :3978), so Azure Bot Service can POST inbound Activities.
#
# Teams has no outbound-only receive path, so a public endpoint is unavoidable.
# We prefer the NO-DOMAIN option so a non-technical owner needs no domain or DNS:
#
#   default (domain-free):  Tailscale Funnel  -> stable https://<host>.<tailnet>.ts.net
#   advanced (own domain):  cloudflared named tunnel -> https://teams-bot.<domain>
#
# A trycloudflare.com QUICK tunnel is intentionally NOT offered: its URL is
# ephemeral and changes on restart, which breaks the statically-registered Azure
# messaging endpoint (confirmed in teams-channel-research.md, sections 2 and 6).
#
# Usage:
#   setup-tunnel.sh [--port 3978] [--mode funnel|cloudflared] [--hostname teams-bot.<domain>]
#
# Prints the public messaging-endpoint URL (……/api/messages) on success.
set -euo pipefail

PORT=3978
MODE="funnel"
HOSTNAME_ARG=""

while [ $# -gt 0 ]; do
  case "$1" in
    --port) PORT="$2"; shift 2 ;;
    --mode) MODE="$2"; shift 2 ;;
    --hostname) HOSTNAME_ARG="$2"; shift 2 ;;
    -h|--help) grep '^#' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *) echo "unknown option: $1" >&2; exit 2 ;;
  esac
done

_fail() { echo "ERROR: $*" >&2; exit 1; }

case "$MODE" in
  funnel)
    command -v tailscale >/dev/null 2>&1 || _fail "tailscale not installed. macOS: 'brew install tailscale' (or the App Store app), then 'tailscale up'."
    tailscale status >/dev/null 2>&1 || _fail "tailscale is not connected. Run 'tailscale up' and sign in, then re-run."
    # Funnel must be enabled for the tailnet (admin console: ACL 'nodeAttrs' funnel
    # + the node needs HTTPS). Starting funnel will error clearly if not allowed.
    echo "Starting Tailscale Funnel for localhost:${PORT} in the background..."
    tailscale funnel --bg "${PORT}" \
      || _fail "Funnel failed. Enable Funnel for this tailnet (admin console -> Access controls -> add a funnel nodeAttr) and enable HTTPS certificates, then retry."
    HOST="$(tailscale status --json 2>/dev/null | python3 -c 'import sys,json; print(json.load(sys.stdin).get("Self",{}).get("DNSName","").rstrip("."))')"
    [ -n "$HOST" ] || _fail "Could not read this node's tailnet DNS name."
    echo
    echo "Public messaging endpoint:"
    echo "  https://${HOST}/api/messages"
    ;;
  cloudflared)
    command -v cloudflared >/dev/null 2>&1 || _fail "cloudflared not installed. macOS: 'brew install cloudflared'."
    [ -n "$HOSTNAME_ARG" ] || _fail "--hostname is required for cloudflared mode (e.g. teams-bot.yourdomain.com on a Cloudflare-managed zone)."
    echo "cloudflared named-tunnel setup is the advanced path. Outline:"
    echo "  1) cloudflared tunnel login"
    echo "  2) cloudflared tunnel create marveen-teams"
    echo "  3) route DNS: cloudflared tunnel route dns marveen-teams ${HOSTNAME_ARG}"
    echo "  4) run: cloudflared tunnel run --url http://localhost:${PORT} marveen-teams"
    echo "     (install as a launchd service for always-on)"
    echo
    echo "Public messaging endpoint:"
    echo "  https://${HOSTNAME_ARG}/api/messages"
    ;;
  *)
    _fail "unknown --mode '$MODE' (use 'funnel' or 'cloudflared')."
    ;;
esac
