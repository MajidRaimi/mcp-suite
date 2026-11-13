#!/usr/bin/env bash
set -euo pipefail

: "${MAC_HOST:?set MAC_HOST to the SSH host name or address of the Mac}"
MAC_USER="${MAC_USER:-$USER}"
MAC_REPO="${MAC_REPO:-\$HOME/mcp-suite}"
MAC_UV="${MAC_UV:-\$HOME/.local/bin/uv}"

exec ssh -T -o BatchMode=yes -o ServerAliveInterval=30 "${MAC_USER}@${MAC_HOST}" \
  "exec \"${MAC_UV}\" run --quiet --no-dev --project \"${MAC_REPO}/servers/office\" office-mcp --transport stdio"
