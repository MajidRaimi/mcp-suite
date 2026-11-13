#!/usr/bin/env bash
set -euo pipefail

: "${MAC_HOST:?set MAC_HOST to the SSH host name or address of the Mac}"
MAC_USER="${MAC_USER:-$USER}"
MAC_REPO="${MAC_REPO:-\$HOME/mcp-suite}"
MAC_UV="${MAC_UV:-\$HOME/.local/bin/uv}"
LOCAL_PORT="${LOCAL_PORT:-8769}"
REMOTE_PORT="${REMOTE_PORT:-8769}"
START_REMOTE="${START_REMOTE:-1}"

ssh_options=(
  -o BatchMode=yes
  -o ExitOnForwardFailure=yes
  -o ServerAliveInterval=30
  -o ServerAliveCountMax=3
  -L "127.0.0.1:${LOCAL_PORT}:127.0.0.1:${REMOTE_PORT}"
)

if [[ "${START_REMOTE}" != "1" ]]; then
  exec ssh -N "${ssh_options[@]}" "${MAC_USER}@${MAC_HOST}"
fi

exec ssh -tt "${ssh_options[@]}" "${MAC_USER}@${MAC_HOST}" \
  "exec \"${MAC_UV}\" run --quiet --no-dev --project \"${MAC_REPO}/servers/office\" office-mcp --transport streamable-http --host 127.0.0.1 --port ${REMOTE_PORT}"
