#!/usr/bin/env bash
# aifs-exec.sh — Shell wrapper for on-demand AIFS filesystem operations
# (Microsoft OneDrive / SharePoint adapter).
#
# Each invocation starts a fresh Node process, executes one operation,
# and exits. No server, no bridge, no process management.
#
# Usage:
#   aifs-exec.sh <tool_name> [json_args]
#   aifs-exec.sh aifs_read '{"path":"/projects/foo/project.md"}'
#   aifs-exec.sh aifs_list '{"path":"/shared/projects"}'
#   aifs-exec.sh aifs_auth_status
#   aifs-exec.sh --help
#
# Environment:
#   AIFS_CONFIG_PATH  Path to agent-index.json (auto-discovered if not set)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

find_config() {
  if [ -n "${AIFS_CONFIG_PATH:-}" ]; then echo "$AIFS_CONFIG_PATH"; return; fi
  local dir="$SCRIPT_DIR"
  while [ "$dir" != "/" ]; do
    if [ -f "$dir/agent-index.json" ]; then echo "$dir/agent-index.json"; return; fi
    dir="$(dirname "$dir")"
  done
  for dir in "$HOME"/mnt/*/; do
    if [ -f "$dir/agent-index.json" ]; then echo "$dir/agent-index.json"; return; fi
  done
  echo ""
}

find_bundle() {
  if [ -f "$SCRIPT_DIR/aifs-exec.bundle.js" ]; then echo "$SCRIPT_DIR/aifs-exec.bundle.js"; return; fi
  if [ -f "$SCRIPT_DIR/../dist/aifs-exec.bundle.js" ]; then echo "$SCRIPT_DIR/../dist/aifs-exec.bundle.js"; return; fi
  if [ -f "$SCRIPT_DIR/exec.mjs" ]; then echo "$SCRIPT_DIR/exec.mjs"; return; fi
  echo ""
}

if [ "${1:-}" = "--help" ] || [ "${1:-}" = "-h" ] || [ -z "${1:-}" ]; then
  echo "Usage: aifs-exec.sh <tool_name> [json_args]"
  echo ""
  echo "Tools: aifs_read aifs_write aifs_list aifs_exists aifs_stat aifs_delete aifs_copy aifs_auth_status aifs_authenticate"
  echo "Environment: AIFS_CONFIG_PATH (path to agent-index.json, auto-discovered if not set)"
  exit 0
fi

CONFIG_PATH="$(find_config)"
if [ -z "$CONFIG_PATH" ]; then
  echo '{"error":"CONFIG_ERROR","message":"Cannot find agent-index.json. Set AIFS_CONFIG_PATH."}'
  exit 1
fi
export AIFS_CONFIG_PATH="$CONFIG_PATH"

EXEC_PATH="$(find_bundle)"
if [ -z "$EXEC_PATH" ]; then
  echo '{"error":"EXEC_ERROR","message":"Cannot find aifs-exec bundle or source."}'
  exit 1
fi

exec node --no-deprecation --no-warnings "$EXEC_PATH" "$@"
