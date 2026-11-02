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
  cat <<'HELP'
Usage: aifs-exec.sh <tool_name> [json_args]

Tools and their JSON arguments (? = optional). Unknown arguments are
ignored without an error, so a misspelled name silently does nothing.

  aifs_read                Read file content
                             {path}
  aifs_write               Write file content
                             {path, content | content_file | content_stdin, encoding?, if_revision?}
  aifs_list                List directory contents
                             {path, recursive?}
  aifs_exists              Check path existence
                             {path}
  aifs_stat                Get file metadata
                             {path}
  aifs_delete              Delete file or empty directory
                             {path}
  aifs_copy                Copy file
                             {source, destination}
  aifs_auth_status         Check authentication state
                             {}
  aifs_authenticate        Initiate/complete OAuth (PKCE) flow
                             {action?: "start" | "complete", auth_code?}
  aifs_resolve_site        Resolve a SharePoint site URL to site and drive ids (create-org)
                             {site_url}
  aifs_resolve_identity    Resolve an email/UPN/object id to the grantable identity
                             {ref | email | subject | identity | upn | user | member | recipient}  (one of; all mean ref)
  aifs_share               Grant access to a path
                             {path, subject, role}
  aifs_unshare             Remove access from a path
                             {path, subject}
  aifs_get_permissions     List permissions on a path
                             {path}
  aifs_search              Find files and folders by name
                             {scope, name_contains?, type?: "file" | "folder" | "any", max_results?}
  aifs_transfer_ownership  Not implemented on OneDrive (returns NOT_IMPLEMENTED)
                             {path, new_owner}
  aifs_write_batch         Write many files in one process
                             {entries: [{path, content | content_file, encoding?}]}
  aifs_stat_batch          Get metadata for many paths in one process
                             {paths: [...]}

aifs_write content — give exactly one:
  content          Inline string. Carried on the command line, so bounded by
                   the OS argument limit (about 128 KiB).
  content_file     Path to a local file the executor reads directly. No size limit.
  content_stdin    true: read the payload from stdin. No size limit.
  If more than one is given: content, then content_file, then content_stdin wins.
  encoding         "base64": a content_file/content_stdin payload is taken as raw
                   bytes (use for binary files; otherwise it is read as UTF-8
                   text). An inline content string must itself be base64.
  if_revision      Revision from aifs_stat; the write fails with
                   REVISION_CONFLICT if the file has changed since.

Paths: an absolute path (/shared/...) addresses the org's SharePoint library.
id:{item_id}[/rel] is your own OneDrive (retried once on the site library if
not found there); id:{drive_id}:{item_id}[/rel] is that drive. aifs_search's
scope takes a path or an id: anchor. Without name_contains nothing is filtered:
you get up to max_results (default 100, max 1000) items under scope.

Environment: AIFS_CONFIG_PATH (path to agent-index.json, auto-discovered if not set)

Examples:
  aifs-exec.sh aifs_read '{"path":"/projects/foo/project.md"}'
  aifs-exec.sh aifs_list '{"path":"/shared/projects"}'
  aifs-exec.sh aifs_write '{"path":"/shared/foo/big.json","content_file":"/tmp/big.json"}'
  aifs-exec.sh aifs_search '{"scope":"/shared","name_contains":"proposal","type":"file"}'
  aifs-exec.sh aifs_auth_status
HELP
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
