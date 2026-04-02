#!/usr/bin/env bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$SCRIPT_DIR/bot_common.sh"

if [[ -f "$PID_FILE" ]]; then
  daemon_pid="$(tr -d '[:space:]' < "$PID_FILE")"
  if is_daemon_pid "$daemon_pid"; then
    echo "running: PID $daemon_pid"
    exit 0
  fi
fi

if mapfile -t valid_pids < <(collect_daemon_pids) && [[ ${#valid_pids[@]} -gt 0 ]]; then
  if [[ ${#valid_pids[@]} -eq 1 ]]; then
    printf '%s\n' "${valid_pids[0]}" > "$PID_FILE"
    echo "running: PID ${valid_pids[0]}"
    exit 0
  fi

  echo "running: multiple PIDs ${valid_pids[*]}"
  exit 1
fi

echo "not running"
exit 1
