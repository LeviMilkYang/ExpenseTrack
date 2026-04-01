#!/usr/bin/env bash

set -euo pipefail

BASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_FILE="$BASE_DIR/telegram_expense_daemon.pid"
DAEMON_SCRIPT="$BASE_DIR/telegram_expense_daemon.py"
STATE_FILE="$BASE_DIR/telegram_bot_config.json"

is_daemon_pid() {
  local pid="$1"
  [[ -n "$pid" ]] || return 1
  [[ "$pid" =~ ^[0-9]+$ ]] || return 1
  kill -0 "$pid" 2>/dev/null || return 1
  [[ -r "/proc/$pid/cmdline" ]] || return 1

  local cmdline
  cmdline="$(tr '\0' ' ' < "/proc/$pid/cmdline" 2>/dev/null || true)"
  [[ "$cmdline" == *"$DAEMON_SCRIPT"* ]]
}

find_daemon_pids() {
  if command -v pgrep >/dev/null 2>&1; then
    pgrep -f "$DAEMON_SCRIPT.*--state-file $STATE_FILE|$DAEMON_SCRIPT" || true
    return
  fi

  ps -ef | grep "$DAEMON_SCRIPT" | grep -v grep | awk '{print $2}' || true
}

if [[ -f "$PID_FILE" ]]; then
  daemon_pid="$(tr -d '[:space:]' < "$PID_FILE")"
  if is_daemon_pid "$daemon_pid"; then
    echo "running: PID $daemon_pid"
    exit 0
  fi
fi

mapfile -t matched_pids < <(find_daemon_pids)

valid_pids=()
for pid in "${matched_pids[@]}"; do
  if is_daemon_pid "$pid"; then
    valid_pids+=("$pid")
  fi
done

if [[ ${#valid_pids[@]} -eq 1 ]]; then
  printf '%s\n' "${valid_pids[0]}" > "$PID_FILE"
  echo "running: PID ${valid_pids[0]}"
  exit 0
fi

if [[ ${#valid_pids[@]} -gt 1 ]]; then
  echo "running: multiple PIDs ${valid_pids[*]}"
  exit 1
fi

echo "not running"
exit 1
