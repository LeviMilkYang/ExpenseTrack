#!/usr/bin/env bash

set -euo pipefail

BASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_FILE="$BASE_DIR/telegram_expense_daemon.pid"
DAEMON_SCRIPT="$BASE_DIR/telegram_expense_daemon.py"

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
    pgrep -f "$DAEMON_SCRIPT" || true
    return
  fi

  ps -ef | grep "$DAEMON_SCRIPT" | grep -v grep | awk '{print $2}' || true
}

resolve_target_pids() {
  local pid
  local candidates=()

  if [[ -f "$PID_FILE" ]]; then
    pid="$(tr -d '[:space:]' < "$PID_FILE")"
    if is_daemon_pid "$pid"; then
      candidates+=("$pid")
    fi
  fi

  while IFS= read -r pid; do
    if is_daemon_pid "$pid"; then
      candidates+=("$pid")
    fi
  done < <(find_daemon_pids)

  if [[ ${#candidates[@]} -eq 0 ]]; then
    return 1
  fi

  printf '%s\n' "${candidates[@]}" | awk '!seen[$0]++'
}

if ! mapfile -t target_pids < <(resolve_target_pids); then
  rm -f "$PID_FILE"
  echo "telegram_expense_daemon.py is not running"
  exit 1
fi

for daemon_pid in "${target_pids[@]}"; do
  kill "$daemon_pid"
done

for _ in {1..10}; do
  all_stopped=true
  for daemon_pid in "${target_pids[@]}"; do
    if is_daemon_pid "$daemon_pid"; then
      all_stopped=false
      break
    fi
  done
  if [[ "$all_stopped" == true ]]; then
    rm -f "$PID_FILE"
    echo "stopped telegram_expense_daemon.py (PID(s) ${target_pids[*]})"
    exit 0
  fi
  sleep 1
done

echo "telegram_expense_daemon.py did not exit within 10 seconds: ${target_pids[*]}"
exit 1
