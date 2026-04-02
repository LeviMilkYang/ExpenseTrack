#!/usr/bin/env bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$SCRIPT_DIR/bot_common.sh"

if ! mapfile -t target_pids < <(collect_daemon_pids) || [[ ${#target_pids[@]} -eq 0 ]]; then
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
