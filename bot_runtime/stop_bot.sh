#!/usr/bin/env bash

set -euo pipefail

BASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_FILE="$BASE_DIR/telegram_expense_daemon.pid"

if [[ ! -f "$PID_FILE" ]]; then
  echo "pid file not found: $PID_FILE"
  exit 1
fi

daemon_pid="$(tr -d '[:space:]' < "$PID_FILE")"

if ! kill -0 "$daemon_pid" 2>/dev/null; then
  rm -f "$PID_FILE"
  echo "process $daemon_pid is not running"
  exit 1
fi

kill "$daemon_pid"

for _ in {1..10}; do
  if ! kill -0 "$daemon_pid" 2>/dev/null; then
    rm -f "$PID_FILE"
    echo "stopped telegram_expense_daemon.py (PID $daemon_pid)"
    exit 0
  fi
  sleep 1
done

echo "process $daemon_pid did not exit within 10 seconds"
exit 1
