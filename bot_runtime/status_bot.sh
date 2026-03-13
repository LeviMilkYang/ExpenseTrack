#!/usr/bin/env bash

set -euo pipefail

BASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_FILE="$BASE_DIR/telegram_expense_daemon.pid"

if [[ -f "$PID_FILE" ]]; then
  daemon_pid="$(tr -d '[:space:]' < "$PID_FILE")"
  if kill -0 "$daemon_pid" 2>/dev/null; then
    echo "running: PID $daemon_pid"
    exit 0
  fi
  echo "stale pid file: $PID_FILE"
  exit 1
fi

echo "not running"
exit 1
