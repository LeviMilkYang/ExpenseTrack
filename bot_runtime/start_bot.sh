#!/usr/bin/env bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$SCRIPT_DIR/bot_common.sh"
STATUS_SCRIPT="$BASE_DIR/status_bot.sh"
PROJECT_DIR="$(resolve_project_dir)"
CURRENT_LOG="$(current_log_path)"

if status_output="$("$STATUS_SCRIPT" 2>&1)"; then
  printf '%s\n' "$status_output"
  exit 0
fi
rm -f "$PID_FILE"

# 启动机器人，机器人内部也会写入该日志文件。
nohup python3 -u "$BASE_DIR/telegram_expense_daemon.py" --verbose --state-file "$STATE_FILE" --excel-path "$PROJECT_DIR/expense.xlsx" >>"$CURRENT_LOG" 2>&1 </dev/null &
daemon_pid=$!
echo "$daemon_pid" > "$PID_FILE"

sleep 2
if "$STATUS_SCRIPT" >/dev/null 2>&1; then
  echo "started telegram_expense_daemon.py with PID $daemon_pid"
  echo "log: $CURRENT_LOG"
  exit 0
fi

rm -f "$PID_FILE"
echo "failed to start telegram_expense_daemon.py"
exit 1
