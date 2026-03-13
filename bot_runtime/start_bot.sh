#!/usr/bin/env bash

set -euo pipefail

BASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$BASE_DIR/.." && pwd)"
STATE_FILE="$BASE_DIR/telegram_bot_config.json"
PID_FILE="$BASE_DIR/telegram_expense_daemon.pid"

# 日志目录
LOG_DIR="$BASE_DIR/logs"
mkdir -p "$LOG_DIR"
MONTH_STR=$(date +"%Y-%m")
# 重定向日志依然指向当前月日志，以便 nohup 异常捕获
CURRENT_LOG="$LOG_DIR/${MONTH_STR}.log"

resolve_project_dir() {
  if [[ -f "$STATE_FILE" ]]; then
    local configured
    configured="$(python3 -c 'import json, sys; from pathlib import Path; data = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8")); print(str(data.get("project_dir", "")).strip())' "$STATE_FILE" 2>/dev/null || true)"
    if [[ -n "$configured" ]]; then
      echo "$configured"
      return
    fi
  fi
  echo "$PROJECT_DIR"
}

PROJECT_DIR="$(resolve_project_dir)"

is_running() {
  local pid="$1"
  [[ -n "$pid" ]] && kill -0 "$pid" 2>/dev/null
}

if [[ -f "$PID_FILE" ]]; then
  existing_pid="$(tr -d '[:space:]' < "$PID_FILE")"
  if is_running "$existing_pid"; then
    echo "telegram_expense_daemon.py is already running with PID $existing_pid"
    exit 0
  fi
  rm -f "$PID_FILE"
fi

# 启动机器人，机器人内部也会写入该日志文件
nohup python3 -u "$BASE_DIR/telegram_expense_daemon.py" --verbose --state-file "$STATE_FILE" --excel-path "$PROJECT_DIR/expense.xlsx" >>"$CURRENT_LOG" 2>&1 </dev/null &
daemon_pid=$!
echo "$daemon_pid" > "$PID_FILE"

sleep 1
if is_running "$daemon_pid"; then
  echo "started telegram_expense_daemon.py with PID $daemon_pid"
  echo "log: $CURRENT_LOG"
  exit 0
fi

rm -f "$PID_FILE"
echo "failed to start telegram_expense_daemon.py"
exit 1
