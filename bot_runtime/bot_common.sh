#!/usr/bin/env bash

BASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$BASE_DIR/.." && pwd)"
STATE_FILE="$BASE_DIR/telegram_bot_config.json"
PID_FILE="$BASE_DIR/telegram_expense_daemon.pid"
DAEMON_SCRIPT="$BASE_DIR/telegram_expense_daemon.py"

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

current_log_path() {
  local log_dir="$BASE_DIR/logs"
  mkdir -p "$log_dir"
  printf '%s/%s.log\n' "$log_dir" "$(date +"%Y-%m")"
}

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

collect_daemon_pids() {
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
