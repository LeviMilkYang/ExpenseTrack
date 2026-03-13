#!/usr/bin/env bash

set -euo pipefail

BASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$BASE_DIR/.." && pwd)"
STATE_FILE="$BASE_DIR/telegram_bot_state.json"

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

echo "validating new daemon startup..."
python3 "$BASE_DIR/telegram_expense_daemon.py" --once --verbose --state-file "$STATE_FILE" --excel-path "$PROJECT_DIR/expense.xlsx"
echo "validation passed"

if "$BASE_DIR/status_bot.sh" >/dev/null 2>&1; then
  echo "stopping current daemon..."
  "$BASE_DIR/stop_bot.sh"
fi

echo "starting new daemon..."
"$BASE_DIR/start_bot.sh"
