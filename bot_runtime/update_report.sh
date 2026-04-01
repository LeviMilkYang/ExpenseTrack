#!/usr/bin/env bash

set -euo pipefail

BASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$BASE_DIR/.." && pwd)"
STATE_FILE="$BASE_DIR/telegram_bot_config.json"

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
CUTOFF_DATE="$(date -d "$(date +%Y-%m-01) -1 day" +%F)"

python3 "$BASE_DIR/generate_expense_report.py" \
  --source "$PROJECT_DIR/expense.xlsx" \
  --report "$PROJECT_DIR/expense_report.xlsx" \
  --cutoff-date "$CUTOFF_DATE"
