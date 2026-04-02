#!/usr/bin/env bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
. "$SCRIPT_DIR/bot_common.sh"
STATUS_SCRIPT="$BASE_DIR/status_bot.sh"
STOP_SCRIPT="$BASE_DIR/stop_bot.sh"
START_SCRIPT="$BASE_DIR/start_bot.sh"
PROJECT_DIR="$(resolve_project_dir)"

echo "validating new daemon startup..."
if ! timeout 40s python3 "$BASE_DIR/telegram_expense_daemon.py" --once --verbose --startup-timeout-seconds 30 --state-file "$STATE_FILE" --excel-path "$PROJECT_DIR/expense.xlsx"; then
  echo "validation failed or timed out"
  exit 1
fi
echo "validation passed"

if "$STATUS_SCRIPT" >/dev/null 2>&1; then
  echo "stopping current daemon..."
  "$STOP_SCRIPT"
fi

echo "starting new daemon..."
"$START_SCRIPT"
