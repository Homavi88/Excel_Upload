#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

echo "============================================"
echo "  Excel Loader for TARGET.xlsx"
echo "============================================"

if ! command -v python3 >/dev/null 2>&1; then
  echo "[ERROR] python3 not found"
  exit 1
fi

PYTHON=python3
if [[ ! -d .venv ]]; then
  echo "Creating virtual environment..."
  if python3 -m venv .venv; then
    # shellcheck disable=SC1091
    source .venv/bin/activate
  else
    echo "[WARN] venv unavailable, using system Python"
    rm -rf .venv
    python3 -m pip install --user -r requirements.txt
    export PATH="$HOME/.local/bin:$PATH"
  fi
else
  # shellcheck disable=SC1091
  source .venv/bin/activate
fi

if [[ -x .venv/bin/python ]]; then
  PYTHON=.venv/bin/python
  "$PYTHON" -m pip install --upgrade pip >/dev/null
  "$PYTHON" -m pip install -r requirements.txt
fi

if [[ ! -f TARGET.xlsx ]]; then
  echo "[ERROR] TARGET.xlsx not found in $(pwd)"
  exit 1
fi

echo
echo "Starting server at http://127.0.0.1:8765"
echo "Press Ctrl+C to stop."
echo
"$PYTHON" -m uvicorn app.main:app --host 127.0.0.1 --port 8765
