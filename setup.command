#!/bin/bash
# One-time setup on macOS. In Terminal: bash setup.command
cd "$(dirname "$0")" || exit 1
PY=$(command -v python3.12 || command -v python3)
if [ -z "$PY" ]; then echo "Python not found. Install Python 3.12 from python.org first."; exit 1; fi
echo "Using $("$PY" --version)"
set -e
"$PY" -m venv .venv
.venv/bin/python -m pip install --upgrade pip --quiet
.venv/bin/pip install -r requirements.txt
.venv/bin/playwright install chromium
echo ""
echo "Setup finished. Start the app with:  bash start.command"
