#!/bin/bash
# Start the app on macOS. In Terminal: bash start.command   (Ctrl+C to stop)
cd "$(dirname "$0")" || exit 1
if [ ! -x .venv/bin/python ]; then echo "Run setup first:  bash setup.command"; exit 1; fi
.venv/bin/python app.py
