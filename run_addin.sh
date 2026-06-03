#!/usr/bin/env bash
# Launch the SheetMind bridge for the Office.js task-pane add-in.
#   ./run_addin.sh          # HTTPS on https://localhost:8765 (for Excel)
#   ./run_addin.sh --http   # plain HTTP (for a quick browser preview)
cd "$(dirname "$0")" || exit 1
PY=".venv/bin/python"
[ -x "$PY" ] || PY="python3"
echo "Starting SheetMind bridge…"
echo "Preview in a browser:  http://localhost:8765/taskpane.html?demo=1   (use --http)"
exec "$PY" server.py "$@"
