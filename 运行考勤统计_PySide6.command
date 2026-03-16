#!/bin/bash
set -e
cd "$(dirname "$0")"
if [ -x ".venv-macgui/bin/python" ]; then
  PYTHON_CMD=".venv-macgui/bin/python"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON_CMD="python3"
else
  echo "python3 not found"
  read -r -p "Press Enter to close..." _
  exit 1
fi
env -u SYSTEM_VERSION_COMPAT "$PYTHON_CMD" attendance_pyside6.py
read -r -p "Press Enter to close..." _
