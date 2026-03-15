#!/bin/zsh
set -e

cd "$(dirname "$0")"
unset SYSTEM_VERSION_COMPAT
PYTHON_BIN=".venv-macgui/bin/python"

if [ ! -x "$PYTHON_BIN" ]; then
  PYTHON_BIN="python3"
fi

env -u SYSTEM_VERSION_COMPAT "$PYTHON_BIN" attendance_nicegui.py

echo
echo "已关闭 NiceGUI 界面。"
echo "按回车键关闭窗口..."
read
