#!/bin/zsh
set +e

cd "$(dirname "$0")"

LOG_FILE="attendance_gui_startup.log"
unset SYSTEM_VERSION_COMPAT
PYTHON_BIN=".venv-macgui/bin/python"

if [ ! -x "$PYTHON_BIN" ]; then
  PYTHON_BIN="python3"
fi

env -u SYSTEM_VERSION_COMPAT "$PYTHON_BIN" attendance_gui.py 2>&1 | tee "$LOG_FILE"
EXIT_CODE=${pipestatus[1]}

echo
if [ "$EXIT_CODE" -ne 0 ]; then
  echo "界面没有正常启动。"
  echo "请查看诊断日志：$LOG_FILE"
else
  echo "已关闭界面。"
fi
echo "按回车键关闭窗口..."
read
