@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>nul

cd /d "%~dp0"

set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py"
if not defined PY_CMD where python >nul 2>nul && set "PY_CMD=python"

if not defined PY_CMD (
    echo [ERROR] Python was not found.
    echo Please install Python 3 first.
    pause
    exit /b 1
)

%PY_CMD% attendance_gui.py
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo [ERROR] Program exited with code %EXIT_CODE%.
    echo If this is the first run, try:
    echo %PY_CMD% -m pip install ttkbootstrap pandas openpyxl xlrd holidays chinese-calendar
    pause
    exit /b %EXIT_CODE%
)
