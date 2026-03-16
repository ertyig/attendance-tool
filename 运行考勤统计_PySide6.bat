@echo off
setlocal EnableExtensions
cd /d "%~dp0"
where py >nul 2>nul && set "PY_CMD=py"
if not defined PY_CMD where python >nul 2>nul && set "PY_CMD=python"
if not defined PY_CMD (
  echo Python was not found.
  pause
  exit /b 1
)
%PY_CMD% attendance_pyside6.py
pause
