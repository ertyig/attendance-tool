@echo off
if /I not "%~1"=="__run__" (
    start "" "%ComSpec%" /k call "%~f0" __run__
    exit /b
)

setlocal EnableExtensions
cd /d "%~dp0"

set "LOG_FILE=%~dp0build_windows_pyside6.log"
set "PY_CMD="

> "%LOG_FILE%" echo [INFO] Start build at %date% %time%
>> "%LOG_FILE%" echo [INFO] Workdir: %cd%

where py >nul 2>nul && set "PY_CMD=py"
if not defined PY_CMD where python >nul 2>nul && set "PY_CMD=python"

if not defined PY_CMD (
    echo [ERROR] Python was not found.
    >> "%LOG_FILE%" echo [ERROR] Python was not found.
    goto :error
)

%PY_CMD% --version >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :error

%PY_CMD% -m pip install --upgrade pip pyinstaller PySide6 pandas openpyxl xlrd holidays chinese-calendar >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :error

%PY_CMD% -m PyInstaller --noconfirm --clean attendance_pyside6.spec >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :error

echo [OK] Build completed.
>> "%LOG_FILE%" echo [OK] Build completed.
echo Output folder:
echo %~dp0dist
goto :end

:error
echo [ERROR] Build failed.
echo Please open this log file:
echo %LOG_FILE%
>> "%LOG_FILE%" echo [ERROR] Build failed.

:end
pause
exit /b 0
