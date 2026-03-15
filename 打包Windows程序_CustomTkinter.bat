@echo off
if /I not "%~1"=="__run__" (
    start "" "%ComSpec%" /k call "%~f0" __run__
    exit /b
)

setlocal EnableExtensions
cd /d "%~dp0"

set "LOG_FILE=%~dp0build_windows_customtkinter.log"
set "PY_CMD="

> "%LOG_FILE%" echo [INFO] Start build at %date% %time%
>> "%LOG_FILE%" echo [INFO] Workdir: %cd%

title Attendance Assistant CustomTkinter Build
echo ==============================
echo Attendance Assistant CustomTkinter Build
echo ==============================
echo.
echo Log file:
echo %LOG_FILE%
echo.

where py >nul 2>nul && set "PY_CMD=py"
if not defined PY_CMD where python >nul 2>nul && set "PY_CMD=python"

if not defined PY_CMD (
    echo [ERROR] Python was not found.
    echo [ERROR] Please install Python 3 first.
    >> "%LOG_FILE%" echo [ERROR] Python was not found.
    goto :error
)

echo [INFO] Python launcher: %PY_CMD%
>> "%LOG_FILE%" echo [INFO] Python launcher: %PY_CMD%
echo.

echo [1/4] Checking Python...
%PY_CMD% --version >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] Python command failed.
    >> "%LOG_FILE%" echo [ERROR] Python command failed.
    goto :error
)

echo [2/4] Installing or updating packages...
%PY_CMD% -m pip install --upgrade pip pyinstaller customtkinter pandas openpyxl xlrd holidays chinese-calendar >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] pip install failed.
    >> "%LOG_FILE%" echo [ERROR] pip install failed.
    goto :error
)

echo [3/4] Building exe...
%PY_CMD% -m PyInstaller --noconfirm --clean attendance_customtkinter.spec >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    >> "%LOG_FILE%" echo [ERROR] PyInstaller build failed.
    goto :error
)

echo [4/4] Done.
echo Output folder:
echo %~dp0dist
>> "%LOG_FILE%" echo [OK] Build completed.
echo.
echo You can close this window now.
goto :end

:error
echo.
echo [ERROR] Build failed.
echo Please open this log file:
echo %LOG_FILE%
echo.
echo Recommended next step:
echo Run the environment check script.
>> "%LOG_FILE%" echo [ERROR] Build ended with failure.

:end
echo.
pause
exit /b 0
