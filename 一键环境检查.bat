@echo off
if /I not "%~1"=="__run__" (
    start "" "%ComSpec%" /k call "%~f0" __run__
    exit /b
)

setlocal EnableExtensions
cd /d "%~dp0"

set "LOG_FILE=%~dp0env_check.log"
set "PY_CMD="

> "%LOG_FILE%" echo [INFO] Start env check at %date% %time%
>> "%LOG_FILE%" echo [INFO] Workdir: %cd%

title Attendance Tool Env Check
echo ==============================
echo Attendance Tool Env Check
echo ==============================
echo.
echo Log file:
echo %LOG_FILE%
echo.

where py >nul 2>nul && set "PY_CMD=py"
if not defined PY_CMD where python >nul 2>nul && set "PY_CMD=python"

if not defined PY_CMD (
    echo [FAIL] Python was not found.
    echo Please install Python 3 first.
    >> "%LOG_FILE%" echo [FAIL] Python was not found.
    goto :end_fail
)

echo [OK] Python launcher found: %PY_CMD%
>> "%LOG_FILE%" echo [OK] Python launcher found: %PY_CMD%
echo.

echo [CHECK] Python version
%PY_CMD% --version >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [FAIL] Cannot run Python.
    >> "%LOG_FILE%" echo [FAIL] Cannot run Python.
    goto :end_fail
)
echo [OK] Python command works.
echo.

echo [CHECK] pip version
%PY_CMD% -m pip --version >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [FAIL] pip is not available.
    >> "%LOG_FILE%" echo [FAIL] pip is not available.
    goto :end_fail
)
echo [OK] pip is available.
echo.

echo [CHECK] PyInstaller
%PY_CMD% -m PyInstaller --version >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [WARN] PyInstaller is not installed.
    >> "%LOG_FILE%" echo [WARN] PyInstaller is not installed.
) else (
    echo [OK] PyInstaller is available.
    >> "%LOG_FILE%" echo [OK] PyInstaller is available.
)
echo.

echo [CHECK] Required Python packages
%PY_CMD% -c "import ttkbootstrap, pandas, openpyxl, xlrd, holidays, chinese_calendar; print('[OK] All required packages are installed.')" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [WARN] Some required packages are missing.
    echo Install with:
    echo %PY_CMD% -m pip install ttkbootstrap pandas openpyxl xlrd holidays chinese-calendar pyinstaller
    >> "%LOG_FILE%" echo [WARN] Some required packages are missing.
) else (
    echo [OK] Runtime dependencies are ready.
    >> "%LOG_FILE%" echo [OK] Runtime dependencies are ready.
)
echo.

echo [CHECK] CustomTkinter package
%PY_CMD% -c "import customtkinter; print('[OK] CustomTkinter is installed.')" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [WARN] CustomTkinter is not installed.
    echo Install with:
    echo %PY_CMD% -m pip install customtkinter
    >> "%LOG_FILE%" echo [WARN] CustomTkinter is not installed.
) else (
    echo [OK] CustomTkinter is available.
    >> "%LOG_FILE%" echo [OK] CustomTkinter is available.
)
echo.

echo [CHECK] NiceGUI package
%PY_CMD% -c "import nicegui; print('[OK] NiceGUI is installed.')" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo [WARN] NiceGUI is not installed.
    echo Install with:
    echo %PY_CMD% -m pip install nicegui
    >> "%LOG_FILE%" echo [WARN] NiceGUI is not installed.
) else (
    echo [OK] NiceGUI is available.
    >> "%LOG_FILE%" echo [OK] NiceGUI is available.
)
echo.

echo [CHECK] Key project files
if exist "attendance_gui.py" (
    echo [OK] attendance_gui.py
    >> "%LOG_FILE%" echo [OK] attendance_gui.py
) else (
    echo [FAIL] attendance_gui.py not found
    >> "%LOG_FILE%" echo [FAIL] attendance_gui.py not found
)

if exist "attendance_report.py" (
    echo [OK] attendance_report.py
    >> "%LOG_FILE%" echo [OK] attendance_report.py
) else (
    echo [FAIL] attendance_report.py not found
    >> "%LOG_FILE%" echo [FAIL] attendance_report.py not found
)

if exist "attendance_gui.spec" (
    echo [OK] attendance_gui.spec
    >> "%LOG_FILE%" echo [OK] attendance_gui.spec
) else (
    echo [WARN] attendance_gui.spec not found
    >> "%LOG_FILE%" echo [WARN] attendance_gui.spec not found
)

if exist "attendance_nicegui.py" (
    echo [OK] attendance_nicegui.py
    >> "%LOG_FILE%" echo [OK] attendance_nicegui.py
) else (
    echo [WARN] attendance_nicegui.py not found
    >> "%LOG_FILE%" echo [WARN] attendance_nicegui.py not found
)

if exist "attendance_nicegui.spec" (
    echo [OK] attendance_nicegui.spec
    >> "%LOG_FILE%" echo [OK] attendance_nicegui.spec
) else (
    echo [WARN] attendance_nicegui.spec not found
    >> "%LOG_FILE%" echo [WARN] attendance_nicegui.spec not found
)

if exist "attendance_customtkinter.py" (
    echo [OK] attendance_customtkinter.py
    >> "%LOG_FILE%" echo [OK] attendance_customtkinter.py
) else (
    echo [WARN] attendance_customtkinter.py not found
    >> "%LOG_FILE%" echo [WARN] attendance_customtkinter.py not found
)

if exist "attendance_customtkinter.spec" (
    echo [OK] attendance_customtkinter.spec
    >> "%LOG_FILE%" echo [OK] attendance_customtkinter.spec
) else (
    echo [WARN] attendance_customtkinter.spec not found
    >> "%LOG_FILE%" echo [WARN] attendance_customtkinter.spec not found
)
echo.

echo [CHECK] Data folder
if exist "data" (
    echo [OK] data
    >> "%LOG_FILE%" echo [OK] data
) else (
    echo [WARN] data folder not found
    >> "%LOG_FILE%" echo [WARN] data folder not found
)
echo.

echo Env check finished.
>> "%LOG_FILE%" echo [INFO] Env check finished.
goto :end_ok

:end_fail
echo.
echo Env check finished with errors.
echo Please check this log file:
echo %LOG_FILE%
>> "%LOG_FILE%" echo [INFO] Env check finished with errors.
pause
exit /b 1

:end_ok
echo.
echo Check completed. Details were written to:
echo %LOG_FILE%
pause
exit /b 0
