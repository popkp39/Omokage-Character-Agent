@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

echo === Omokage-Character-Agent Setup ===
echo.

REM Python check
where python >nul 2>&1
if errorlevel 1 goto :no_python

for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo [OK] %%v

REM Python version check (3.10+)
python -c "import sys; exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if errorlevel 1 goto :old_python

REM Create virtual environment
if exist ".venv" goto :venv_exists
echo.
python src\_create_venv.py
if errorlevel 1 goto :venv_create_failed
goto :activate_venv

:venv_exists
echo [OK] .venv already exists

:activate_venv
REM Activate virtual environment
if not exist ".venv\Scripts\activate.bat" goto :no_venv
call .venv\Scripts\activate.bat

REM Install dependencies
set VERIFY_FAILED=0
echo.
python src\_pip_install.py
if errorlevel 1 set VERIFY_FAILED=1

REM Verify
echo.
echo --- Verify ---
python -c "import sys; sys.path.insert(0,'src'); import config; print('[OK] config.py')"
if errorlevel 1 (
    echo [WARN] config.py failed to load
    set VERIFY_FAILED=1
)
python -c "import sys; sys.path.insert(0,'src'); import send_to_avatar; print('[OK] send_to_avatar.py')"
if errorlevel 1 (
    echo [WARN] send_to_avatar.py failed to load
    set VERIFY_FAILED=1
)

echo.
if %VERIFY_FAILED%==1 (
    echo === Setup complete, but some checks failed ===
) else (
    echo === Setup complete ===
)
echo.
echo Next steps:
echo   1. Start VOICEVOX
echo   2. Start VMagicMirror
echo   3. Double-click "Open Settings.bat"
echo   4. Follow the on-screen instructions
echo   See "Getting Started Guide" for details
echo.
echo Press any key to close.
pause >nul
exit /b 0

:no_python
echo [ERROR] Python not found. Please install Python 3.10 or later.
echo.
echo Press any key to close.
pause >nul
exit /b 1

:old_python
echo [ERROR] Python 3.10 or later is required. Your version is too old.
echo.
echo Press any key to close.
pause >nul
exit /b 1

:no_venv
echo [ERROR] .venv activate not found.
echo.
echo Press any key to close.
pause >nul
exit /b 1

:venv_create_failed
echo [ERROR] Failed to create virtual environment.
echo.
echo Press any key to close.
pause >nul
exit /b 1
