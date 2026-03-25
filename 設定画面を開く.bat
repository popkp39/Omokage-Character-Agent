@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" goto :no_setup

start "" .venv\Scripts\pythonw.exe src\_launch_config.py
exit /b 0

:no_setup
echo.
echo   Setup is not complete.
echo.
echo   Please run "Setup.bat" first,
echo   or "Migration.bat" if upgrading from an older version.
echo.
echo   Press any key to close.
pause >nul
exit /b 1
