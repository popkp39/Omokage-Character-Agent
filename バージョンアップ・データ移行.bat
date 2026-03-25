@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

REM Use .venv if available
set VENV_EXISTED=0
if exist ".venv\Scripts\python.exe" (
    set VENV_EXISTED=1
    set PYTHON=.venv\Scripts\python.exe
    goto :run_migration
)

REM Fall back to system Python
where python >nul 2>&1
if errorlevel 1 goto :no_python

python -c "import sys; exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if errorlevel 1 goto :old_python

set PYTHON=python

:run_migration
echo ============================================
echo   Omokage-Character-Agent
echo   Migration Tool
echo ============================================
echo.
echo   Migrate data (settings, presets, characters, logs, .venv)
echo   from an older version to this folder.
echo.
echo   A folder selection dialog will open.
echo   Select the old OmokageCharacterAgent folder.
echo.

%PYTHON% src\_migrate_data.py
if errorlevel 1 goto :migration_failed

REM Run pip install if .venv was copied during migration
if "%VENV_EXISTED%"=="0" if exist ".venv\Scripts\python.exe" (
    echo.
    echo --- pip install ---
    call .venv\Scripts\activate.bat
    python src\_pip_install.py
    echo.
    echo --- check ---
    python -c "import sys; sys.path.insert(0,'src'); import config; print('[OK] config.py')"
    python -c "import sys; sys.path.insert(0,'src'); import send_to_avatar; print('[OK] send_to_avatar.py')"
)

echo.
echo ============================================
echo.
echo   Migration complete!
echo.
echo   Next steps:
echo     - Open "Settings" to verify paths
echo     - Test that everything works
echo.
echo ============================================
echo.
echo Press any key to close.
pause >nul
exit /b 0

:migration_failed
echo.
echo ============================================
echo.
echo   Migration did not complete.
echo   Check the messages above.
echo.
echo ============================================
echo.
echo Press any key to close.
pause >nul
exit /b 1

:no_python
echo.
echo  [ERROR] Python not found.
echo  Please install Python 3.10+ or run Setup first.
echo.
echo Press any key to close.
pause >nul
exit /b 1

:old_python
echo.
echo  [ERROR] Python 3.10 or later is required.
echo.
echo Press any key to close.
pause >nul
exit /b 1
