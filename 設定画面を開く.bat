@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" goto :no_setup

start "" .venv\Scripts\pythonw.exe src\_launch_config.py
exit /b 0

:no_setup
python -c "print(); print('  セットアップが完了していません。'); print(); print('  新規の場合は「初回セットアップ.bat」を、'); print('  旧バージョンからの移行の場合は「バージョンアップ・データ移行.bat」を'); print('  先に実行してから、もう一度お試しください。'); print(); print('  何かキーを押すと閉じます。')"
pause >nul
exit /b 1
