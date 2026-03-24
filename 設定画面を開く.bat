@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" goto :no_setup

start "" .venv\Scripts\pythonw.exe src\_launch_config.py
exit /b 0

:no_setup
echo.
echo  セットアップが完了していません。
echo.
echo  先に「初回セットアップ.bat」をダブルクリックして
echo  セットアップを完了してから、もう一度実行してください。
echo.
echo  何かキーを押すと閉じます。
pause >nul
exit /b 1
