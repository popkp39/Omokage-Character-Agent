@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
    echo セットアップが完了していません。
    echo 先に 初回セットアップ.bat を実行してください。
    pause
    exit /b 1
)

start "" .venv\Scripts\pythonw.exe src\_launch_config.py
