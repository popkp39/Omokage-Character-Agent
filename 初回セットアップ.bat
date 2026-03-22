@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

echo === Omokage-Character-Agent Setup ===
echo.

REM Python 存在確認
where python >nul 2>&1
if errorlevel 1 goto :no_python

for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo [OK] %%v

REM Python バージョンチェック（3.10 以上）
python -c "import sys; exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python 3.10 以上が必要です。現在のバージョンが古すぎます。
    pause
    exit /b 1
)

REM 仮想環境の作成
if exist ".venv" (
    echo [OK] .venv は既に存在します
) else (
    echo.
    echo --- 仮想環境を作成しています ---
    python -m venv .venv
    echo [OK] .venv を作成しました
)

REM 仮想環境を有効化
if not exist ".venv\Scripts\activate.bat" goto :no_venv
call .venv\Scripts\activate.bat

REM 依存パッケージのインストール
echo.
echo --- 依存パッケージをインストールしています ---
pip install -r requirements.txt --quiet
echo [OK] 依存パッケージをインストールしました

REM 動作確認
echo.
echo --- 動作確認 ---
python -c "import sys; sys.path.insert(0,'src'); import config; print('[OK] config.py を読み込めました')"
python -c "import sys; sys.path.insert(0,'src'); import send_to_avatar; print('[OK] send_to_avatar.py を読み込めました')"

echo.
echo === セットアップ完了 ===
echo.
echo 次のステップ:
echo   1. VOICEVOX を起動してください
echo   2. 設定画面を開く.bat をダブルクリックしてください
echo   3. 画面の案内に沿って初期設定を進めてください
echo   詳しくは はじめに_初回セットアップ手順.md をご覧ください
echo.
pause
exit /b 0

:no_python
echo [ERROR] Python が見つかりません。Python 3.10 以上をインストールしてください。
pause
exit /b 1

:no_venv
echo [ERROR] 仮想環境の activate が見つかりません
pause
exit /b 1
