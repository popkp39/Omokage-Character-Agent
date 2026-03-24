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
if errorlevel 1 goto :old_python

REM 仮想環境の作成
if exist ".venv" goto :venv_exists
echo.
python src\_create_venv.py
goto :activate_venv

:venv_exists
echo [OK] .venv は既に存在します

:activate_venv
REM 仮想環境を有効化
if not exist ".venv\Scripts\activate.bat" goto :no_venv
call .venv\Scripts\activate.bat

REM 依存パッケージのインストール
echo.
echo --- 依存パッケージをインストールしています ---
python src\_pip_install.py
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
echo   2. VMagicMirror を起動してください
echo   3. 「設定画面を開く.bat」をダブルクリックしてください
echo   4. 画面の案内に沿って初期設定を進めてください
echo   詳しくは「はじめに_初回セットアップ手順.md」をご覧ください
echo.
echo 何かキーを押すと閉じます。
pause >nul
exit /b 0

:no_python
echo [ERROR] Python が見つかりません。Python 3.10 以上をインストールしてください。
echo.
echo 何かキーを押すと閉じます。
pause >nul
exit /b 1

:old_python
echo [ERROR] Python 3.10 以上が必要です。現在のバージョンが古すぎます。
echo.
echo 何かキーを押すと閉じます。
pause >nul
exit /b 1

:no_venv
echo [ERROR] 仮想環境の activate が見つかりません。
echo.
echo 何かキーを押すと閉じます。
pause >nul
exit /b 1
