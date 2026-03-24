@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

REM .venv があればそれを使う（既存フラグを記録）
set VENV_EXISTED=0
if exist ".venv\Scripts\python.exe" (
    set VENV_EXISTED=1
    set PYTHON=.venv\Scripts\python.exe
    goto :run_migration
)

REM .venv がない場合はシステム Python で実行（.venv は移行でコピーされる）
where python >nul 2>&1
if errorlevel 1 goto :no_python

python -c "import sys; exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if errorlevel 1 goto :old_python

set PYTHON=python

:run_migration
%PYTHON% -c "print('============================================'); print('  Omokage-Character-Agent'); print('  バージョンアップ・データ移行ツール'); print('============================================'); print(); print('  旧バージョンからデータ（設定・プリセット・キャラ設定・ログ・.venv）を'); print('  このフォルダへ移行します。'); print(); print('  フォルダ選択ダイアログが開きます。'); print('  旧バージョンの OmokageCharacterAgent フォルダを選択してください。'); print()"

%PYTHON% src\_migrate_data.py
if errorlevel 1 goto :migration_failed

REM .venv がコピーされた場合のみ、差分 pip install を実行
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

%PYTHON% -c "print(); print('============================================'); print(); print('  バージョンアップ・データ移行が完了しました!'); print(); print('  次のステップ:'); print('    - 「設定画面を開く.bat」でパスの確認'); print('    - 動作テスト'); print(); print('============================================'); print()"

echo 何かキーを押すと閉じます。
pause >nul
exit /b 0

:migration_failed
%PYTHON% -c "print(); print('============================================'); print(); print('  移行は完了していません。'); print('  上のメッセージを確認してください。'); print(); print('============================================'); print()"
echo 何かキーを押すと閉じます。
pause >nul
exit /b 1

:no_python
echo.
echo  [エラー] Python が見つかりません。
echo  Python 3.10 以上をインストールするか、初回セットアップを先に実行してください。
echo.
echo 何かキーを押すと閉じます。
pause >nul
exit /b 1

:old_python
echo.
echo  [エラー] Python 3.10 以上が必要です。
echo.
echo 何かキーを押すと閉じます。
pause >nul
exit /b 1
