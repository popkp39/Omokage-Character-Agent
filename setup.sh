#!/bin/bash
# Omokage-Character-Agent (OCA) — セットアップスクリプト
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "=== Omokage-Character-Agent Setup ==="
echo ""

# Python バージョン確認
if ! command -v python &> /dev/null && ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python が見つかりません。Python 3.10 以上をインストールしてください。"
    exit 1
fi

PYTHON_CMD="python3"
if ! command -v python3 &> /dev/null; then
    PYTHON_CMD="python"
fi

# Python 3.10 以上であることを確認
PYTHON_VERSION=$($PYTHON_CMD --version 2>&1)
if ! $PYTHON_CMD -c "import sys; assert sys.version_info >= (3, 10)" 2>/dev/null; then
    echo "[ERROR] $PYTHON_VERSION が検出されましたが、Python 3.10 以上が必要です。"
    exit 1
fi
echo "[OK] $PYTHON_VERSION"

# 仮想環境の作成
if [ ! -d ".venv" ]; then
    echo ""
    if ! $PYTHON_CMD src/_create_venv.py; then
        echo "[ERROR] 仮想環境の作成に失敗しました"
        exit 1
    fi
else
    echo "[OK] .venv は既に存在します"
fi

# 仮想環境を有効化
if [ -f ".venv/Scripts/activate" ]; then
    source .venv/Scripts/activate
elif [ -f ".venv/bin/activate" ]; then
    source .venv/bin/activate
else
    echo "[ERROR] 仮想環境の activate が見つかりません"
    exit 1
fi

# 依存パッケージのインストール（失敗しても動作確認へ進む）
VERIFY_FAILED=0
echo ""
python src/_pip_install.py || VERIFY_FAILED=1

# 動作確認
echo ""
echo "--- 動作確認 ---"
if ! python -c "import sys; sys.path.insert(0,'src'); import config; print('[OK] config.py を読み込めました')"; then
    echo "[WARN] config.py の読み込みに失敗しました"
    VERIFY_FAILED=1
fi
if ! python -c "import sys; sys.path.insert(0,'src'); import send_to_avatar; print('[OK] send_to_avatar.py を読み込めました')"; then
    echo "[WARN] send_to_avatar.py の読み込みに失敗しました"
    VERIFY_FAILED=1
fi

echo ""
if [ $VERIFY_FAILED -eq 1 ]; then
    echo "=== セットアップは完了しましたが、動作確認で問題が見つかりました ==="
else
    echo "=== セットアップ完了 ==="
fi
echo ""
echo "次のステップ:"
echo "  1. VOICEVOX を起動してください"
echo "  2. 設定画面を開いてください:"
if [ -f ".venv/Scripts/activate" ]; then
    echo "     source .venv/Scripts/activate && python src/_launch_config.py"
else
    echo "     source .venv/bin/activate && python src/_launch_config.py"
fi
echo "  3. 画面の案内に沿って初期設定を進めてください"
echo ""
echo "  詳しくは はじめに_初回セットアップ手順.md をご覧ください"
