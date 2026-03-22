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
    echo "--- 仮想環境を作成しています ---"
    $PYTHON_CMD -m venv .venv
    echo "[OK] .venv を作成しました"
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

# 依存パッケージのインストール
echo ""
echo "--- 依存パッケージをインストールしています ---"
pip install -r requirements.txt --quiet --disable-pip-version-check
echo "[OK] 依存パッケージをインストールしました"

# 動作確認
echo ""
echo "--- 動作確認 ---"
python -c "import sys; sys.path.insert(0,'src'); import config; print('[OK] config.py を読み込めました')"
python -c "import sys; sys.path.insert(0,'src'); import send_to_avatar; print('[OK] send_to_avatar.py を読み込めました')"

echo ""
echo "=== セットアップ完了 ==="
echo ""
echo "次のステップ:"
echo "  1. VOICEVOX を起動してください"
echo "  2. 設定画面を開いてください:"
if [ -f ".venv/Scripts/activate" ]; then
    echo "     source .venv/Scripts/activate && python src/config.py"
else
    echo "     source .venv/bin/activate && python src/config.py"
fi
echo "  3. 画面の案内に沿って初期設定を進めてください"
echo ""
echo "  詳しくは はじめに_初回セットアップ手順.md をご覧ください"
