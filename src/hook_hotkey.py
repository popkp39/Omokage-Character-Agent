"""Hook イベント発生時に VMagicMirror へ表情ホットキーを送信する軽量スクリプト。

使い方:
    python hook_hotkey.py <イベント名>

例:
    python hook_hotkey.py Stop
    python hook_hotkey.py PostToolUseFailure

設定の hook_expression_mapping にイベント名→表情IDのマッピングがあれば
そのホットキーを送信する。マッピングがなければ何もしない。

■ デバウンス方式
短時間に連続発火するイベント（ToolUse成功→ToolUse開始→…→応答完了）を
そのまま送ると表情がパラパラ漫画になるため、デバウンスで「最後の1つ」だけ送信する。

仕組み:
  1. イベント発火 → 表情IDとタイムスタンプをファイルに書き込む
  2. cooldown_ms だけ sleep する
  3. sleep 後にファイルを再読み込みし、自分が書いた値がまだ最新なら送信する
     （後続イベントに上書きされていたら、自分は破棄される）
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import config
import send_to_avatar

# デバウンス用ペンディングファイル（src/ 直下に生成）
_PENDING_FILE = Path(__file__).with_name(".hook_pending")


def _write_pending(expression_id: int, timestamp: str) -> None:
    """送信候補をファイルに書き込む。後続イベントが来れば上書きされる。"""
    try:
        data = json.dumps({"expression_id": expression_id, "ts": timestamp})
        _PENDING_FILE.write_text(data, encoding="utf-8")
    except OSError:
        pass


def _read_pending() -> tuple[int, str] | None:
    """ペンディングファイルを読み取る。"""
    try:
        data = json.loads(_PENDING_FILE.read_text(encoding="utf-8"))
        return int(data["expression_id"]), str(data["ts"])
    except (OSError, json.JSONDecodeError, KeyError, TypeError, ValueError):
        return None


def main() -> int:
    if len(sys.argv) < 2:
        print("hook_hotkey: イベント名が指定されていません", file=sys.stderr)
        return 1

    event_name = sys.argv[1]

    try:
        settings = config.load_settings()
    except Exception as error:
        print(f"hook_hotkey: 設定の読み込みに失敗: {error}", file=sys.stderr)
        return 1

    if not bool(settings.get("avatar_enabled", True)):
        return 0

    if not bool(settings.get("hook_hotkey_enabled", False)):
        return 0

    mapping = settings.get("hook_expression_mapping")
    if not isinstance(mapping, dict):
        print("hook_hotkey: hook_expression_mapping が不正です", file=sys.stderr)
        return 1

    expression_id = mapping.get(event_name, 0)
    if not expression_id:
        return 0

    try:
        expression_id = int(expression_id)
    except (TypeError, ValueError):
        print(f"hook_hotkey: 表情ID が不正です: {expression_id!r}", file=sys.stderr)
        return 1

    if not config.EXPRESSION_ID_MIN <= expression_id <= config.EXPRESSION_ID_MAX:
        print(f"hook_hotkey: 表情ID が範囲外です: {expression_id}", file=sys.stderr)
        return 1

    # クールダウン（デバウンス間隔）
    try:
        cooldown_ms = int(settings.get("hook_cooldown_ms", config.HOOK_COOLDOWN_MS))
    except (TypeError, ValueError):
        cooldown_ms = config.HOOK_COOLDOWN_MS

    # デバウンス: 自分のイベントを書き込み → 待機 → 最新なら送信
    # float の等値比較を避けるため、文字列として比較する
    my_ts = repr(time.time())
    _write_pending(expression_id, my_ts)

    if cooldown_ms > 0:
        time.sleep(cooldown_ms / 1000.0)

    # sleep 後に再読み込み: 後続イベントに上書きされていたら自分は破棄
    current = _read_pending()
    if current is None:
        return 0
    current_eid, current_ts = current
    if current_ts != my_ts:
        return 0

    # 本命（send_to_avatar.py）が直近に送信していたら遠慮する
    # 音声再生中の Hook 上書きを防ぐため、クールダウンとは独立の固定値を使う
    _AVATAR_GUARD_MS = 3000
    avatar_sent_at = send_to_avatar.get_avatar_sent_time()
    if avatar_sent_at > 0:
        elapsed_ms = (time.time() - avatar_sent_at) * 1000
        if elapsed_ms < _AVATAR_GUARD_MS:
            return 0

    # 自分がまだ最新 → 送信
    try:
        send_to_avatar.send_hotkey(current_eid, settings)
    except Exception as error:
        print(f"hook_hotkey: {event_name} のホットキー送信に失敗: {error}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
