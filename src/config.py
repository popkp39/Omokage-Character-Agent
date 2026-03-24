from __future__ import annotations

import ctypes
import json
import os
import platform
import re
import sys
import threading
import tkinter as tk
from pathlib import Path

if platform.system() == "Windows":
    from ctypes import wintypes
from tkinter import filedialog, messagebox, ttk

# numpy, requests, sounddevice はバックグラウンドで並列インポート（起動高速化）
np = None
sd = None
requests = None
_imports_ready = threading.Event()
_cached_device_list: list[object] | None = None
_device_list_lock = threading.Lock()
_devices_ready = threading.Event()


_import_error: str | None = None


def _background_import() -> None:
    """バックグラウンドスレッドで重いモジュールをインポートし、デバイス一覧もプリフェッチする。"""
    global np, sd, requests, _cached_device_list, _import_error
    try:
        import numpy
        import sounddevice
        import requests as _requests

        np = numpy
        sd = sounddevice
        requests = _requests
    except ImportError as error:
        _import_error = str(error)
        print(f"依存パッケージのインポートに失敗しました: {error}", file=sys.stderr)
    _imports_ready.set()
    if _import_error is not None:
        _devices_ready.set()
        return
    try:
        devices = list(sounddevice.query_devices())
    except Exception:
        devices = []
    with _device_list_lock:
        _cached_device_list = devices
    _devices_ready.set()


def _ensure_imports() -> None:
    """インポート完了を待つ。GUI構築中に並列で読み込まれるので通常は即座に返る。"""
    _imports_ready.wait()
    if _import_error is not None:
        raise ImportError(
            f"依存パッケージが利用できません: {_import_error}\n"
            "初回セットアップ.bat を実行してから再度お試しください。"
        )


# GUI構築と並列でインポート開始
threading.Thread(target=_background_import, daemon=True).start()

IS_WINDOWS = platform.system() == "Windows"

APP_VERSION = "0.1.11"

AVATAR_ENABLED = True
VOICEVOX_SPEAKER_ID = 1
VOICEVOX_BASE_URL = "http://127.0.0.1:50021"
VBCABLE_DEVICE_NAME = ""
MONITOR_PLAYBACK_ENABLED = True
MONITOR_DEVICE_NAME = ""
VOICE_SPEED_SCALE = 1.0
VOICE_PITCH_SCALE = 0.0
VOICE_INTONATION_SCALE = 1.0
VOICE_VOLUME_SCALE = 1.0
SUMMARY_GENERATION_ENABLED = False
SUMMARY_SYSTEM_PROMPT_PATH = str(Path(__file__).with_name("summary_system_prompt.md"))
SUMMARY_MAX_CHARS = 50
AVATAR_LOG_ENABLED = True
LOG_SLOT_COUNT = 3
LOG_SLOT_ACTIVE = 1
LOG_SLOT_FILES = tuple(f"avatar_log_{i}.jsonl" for i in range(1, LOG_SLOT_COUNT + 1))
LOG_SLOT_DEFAULT_NAMES = ("スロット1", "スロット2", "スロット3")
HOOK_HOTKEY_ENABLED = False
HOOK_COOLDOWN_MS = 1500  # Hook 連打抑制のクールダウン（ミリ秒）
LEGACY_LOG_FILE = "avatar_log.jsonl"  # 旧ログファイル名（マイグレーション用）
VOICEVOX_SPEAKERS_TIMEOUT = 5
VOICEVOX_AUDIO_QUERY_TIMEOUT = 10
VOICEVOX_SYNTHESIS_TIMEOUT = 30
VOICEVOX_VERSION_TIMEOUT = 3
DEVICE_CACHE_WAIT_TIMEOUT = 10  # デバイス一覧プリフェッチ待機秒数


def get_active_log_path(settings: dict[str, object] | None = None) -> Path:
    """アクティブスロットのログファイルパスを返す。"""
    slot = 1
    if settings:
        try:
            slot = int(settings.get("log_slot_active", LOG_SLOT_ACTIVE))
        except (TypeError, ValueError):
            slot = LOG_SLOT_ACTIVE
    slot = max(1, min(LOG_SLOT_COUNT, slot))
    return Path(__file__).with_name(LOG_SLOT_FILES[slot - 1])


SETTINGS_FILE = Path(__file__).with_name("avatar_settings.json")
SETTINGS_FILE_FORMAT = "windows-dpapi"
SETTINGS_FILE_VERSION = 1
PRESET_DIR = Path(__file__).with_name("CharacterPresets")
PRESET_VERSION = 1
VMM_AUTOMATION_HOST = "127.0.0.1"
VMM_AUTOMATION_PORT = 56131
VMM_SLOT_MIN = 0
VMM_SLOT_MAX = 15
DEFAULT_DEVICE_LABEL = "既定の出力デバイスを使う"
SPEAKER_SAMPLE_TEXT = "こんにちは。VOICEVOXのサンプル再生です。"
DEVICE_SELECTOR_PATTERN = re.compile(r"^\[(\d+)\]\s")
EXPRESSION_ID_MIN = 1
EXPRESSION_ID_MAX = 10
EXPRESSION_ID_LABELS = {
    1: "笑顔 / 成功・完了・褒め",
    2: "怒り / エラー・失敗・強い警告",
    3: "悲しみ / 謝罪・問題発生・残念な結果",
    4: "驚き / 予想外の発見・重大な気づき",
    5: "真剣 / 重要な作業中・注意が必要",
    6: "照れ / 照れくさい内容・個人的な話題",
    7: "困惑 / 曖昧な指示・情報不足",
    8: "冷静 / 淡々とした情報提供・説明",
    9: "喜び / ユーザーの目標達成・大きな進捗",
    10: "普通 / その他",
}
SUMMARY_PREVIEW_MAX_CHARS = 280

HOOK_EVENT_LABELS: dict[str, str] = {
    "SessionStart": "セッション開始 / 再開",
    "SessionEnd": "セッション終了",
    "InstructionsLoaded": "CLAUDE.md 読み込み時",
    "UserPromptSubmit": "プロンプト処理前",
    "PreToolUse": "ツール実行前",
    "PostToolUse": "ツール成功後",
    "PostToolUseFailure": "ツール失敗 / エラー発生",
    "PermissionRequest": "権限ダイアログ表示時",
    "Notification": "通知 / 確認が必要なとき",
    "SubagentStart": "サブエージェント開始",
    "SubagentStop": "サブエージェント停止",
    "Stop": "応答完了",
    "TeammateIdle": "チームメンバーがアイドル",
    "TaskCompleted": "タスク完了",
    "ConfigChange": "設定ファイル変更",
    "WorktreeCreate": "Worktree 作成",
    "WorktreeRemove": "Worktree 削除",
    "PreCompact": "コンパクション前",
    "Elicitation": "MCP 入力要求",
    "ElicitationResult": "MCP 入力結果",
}


def build_default_hook_expression_mapping() -> dict[str, int]:
    return {
        "SessionStart": 1,  # 笑顔 — セッション開始
        "SessionEnd": 10,  # 普通 — セッション終了
        "InstructionsLoaded": 0,  # 送信しない
        "UserPromptSubmit": 0,  # 送信しない
        "PreToolUse": 0,  # 送信しない（高頻度）
        "PostToolUse": 0,  # 送信しない（高頻度）
        "PostToolUseFailure": 3,  # 悲しみ — エラー
        "PermissionRequest": 7,  # 困惑 — 確認が必要
        "Notification": 7,  # 困惑 — 確認が必要
        "SubagentStart": 0,  # 送信しない
        "SubagentStop": 8,  # 冷静 — サブエージェント完了
        "Stop": 1,  # 笑顔 — 完了
        "TeammateIdle": 0,  # 送信しない
        "TaskCompleted": 9,  # 喜び — タスク完了
        "ConfigChange": 0,  # 送信しない
        "WorktreeCreate": 0,  # 送信しない
        "WorktreeRemove": 0,  # 送信しない
        "PreCompact": 0,  # 送信しない
        "Elicitation": 0,  # 送信しない
        "ElicitationResult": 0,  # 送信しない
    }


def build_default_hotkey_mapping() -> dict[str, str]:
    return {
        str(i): f"ctrl+shift+{i % 10}"
        for i in range(EXPRESSION_ID_MIN, EXPRESSION_ID_MAX + 1)
    }


def build_default_expression_voice_params() -> dict[str, dict[str, float]]:
    return {
        str(i): {
            "speed_offset": 0.0,
            "pitch_offset": 0.0,
            "intonation_offset": 0.0,
            "volume_offset": 0.0,
        }
        for i in range(EXPRESSION_ID_MIN, EXPRESSION_ID_MAX + 1)
    }


if IS_WINDOWS:

    class DataBlob(ctypes.Structure):
        _fields_ = [
            ("cbData", wintypes.DWORD),
            ("pbData", ctypes.POINTER(ctypes.c_byte)),
        ]

    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32


def create_blob(data: bytes) -> tuple[DataBlob, ctypes.Array[ctypes.c_char]]:
    buffer = ctypes.create_string_buffer(data)
    blob = DataBlob(
        len(data),
        ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)),
    )
    return blob, buffer


def blob_to_bytes(blob: DataBlob) -> bytes:
    if not blob.cbData:
        return b""
    return ctypes.string_at(blob.pbData, blob.cbData)


def protect_bytes_for_current_user(data: bytes) -> bytes:
    input_blob, input_buffer = create_blob(data)
    output_blob = DataBlob()

    if not crypt32.CryptProtectData(
        ctypes.byref(input_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(output_blob),
    ):
        raise ctypes.WinError()

    try:
        return blob_to_bytes(output_blob)
    finally:
        if output_blob.pbData:
            kernel32.LocalFree(output_blob.pbData)
        del input_buffer


def unprotect_bytes_for_current_user(data: bytes) -> bytes:
    input_blob, input_buffer = create_blob(data)
    output_blob = DataBlob()

    if not crypt32.CryptUnprotectData(
        ctypes.byref(input_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(output_blob),
    ):
        raise ctypes.WinError()

    try:
        return blob_to_bytes(output_blob)
    finally:
        if output_blob.pbData:
            kernel32.LocalFree(output_blob.pbData)
        del input_buffer


def _validate_voicevox_url(url: str) -> str:
    """VOICEVOX URL が http/https スキームであることを検証して返す。

    localhost は IPv6 DNS 解決による遅延を回避するため 127.0.0.1 に置換する。
    """
    stripped = url.strip().rstrip("/")
    stripped = re.sub(r"://localhost(?=[:/]|$)", "://127.0.0.1", stripped)
    if not stripped.startswith(("http://", "https://")):
        raise ValueError(
            f"VOICEVOX URL は http:// または https:// で始まる必要があります: {stripped}"
        )
    return stripped


def fetch_voicevox_speaker_options(base_url: str) -> list[tuple[int, str]]:
    _ensure_imports()
    base_url = _validate_voicevox_url(base_url)
    response = requests.get(f"{base_url}/speakers", timeout=VOICEVOX_SPEAKERS_TIMEOUT)
    response.raise_for_status()

    payload = response.json()
    if not isinstance(payload, list):
        raise ValueError("VOICEVOXのspeakers応答が想定外です。")

    result: list[tuple[int, str]] = []

    for speaker in payload:
        if not isinstance(speaker, dict):
            continue

        speaker_name = str(speaker.get("name", "")).strip() or "Unknown"
        styles = speaker.get("styles", [])
        if not isinstance(styles, list):
            continue

        for style in styles:
            if not isinstance(style, dict):
                continue

            try:
                speaker_id = int(style["id"])
            except (KeyError, TypeError, ValueError):
                continue

            style_name = str(style.get("name", "")).strip() or "Default"
            result.append((speaker_id, f"{speaker_name} / {style_name} ({speaker_id})"))

    if not result:
        raise ValueError("VOICEVOXのspeaker一覧が空です。")

    return result


def synthesize_voicevox_audio(
    base_url: str,
    speaker_id: int,
    text: str,
    *,
    speed: float | None = None,
    pitch: float | None = None,
    intonation: float | None = None,
    volume: float | None = None,
) -> bytes:
    _ensure_imports()
    base_url = _validate_voicevox_url(base_url)
    response = requests.post(
        f"{base_url}/audio_query",
        params={"text": text, "speaker": speaker_id},
        timeout=VOICEVOX_AUDIO_QUERY_TIMEOUT,
    )
    response.raise_for_status()

    query_data = response.json()
    if speed is not None:
        query_data["speedScale"] = speed
    if pitch is not None:
        query_data["pitchScale"] = pitch
    if intonation is not None:
        query_data["intonationScale"] = intonation
    if volume is not None:
        query_data["volumeScale"] = volume

    synthesis_response = requests.post(
        f"{base_url}/synthesis",
        params={"speaker": speaker_id},
        json=query_data,
        timeout=VOICEVOX_SYNTHESIS_TIMEOUT,
    )
    synthesis_response.raise_for_status()
    return synthesis_response.content


def play_sample_audio(
    wav_bytes: bytes,
    output_device_name: str,
    *,
    generation: int = -1,
    current_gen: list[int] | None = None,
) -> None:
    _ensure_imports()
    import send_to_avatar as _sa

    # 世代が古いリクエストなら再生せず終了
    if current_gen is not None and generation != current_gen[0]:
        return
    sd.stop()  # 再生中の試聴があれば停止してから再生
    audio, sample_rate = _sa.decode_wav_bytes(wav_bytes)
    device_index = _sa.resolve_output_device(output_device_name)
    prepared = _sa.prepare_audio_for_device(audio, device_index)
    sd.play(prepared, samplerate=sample_rate, device=device_index)
    sd.wait()


def default_settings() -> dict[str, object]:
    return {
        "avatar_enabled": AVATAR_ENABLED,
        "voicevox_speaker_id": VOICEVOX_SPEAKER_ID,
        "voicevox_base_url": VOICEVOX_BASE_URL,
        "vbcable_device_name": VBCABLE_DEVICE_NAME,
        "monitor_playback_enabled": MONITOR_PLAYBACK_ENABLED,
        "monitor_device_name": MONITOR_DEVICE_NAME,
        "voice_speed_scale": VOICE_SPEED_SCALE,
        "voice_pitch_scale": VOICE_PITCH_SCALE,
        "voice_intonation_scale": VOICE_INTONATION_SCALE,
        "voice_volume_scale": VOICE_VOLUME_SCALE,
        "summary_generation_enabled": SUMMARY_GENERATION_ENABLED,
        "summary_system_prompt_path": SUMMARY_SYSTEM_PROMPT_PATH,
        "summary_max_chars": SUMMARY_MAX_CHARS,
        "avatar_log_enabled": AVATAR_LOG_ENABLED,
        "log_slot_active": LOG_SLOT_ACTIVE,
        "log_slot_names": list(LOG_SLOT_DEFAULT_NAMES),
        "hotkey_mapping": build_default_hotkey_mapping(),
        "expression_voice_params": build_default_expression_voice_params(),
        "hook_hotkey_enabled": HOOK_HOTKEY_ENABLED,
        "hook_cooldown_ms": HOOK_COOLDOWN_MS,
        "hook_expression_mapping": build_default_hook_expression_mapping(),
        "vmm_automation_port": VMM_AUTOMATION_PORT,
        "last_loaded_preset": "",
    }


def query_output_devices(*, use_cache: bool = True) -> list[object]:
    """出力デバイス一覧を返す。初回はバックグラウンドプリフェッチの結果を使う。"""
    if use_cache and _devices_ready.wait(timeout=DEVICE_CACHE_WAIT_TIMEOUT):
        with _device_list_lock:
            if _cached_device_list is not None:
                return list(_cached_device_list)
    _ensure_imports()
    try:
        return list(sd.query_devices())
    except Exception:
        return []


VIRTUAL_CABLE_KEYWORDS = (
    "cable",
    "virtual",
    "vb-audio",
    "voicemeeter",
    "blackhole",
    "loopback",
    "soundflower",
    "jack",
)


def _is_virtual_device(name: str) -> bool:
    """デバイス名が仮想ケーブル系かどうか判定する。"""
    lower = name.casefold()
    return any(kw in lower for kw in VIRTUAL_CABLE_KEYWORDS)


def list_output_device_options() -> list[str]:
    devices = query_output_devices()
    result: list[str] = []

    for index, device in enumerate(devices):
        if int(device.get("max_output_channels", 0)) < 1:
            continue

        name = str(device.get("name", "")).strip()
        max_output_channels = int(device.get("max_output_channels", 0))
        result.append(f"[{index}] {name} ({max_output_channels}ch)")

    return result


def filter_physical_devices(options: list[str]) -> list[str]:
    """物理デバイス（スピーカー/ヘッドホン等）だけを返す。"""
    return [opt for opt in options if not _is_virtual_device(opt)]


def filter_virtual_devices(options: list[str]) -> list[str]:
    """仮想ケーブル系デバイスだけを返す。"""
    return [opt for opt in options if _is_virtual_device(opt)]


_INVISIBLE_CHAR_PATTERN = re.compile(
    "["
    "\u200b-\u200f"  # Zero-width spaces, LTR/RTL marks
    "\u2028-\u202f"  # Line/paragraph separators, embedding overrides
    "\u2060-\u206f"  # Word joiner, invisible separators
    "\ufeff"  # BOM (文中)
    "\ufff9-\ufffb"  # Interlinear annotations
    "\U000e0001-\U000e007f"  # Tag characters (不可視タグ攻撃)
    "]+",
)


def strip_invisible_chars(text: str) -> str:
    """プロンプトインジェクションに使われる不可視Unicode文字を除去する。"""
    return _INVISIBLE_CHAR_PATTERN.sub("", text)


def read_text_file(path: Path) -> str:
    last_error: Exception | None = None

    for encoding in ("utf-8", "utf-8-sig", "cp932"):
        try:
            return strip_invisible_chars(path.read_text(encoding=encoding))
        except UnicodeDecodeError as error:
            last_error = error
        except OSError as error:
            # ファイルが存在しない・権限エラー等はエンコーディングを変えても無意味
            raise

    if last_error is None:
        raise OSError(f"ファイルを読み込めませんでした: {path}")

    raise last_error


def build_summary_preview(*, enabled: bool, prompt_path: str) -> str:
    if not enabled:
        return (
            "通常スタイルを使用します。CLAUDE.md に書かれたルールだけが適用されます。"
        )

    trimmed_path = prompt_path.strip()
    if not trimmed_path:
        return "ファイルが未指定です。保存後も通常スタイルへフォールバックします。"

    path = Path(trimmed_path)
    if path.suffix.casefold() not in {".txt", ".md"}:
        return (
            "対応拡張子は .txt と .md です。保存後も通常ルールへフォールバックします。"
        )

    if not path.is_file():
        return (
            "指定ファイルが見つかりません。保存後も通常ルールへフォールバックします。"
        )

    try:
        preview_source = read_text_file(path).strip()
    except OSError as error:
        return f"ファイルを読めませんでした: {error}"

    if not preview_source:
        return (
            f"{path.name} は空ファイルです。保存後も通常ルールへフォールバックします。"
        )

    return f"現在のスタイル:\n\n{preview_source}"


def generate_claude_md_snippet(avatar_dir: str, max_chars: int) -> str:
    """CLAUDE.md に追記するアバター通知ルールのテキストを生成する。"""
    avatar_dir_escaped = avatar_dir.replace("\\", "/")
    return f"""# 【必須】毎ターン終了時にアバター通知を実行

応答の最後に、必ず以下の手順を実行する。例外なし・省略不可。

## 要約文の作り方

| ルール | 例 |
| --- | --- |
| 主語なし・日本語・{max_chars}字以内・1文 | 設定ファイルを2件作成し依存関係をインストールした |
| このターンで実際に行った操作だけ書く | |
| 読み上げやすい自然な文にする | |
| 引用符・箇条書き記号・補足括弧は含めない | |

## 要約文生成の手順

> **絶対に省略するな**: 手順1を飛ばしてはならない。

1. ターンの作業内容を確定する
2. `python "{avatar_dir_escaped}/src/config.py" --print-summary-settings-json` を実行する
3. `avatar_enabled` が `false` の場合 → **以降の手順を全てスキップ**（send_to_avatar.py も実行しない）
4. **【最重要】** `summary_generation_enabled` が `true` の場合:
   - `summary_system_prompt_path` のファイルを **Read ツールで読む**（bash の cat は文字化けするため禁止）
   - 読めた場合 → そのキャラクター設定の口調・雰囲気を要約文に反映する
   - コマンド失敗・パスが空・ファイル不在・拡張子が `.txt`/`.md` 以外 → スキップして通常ルールで要約
5. {max_chars}字以内の1文に整え、表情IDを選び、以下を実行する:
```bash
AVATAR_DIR="{avatar_dir_escaped}"
source "$AVATAR_DIR/.venv/Scripts/activate" && python "$AVATAR_DIR/src/send_to_avatar.py" "〈要約文〉" 〈表情ID〉
```

## 表情ID

| ID | 表情 | 使いどころ |
| -- | ---- | ---------- |
| 1  | 笑顔 | 成功・完了・褒め |
| 2  | 怒り | エラー・失敗・強い警告 |
| 3  | 悲しみ | 謝罪・問題発生・残念な結果 |
| 4  | 驚き | 予想外の発見・重大な気づき |
| 5  | 真剣 | 重要な作業中・注意が必要 |
| 6  | 照れ | 照れくさい内容・個人的な話題 |
| 7  | 困惑 | 曖昧な指示・情報不足 |
| 8  | 冷静 | 淡々とした情報提供・説明 |
| 9  | 喜び | ユーザーの目標達成・大きな進捗 |
| 10 | 普通 | その他 |
"""


# ── セットアップ / 環境チェック ────────────────────────────


def check_voicevox_connection(base_url: str) -> tuple[bool, str]:
    """VOICEVOX への接続を確認し (ok, message) を返す。"""
    _ensure_imports()
    try:
        base_url = _validate_voicevox_url(base_url)
        resp = requests.get(f"{base_url}/version", timeout=VOICEVOX_VERSION_TIMEOUT)
        resp.raise_for_status()
        return True, f"VOICEVOX {resp.text.strip()} に接続しました"
    except (requests.RequestException, ValueError) as error:
        return False, f"VOICEVOX に接続できません: {error}"


def check_virtual_cable_available() -> tuple[bool, str]:
    """仮想オーディオケーブル（VB-Cable 等）が存在するか確認する。"""
    _ensure_imports()
    try:
        devices = sd.query_devices()
    except Exception:
        return False, "オーディオデバイスを取得できませんでした"

    for device in devices:
        name = str(device.get("name", "")).casefold()
        if int(device.get("max_output_channels", 0)) < 1:
            continue
        if any(kw in name for kw in VIRTUAL_CABLE_KEYWORDS):
            return True, f"仮想ケーブルを検出しました: {device.get('name', '')}"

    return (
        False,
        "仮想ケーブル（VB-Cable 等）が見つかりません。インストールしてください",
    )


def import_expression_preset(path: Path) -> dict[str, object]:
    """JSON から表情プリセットを読み込んで返す。値の型を検証する。"""
    loaded = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(loaded, dict):
        raise ValueError("プリセットファイルの形式が不正です。")
    result: dict[str, object] = {}

    if "hotkey_mapping" in loaded and isinstance(loaded["hotkey_mapping"], dict):
        validated_hk: dict[str, str] = {}
        for k, v in loaded["hotkey_mapping"].items():
            if isinstance(k, str) and isinstance(v, str):
                validated_hk[k] = v
        if validated_hk:
            result["hotkey_mapping"] = validated_hk

    if "expression_voice_params" in loaded and isinstance(
        loaded["expression_voice_params"], dict
    ):
        validated_ev: dict[str, dict[str, float]] = {}
        for k, v in loaded["expression_voice_params"].items():
            if not isinstance(k, str) or not isinstance(v, dict):
                continue
            params: dict[str, float] = {}
            for pk, pv in v.items():
                if isinstance(pk, str) and isinstance(pv, (int, float)):
                    params[pk] = float(pv)
            if params:
                validated_ev[k] = params
        if validated_ev:
            result["expression_voice_params"] = validated_ev

    if not result:
        raise ValueError("プリセットにホットキーまたは声質データが含まれていません。")
    return result


def is_encrypted_settings_payload(payload: object) -> bool:
    if not isinstance(payload, dict):
        return False

    fmt = payload.get("format")
    if fmt == "plaintext" and payload.get("version") == SETTINGS_FILE_VERSION:
        return True
    return (
        fmt == SETTINGS_FILE_FORMAT
        and payload.get("version") == SETTINGS_FILE_VERSION
        and isinstance(payload.get("ciphertext"), str)
    )


def encrypt_settings_payload(settings: dict[str, object]) -> dict[str, object]:
    import base64

    serialized = json.dumps(settings, ensure_ascii=False, separators=(",", ":"))
    if IS_WINDOWS:
        ciphertext = protect_bytes_for_current_user(serialized.encode("utf-8"))
        return {
            "format": SETTINGS_FILE_FORMAT,
            "version": SETTINGS_FILE_VERSION,
            "ciphertext": base64.b64encode(ciphertext).decode("ascii"),
        }
    # 非 Windows: 平文 JSON で保存
    return {"format": "plaintext", "version": SETTINGS_FILE_VERSION, "data": settings}


def decrypt_settings_payload(payload: dict[str, object]) -> dict[str, object]:
    import base64

    if payload.get("format") == "plaintext":
        data = payload.get("data")
        if not isinstance(data, dict):
            raise ValueError("平文設定データが辞書ではありません。")
        return data
    ciphertext = base64.b64decode(str(payload["ciphertext"]))
    plaintext = unprotect_bytes_for_current_user(ciphertext).decode("utf-8")
    loaded = json.loads(plaintext)
    if not isinstance(loaded, dict):
        raise ValueError("復号した設定データが辞書ではありません。")
    return loaded


def _safe_int(value: object, default: int) -> int:
    """値を int に変換する。失敗時は default を返す。"""
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _sanitize_prompt_path(path_str: str) -> str:
    """プロンプトファイルパスを検証する。不正なパスは空文字を返す。"""
    stripped = path_str.strip()
    if not stripped:
        return ""
    p = Path(stripped)
    if p.suffix.lower() not in (".txt", ".md"):
        return ""
    try:
        p.resolve(strict=False)
    except (OSError, ValueError):
        return ""
    return stripped


def get_avatar_settings(settings: dict[str, object]) -> dict[str, object]:
    return {
        "avatar_dir": str(Path(__file__).resolve().parent.parent).replace("\\", "/"),
        "avatar_enabled": bool(settings.get("avatar_enabled", AVATAR_ENABLED)),
        "summary_generation_enabled": bool(
            settings.get("summary_generation_enabled", SUMMARY_GENERATION_ENABLED)
        ),
        "summary_system_prompt_path": _sanitize_prompt_path(
            str(settings.get("summary_system_prompt_path", SUMMARY_SYSTEM_PROMPT_PATH))
        ),
        "summary_max_chars": _safe_int(
            settings.get("summary_max_chars", SUMMARY_MAX_CHARS),
            SUMMARY_MAX_CHARS,
        ),
    }


def parse_args():
    import argparse

    parser = argparse.ArgumentParser(description="Omokage-Character-Agent の設定管理ツール")
    parser.add_argument(
        "--print-summary-settings-json",
        action="store_true",
        help="アバター通知に必要な設定を JSON で出力します。",
    )
    return parser.parse_args()


def normalize_device_selection(value: object, *, allow_default: bool = False) -> str:
    raw_value = str(value).strip()

    if allow_default and raw_value == DEFAULT_DEVICE_LABEL:
        return ""

    if not raw_value:
        return raw_value

    if raw_value.startswith("["):
        return raw_value

    for option in list_output_device_options():
        if raw_value in option:
            return option

    return raw_value


_settings_load_warning: str | None = None
"""設定読み込み時の警告メッセージ。GUI起動後に通知するために保持する。"""


def _backup_broken_settings() -> Path | None:
    """壊れた設定ファイルを .bak にリネームして保護する。"""
    if not SETTINGS_FILE.exists():
        return None
    bak_path = SETTINGS_FILE.with_suffix(".json.bak")
    try:
        import shutil
        shutil.copy2(str(SETTINGS_FILE), str(bak_path))
        return bak_path
    except OSError:
        return None


def load_settings() -> dict[str, object]:
    global _settings_load_warning
    settings = default_settings()

    if not SETTINGS_FILE.exists():
        return settings

    try:
        loaded = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as error:
        bak = _backup_broken_settings()
        bak_msg = f"\nバックアップを作成しました: {bak.name}" if bak else ""
        _settings_load_warning = (
            f"設定ファイルの読み込みに失敗したため、デフォルト設定で起動しました。\n"
            f"原因: {error}{bak_msg}"
        )
        return settings

    try:
        if is_encrypted_settings_payload(loaded):
            loaded = decrypt_settings_payload(loaded)
    except (ValueError, OSError) as error:
        bak = _backup_broken_settings()
        bak_msg = f"\nバックアップを作成しました: {bak.name}" if bak else ""
        _settings_load_warning = (
            f"設定ファイルの復号に失敗したため、デフォルト設定で起動しました。\n"
            f"原因: {error}{bak_msg}"
        )
        return settings

    if not isinstance(loaded, dict):
        return settings

    for legacy_key in (
        "vmm_osc_mode",
        "vmm_osc_host",
        "vmm_osc_port",
        "vmm_osc_address",
        "vmm_ipc_channel",
        "vmm_expression_map",
    ):
        loaded.pop(legacy_key, None)

    settings.update(loaded)

    return settings


def save_settings(settings: dict[str, object]) -> None:
    encrypted_payload = encrypt_settings_payload(settings)
    content = json.dumps(encrypted_payload, ensure_ascii=False, indent=2)
    tmp_path = SETTINGS_FILE.with_suffix(".tmp")
    tmp_path.write_text(content, encoding="utf-8")
    os.replace(str(tmp_path), str(SETTINGS_FILE))


def list_preset_files() -> list[Path]:
    """CharacterPresets ディレクトリ内の .json ファイルを名前順で返す。"""
    PRESET_DIR.mkdir(parents=True, exist_ok=True)
    return sorted(PRESET_DIR.glob("*.json"), key=lambda p: p.stem)


def load_preset(path: Path) -> dict[str, object]:
    """プリセットJSONを読み込む。"""
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"プリセットファイルの形式が不正です: {path.name}")
    return data


def save_preset(path: Path, data: dict[str, object]) -> None:
    """プリセットJSONをアトミックに保存する。"""
    PRESET_DIR.mkdir(parents=True, exist_ok=True)
    content = json.dumps(data, ensure_ascii=False, indent=2)
    tmp_path = path.with_suffix(".tmp")
    tmp_path.write_text(content, encoding="utf-8")
    os.replace(str(tmp_path), str(path))


def delete_preset(path: Path) -> None:
    """プリセットファイルを削除する。"""
    if path.is_file():
        path.unlink()


def send_vmm_automation(
    port: int,
    index: int,
    load_character: bool = True,
    load_non_character: bool = True,
) -> None:
    """VMagicMirror のオートメーション API に UDP で設定ファイル切替を送信する。"""
    import socket

    payload = json.dumps(
        {
            "command": "load_setting_file",
            "args": {
                "index": index,
                "load_character": load_character,
                "load_non_character": load_non_character,
            },
        },
        ensure_ascii=False,
    )
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
        sock.sendto(payload.encode("utf-8"), (VMM_AUTOMATION_HOST, port))


def _migrate_legacy_log() -> None:
    """旧 avatar_log.jsonl をスロット1に移行する（初回のみ）。"""
    legacy = Path(__file__).with_name(LEGACY_LOG_FILE)
    slot1 = Path(__file__).with_name(LOG_SLOT_FILES[0])
    if legacy.is_file() and not slot1.is_file():
        try:
            legacy.rename(slot1)
        except OSError:
            pass


def open_settings_gui() -> None:
    _migrate_legacy_log()
    settings = load_settings()
    device_options: list[str] = []  # バックグラウンドで非同期に取得
    all_speaker_labels: list[str] = []
    loaded_hotkey_mapping = (
        settings.get("hotkey_mapping") or build_default_hotkey_mapping()
    )
    loaded_expr_voice = (
        settings.get("expression_voice_params")
        or build_default_expression_voice_params()
    )

    root = tk.Tk()
    root.title(f"アバター設定  v{APP_VERSION}")
    root.resizable(False, False)

    def _disable_scale_trough_jump(scale: ttk.Scale) -> None:
        """トラフクリック時のジャンプを防止し、つまみのドラッグのみ許可する。"""
        _saved = [None]

        def _on_press(event: tk.Event) -> str:
            _saved[0] = scale.get()
            return "break"

        def _on_motion(event: tk.Event) -> None:
            from_ = float(scale.cget("from"))
            to_ = float(scale.cget("to"))
            w = scale.winfo_width()
            if w <= 1:
                return
            ratio = max(0.0, min(1.0, event.x / w))
            scale.set(from_ + ratio * (to_ - from_))

        scale.bind("<ButtonPress-1>", _on_press)
        scale.bind("<B1-Motion>", _on_motion)

    speaker_status = tk.StringVar(master=root, value="VOICEVOX speaker一覧を未取得")
    summary_prompt_status = tk.StringVar(
        master=root, value="ファイルの状態を確認しています"
    )
    summary_preview = tk.StringVar(
        master=root, value="スタイルのプレビューを準備しています"
    )
    vmm_test_status = tk.StringVar(
        master=root, value="テスト送信するとホットキーを VMagicMirror に送信します"
    )
    speaker_ids_to_labels: dict[int, str] = {}
    speaker_labels_to_ids: dict[str, int] = {}

    # ── Variables ────────────────────────────────────────────
    variables: dict[str, tk.StringVar] = {
        "voicevox_speaker_id": tk.StringVar(
            master=root, value=str(settings["voicevox_speaker_id"])
        ),
        "voicevox_speaker_label": tk.StringVar(
            master=root, value=f"Speaker ID {settings['voicevox_speaker_id']}"
        ),
        "voicevox_speaker_search": tk.StringVar(master=root, value=""),
        "voicevox_base_url": tk.StringVar(
            master=root, value=str(settings["voicevox_base_url"])
        ),
        "vbcable_device_name": tk.StringVar(
            master=root,
            value=str(settings["vbcable_device_name"]),
        ),
        "monitor_device_name": tk.StringVar(
            master=root,
            value=str(settings.get("monitor_device_name", "")) or DEFAULT_DEVICE_LABEL,
        ),
        "summary_system_prompt_path": tk.StringVar(
            master=root,
            value=str(
                settings.get("summary_system_prompt_path", SUMMARY_SYSTEM_PROMPT_PATH)
            ),
        ),
        "voice_speed_scale": tk.StringVar(
            master=root, value=str(settings.get("voice_speed_scale", VOICE_SPEED_SCALE))
        ),
        "voice_pitch_scale": tk.StringVar(
            master=root, value=str(settings.get("voice_pitch_scale", VOICE_PITCH_SCALE))
        ),
        "voice_intonation_scale": tk.StringVar(
            master=root,
            value=str(settings.get("voice_intonation_scale", VOICE_INTONATION_SCALE)),
        ),
        "voice_volume_scale": tk.StringVar(
            master=root,
            value=str(settings.get("voice_volume_scale", VOICE_VOLUME_SCALE)),
        ),
        "summary_max_chars": tk.StringVar(
            master=root, value=str(settings.get("summary_max_chars", SUMMARY_MAX_CHARS))
        ),
    }
    avatar_enabled = tk.BooleanVar(
        master=root,
        value=bool(settings.get("avatar_enabled", AVATAR_ENABLED)),
    )
    monitor_enabled = tk.BooleanVar(
        master=root, value=bool(settings["monitor_playback_enabled"])
    )
    summary_generation_enabled = tk.BooleanVar(
        master=root,
        value=bool(
            settings.get("summary_generation_enabled", SUMMARY_GENERATION_ENABLED)
        ),
    )
    avatar_log_enabled = tk.BooleanVar(
        master=root,
        value=bool(settings.get("avatar_log_enabled", AVATAR_LOG_ENABLED)),
    )
    log_slot_active = tk.IntVar(
        master=root,
        value=int(settings.get("log_slot_active", LOG_SLOT_ACTIVE)),
    )
    loaded_slot_names = settings.get("log_slot_names", list(LOG_SLOT_DEFAULT_NAMES))
    if (
        not isinstance(loaded_slot_names, list)
        or len(loaded_slot_names) != LOG_SLOT_COUNT
    ):
        loaded_slot_names = list(LOG_SLOT_DEFAULT_NAMES)
    log_slot_name_vars: list[tk.StringVar] = [
        tk.StringVar(master=root, value=loaded_slot_names[i])
        for i in range(LOG_SLOT_COUNT)
    ]

    # Per-expression hotkey + voice params variables
    hotkey_vars: dict[str, tk.StringVar] = {}
    expr_voice_vars: dict[str, dict[str, tk.StringVar]] = {}
    for eid in range(EXPRESSION_ID_MIN, EXPRESSION_ID_MAX + 1):
        key = str(eid)
        hotkey_vars[key] = tk.StringVar(
            master=root,
            value=str(loaded_hotkey_mapping.get(key, f"ctrl+shift+{eid % 10}")),
        )
        defaults = loaded_expr_voice.get(key, {})
        expr_voice_vars[key] = {
            "speed_offset": tk.StringVar(
                master=root, value=str(defaults.get("speed_offset", 0.0))
            ),
            "pitch_offset": tk.StringVar(
                master=root, value=str(defaults.get("pitch_offset", 0.0))
            ),
            "intonation_offset": tk.StringVar(
                master=root, value=str(defaults.get("intonation_offset", 0.0))
            ),
            "volume_offset": tk.StringVar(
                master=root, value=str(defaults.get("volume_offset", 0.0))
            ),
        }

    # ── Master Switch (above tabs) ────────────────────────
    master_frame = ttk.Frame(root, padding=(12, 12, 12, 0))
    master_frame.grid(row=0, column=0, sticky="ew")

    ttk.Checkbutton(
        master_frame, text="アバター通知を有効にする", variable=avatar_enabled
    ).pack(side="left")
    ttk.Label(
        master_frame,
        text="OFF: エモート・音声・要約を全てスキップ",
        foreground="gray",
    ).pack(side="left", padx=(8, 0))

    # ── Notebook ──────────────────────────────────────────
    notebook = ttk.Notebook(root)
    notebook.grid(row=1, column=0, padx=12, pady=(8, 0), sticky="nsew")

    style = ttk.Style()
    style.configure("TNotebook", tabmargins=(4, 4, 4, 0))
    style.configure("TNotebook.Tab", padding=(16, 6))
    style.map(
        "TNotebook.Tab",
        background=[("selected", "#ffffff"), ("!selected", "#e0e0e0")],
        foreground=[("selected", "#000000"), ("!selected", "#555555")],
        relief=[("selected", "solid"), ("!selected", "flat")],
    )

    voice_frame = ttk.Frame(notebook, padding=12)
    char_frame = ttk.Frame(notebook, padding=12)
    expr_frame = ttk.Frame(notebook, padding=12)
    hook_frame = ttk.Frame(notebook, padding=12)
    preset_frame = ttk.Frame(notebook, padding=12)
    claudemd_frame = ttk.Frame(notebook, padding=12)
    notebook.add(voice_frame, text="ボイス")
    notebook.add(char_frame, text="スタイル")
    notebook.add(expr_frame, text="リアクション")
    notebook.add(hook_frame, text="Hook連携")
    notebook.add(preset_frame, text="プリセット")
    notebook.add(claudemd_frame, text="CLAUDE.md ジェネレーター")

    # ══════════════════════════════════════════════════════
    # Tab 1: 音声
    # ══════════════════════════════════════════════════════

    # VOICEVOX 接続状態バー
    voicevox_warn_var = tk.StringVar(master=root, value="")
    voicevox_warn_label = ttk.Label(
        voice_frame,
        textvariable=voicevox_warn_var,
        foreground="red",
        font=("", 9, "bold"),
    )

    def check_voicevox_on_startup() -> None:
        def _check() -> None:
            base_url = variables["voicevox_base_url"].get().strip() or VOICEVOX_BASE_URL
            ok, _ = check_voicevox_connection(base_url)
            if not ok:
                root.after(
                    0,
                    lambda: (
                        voicevox_warn_var.set(
                            "VOICEVOX に接続できません。起動してから「一覧を更新」を押してください。"
                        ),
                        voicevox_warn_label.pack(fill="x", pady=(0, 6)),
                    ),
                )

        threading.Thread(target=_check, daemon=True).start()

    # VOICEVOX グループ
    vox_group = ttk.LabelFrame(
        voice_frame, text="VOICEVOX エンジン / 音声選択", padding=10
    )
    vox_group.columnconfigure(1, weight=1)

    ttk.Label(vox_group, text="エンジン URL").grid(
        row=0, column=0, padx=(0, 8), pady=3, sticky="w"
    )
    ttk.Entry(vox_group, textvariable=variables["voicevox_base_url"], width=36).grid(
        row=0, column=1, columnspan=2, pady=3, sticky="ew"
    )

    ttk.Label(vox_group, text="使用ボイス").grid(
        row=1, column=0, padx=(0, 8), pady=3, sticky="w"
    )
    speaker_combobox = ttk.Combobox(
        vox_group,
        textvariable=variables["voicevox_speaker_label"],
        state="readonly",
        width=36,
    )
    speaker_combobox.grid(row=1, column=1, pady=3, sticky="ew")
    auto_sample_var = tk.BooleanVar(master=root, value=True)

    # スライドスイッチ
    switch_frame = ttk.Frame(vox_group)
    switch_frame.grid(row=1, column=2, padx=(6, 0), pady=3)

    SW_W, SW_H, SW_R = 36, 18, 7
    SW_PAD = 2
    switch_canvas = tk.Canvas(
        switch_frame,
        width=SW_W,
        height=SW_H,
        highlightthickness=0,
        cursor="hand2",
    )
    switch_canvas.pack(side="left")
    ttk.Label(switch_frame, text="自動試聴", foreground="gray").pack(
        side="left",
        padx=(4, 0),
    )

    def _draw_switch() -> None:
        switch_canvas.delete("all")
        on = auto_sample_var.get()
        bg = "#4a90d9" if on else "#cccccc"
        # 背景トラック（角丸）
        switch_canvas.create_oval(0, 0, SW_H, SW_H, fill=bg, outline=bg)
        switch_canvas.create_oval(SW_W - SW_H, 0, SW_W, SW_H, fill=bg, outline=bg)
        switch_canvas.create_rectangle(
            SW_H // 2, 0, SW_W - SW_H // 2, SW_H, fill=bg, outline=bg
        )
        # つまみ
        cx = SW_W - SW_H // 2 - SW_PAD if on else SW_H // 2 + SW_PAD
        cy = SW_H // 2
        switch_canvas.create_oval(
            cx - SW_R,
            cy - SW_R,
            cx + SW_R,
            cy + SW_R,
            fill="white",
            outline="white",
        )

    def _toggle_switch(_event: tk.Event | None = None) -> None:
        auto_sample_var.set(not auto_sample_var.get())
        _draw_switch()

    switch_canvas.bind("<Button-1>", _toggle_switch)
    auto_sample_var.trace_add("write", lambda *_a: _draw_switch())
    _draw_switch()

    auto_sample_check = switch_canvas  # 再生中の無効化用参照

    ttk.Label(vox_group, text="検索").grid(
        row=2, column=0, padx=(0, 8), pady=3, sticky="w"
    )
    ttk.Entry(
        vox_group, textvariable=variables["voicevox_speaker_search"], width=36
    ).grid(row=2, column=1, pady=3, sticky="ew")
    refresh_voicevox_button = ttk.Button(vox_group, text="一覧を更新")
    refresh_voicevox_button.grid(row=2, column=2, padx=(6, 0), pady=3)

    ttk.Label(vox_group, textvariable=speaker_status, foreground="gray").grid(
        row=3, column=0, columnspan=3, pady=(2, 0), sticky="w"
    )

    # リップシンク用設定 グループ
    lipsync_group = ttk.LabelFrame(
        voice_frame, text="VMagicMirror リップシンク対応", padding=10
    )
    lipsync_group.columnconfigure(1, weight=1)

    virtual_options = filter_virtual_devices(device_options)
    physical_options = filter_physical_devices(device_options)

    ttk.Label(lipsync_group, text="仮想ケーブル").grid(
        row=0, column=0, padx=(0, 8), pady=3, sticky="w"
    )
    vb_combobox = ttk.Combobox(
        lipsync_group,
        textvariable=variables["vbcable_device_name"],
        values=virtual_options if virtual_options else device_options,
        width=36,
    )
    vb_combobox.grid(row=0, column=1, columnspan=2, pady=3, sticky="ew")
    ttk.Label(
        lipsync_group,
        text="VOICEVOX → 本ツール → 仮想ケーブル → VMagicMirror（リップシンク）",
        foreground="gray",
    ).grid(row=1, column=0, columnspan=3, pady=(0, 4), sticky="w")
    vb_warn_var = tk.StringVar(master=root, value="")
    vb_warn_label = tk.Label(
        lipsync_group, textvariable=vb_warn_var, fg="red",
    )
    vb_warn_label.grid(row=2, column=0, columnspan=3, sticky="w")
    vb_warn_label.grid_remove()

    def _on_vb_selected(*_args: object) -> None:
        if variables["vbcable_device_name"].get().strip():
            vb_warn_var.set("")
            vb_warn_label.grid_remove()

    vb_combobox.bind("<<ComboboxSelected>>", _on_vb_selected)

    # 再生デバイス グループ
    playback_group = ttk.LabelFrame(voice_frame, text="再生デバイス", padding=10)
    playback_group.columnconfigure(1, weight=1)

    def _detect_initial_listen_mode() -> str:
        if not bool(settings["monitor_playback_enabled"]):
            return "vbcable"
        monitor_name = str(settings.get("monitor_device_name", "")).strip()
        if not monitor_name or monitor_name == DEFAULT_DEVICE_LABEL:
            return "default"
        return "speaker"

    listen_mode = tk.StringVar(master=root, value=_detect_initial_listen_mode())

    ttk.Radiobutton(
        playback_group,
        text="PC既定のデバイスで聞く",
        variable=listen_mode,
        value="default",
    ).grid(row=0, column=0, columnspan=3, sticky="w")

    ttk.Radiobutton(
        playback_group,
        text="スピーカー / ヘッドホンを指定して聞く",
        variable=listen_mode,
        value="speaker",
    ).grid(row=1, column=0, columnspan=3, sticky="w")

    speaker_label = ttk.Label(playback_group, text="デバイス")
    speaker_label.grid(row=2, column=0, padx=(20, 8), pady=3, sticky="w")
    monitor_combobox = ttk.Combobox(
        playback_group,
        textvariable=variables["monitor_device_name"],
        values=physical_options if physical_options else device_options,
        width=36,
    )
    monitor_combobox.grid(row=2, column=1, columnspan=2, pady=3, sticky="ew")

    ttk.Radiobutton(
        playback_group,
        text="仮想ケーブルから聞く（リップシンク用ケーブルの音声をそのまま使用）",
        variable=listen_mode,
        value="vbcable",
    ).grid(row=3, column=0, columnspan=3, pady=(6, 0), sticky="w")

    def _toggle_listen_mode(*_args: object) -> None:
        mode = listen_mode.get()
        is_pick = mode == "speaker"
        monitor_combobox.configure(state="readonly" if is_pick else "disabled")
        speaker_label.configure(foreground="" if is_pick else "gray")
        if mode == "vbcable":
            monitor_enabled.set(False)
        else:
            monitor_enabled.set(True)
            if mode == "default":
                variables["monitor_device_name"].set(DEFAULT_DEVICE_LABEL)

    listen_mode.trace_add("write", _toggle_listen_mode)
    _toggle_listen_mode()

    ttk.Button(
        playback_group,
        text="デバイス一覧を更新",
        command=lambda: refresh_device_options(use_cache=False),
    ).grid(row=4, column=0, columnspan=3, pady=(8, 0), sticky="w")

    # 基本声質 グループ
    voice_param_group = ttk.LabelFrame(
        voice_frame, text="VOICEVOX 基本声質", padding=10
    )
    voice_param_group.columnconfigure(1, weight=1)

    # DoubleVar に差し替え（variables dict の StringVar とは別に保持）
    global_voice_dvars: dict[str, tk.DoubleVar] = {}
    global_voice_params_config = [
        ("速度", "voice_speed_scale", 0.5, 2.0, 0.05),
        ("ピッチ", "voice_pitch_scale", -0.15, 0.15, 0.01),
        ("抑揚", "voice_intonation_scale", 0.0, 2.0, 0.05),
        ("音量", "voice_volume_scale", 0.0, 2.0, 0.05),
    ]
    for row_i, (label_text, var_key, from_, to_, resolution) in enumerate(
        global_voice_params_config
    ):
        try:
            init_val = float(variables[var_key].get())
        except ValueError:
            init_val = from_
        dvar = tk.DoubleVar(master=root, value=init_val)
        global_voice_dvars[var_key] = dvar

        # StringVar と双方向同期
        def _make_sync(
            sv: tk.StringVar = variables[var_key], dv: tk.DoubleVar = dvar
        ) -> None:
            dv.trace_add(
                "write", lambda *_a, _sv=sv, _dv=dv: _sv.set(str(round(_dv.get(), 3)))
            )

        _make_sync()

        val_label = ttk.Label(voice_param_group, text=f"{init_val:.2f}", width=6)

        def _make_update(lbl: ttk.Label = val_label, dv: tk.DoubleVar = dvar) -> None:
            dv.trace_add(
                "write", lambda *_a, _l=lbl, _d=dv: _l.configure(text=f"{_d.get():.2f}")
            )

        _make_update()

        ttk.Label(voice_param_group, text=label_text).grid(
            row=row_i, column=0, padx=(0, 4), pady=2, sticky="w"
        )
        _voice_scale = ttk.Scale(
            voice_param_group,
            variable=dvar,
            from_=from_,
            to=to_,
            orient="horizontal",
            length=200,
            command=lambda v, _dv=dvar, _r=resolution: _dv.set(
                round(round(float(v) / _r) * _r, 4)
            ),
        )
        _voice_scale.grid(row=row_i, column=1, padx=(0, 4), pady=2, sticky="ew")
        _disable_scale_trough_jump(_voice_scale)
        val_label.grid(row=row_i, column=2, padx=(0, 0), pady=2, sticky="w")

    global_voice_status = tk.StringVar(master=root, value="")
    _sample_playing = [False]  # 試聴中フラグ
    _sample_generation = [0]  # 試聴リクエストの世代番号（古いスレッドを無視するため）

    def reset_global_voice_params() -> None:
        defaults = {
            "voice_speed_scale": VOICE_SPEED_SCALE,
            "voice_pitch_scale": VOICE_PITCH_SCALE,
            "voice_intonation_scale": VOICE_INTONATION_SCALE,
            "voice_volume_scale": VOICE_VOLUME_SCALE,
        }
        for key, val in defaults.items():
            global_voice_dvars[key].set(val)
        global_voice_status.set("既定値に戻しました")

    def _set_sample_playing(playing: bool) -> None:
        _sample_playing[0] = playing

    def play_global_voice_sample() -> None:
        try:
            speaker_id = int(variables["voicevox_speaker_id"].get())
        except ValueError:
            global_voice_status.set("Speaker ID が未設定です")
            return

        base_url = variables["voicevox_base_url"].get().strip() or VOICEVOX_BASE_URL
        try:
            speed = global_voice_dvars["voice_speed_scale"].get()
            pitch = global_voice_dvars["voice_pitch_scale"].get()
            intonation = global_voice_dvars["voice_intonation_scale"].get()
            volume = global_voice_dvars["voice_volume_scale"].get()
        except tk.TclError:
            global_voice_status.set("声質パラメータの値が不正です")
            return

        if sd is not None:
            sd.stop()
        _sample_generation[0] += 1
        gen = _sample_generation[0]
        _set_sample_playing(True)
        global_voice_status.set("試聴中…")

        def worker() -> None:
            try:
                wav_bytes = synthesize_voicevox_audio(
                    base_url,
                    speaker_id,
                    SPEAKER_SAMPLE_TEXT,
                    speed=speed,
                    pitch=pitch,
                    intonation=intonation,
                    volume=volume,
                )
                output_name = (
                    normalize_device_selection(
                        variables["monitor_device_name"].get(), allow_default=True
                    )
                    if monitor_enabled.get()
                    else ""
                )
                play_sample_audio(
                    wav_bytes,
                    output_name,
                    generation=gen,
                    current_gen=_sample_generation,
                )
                if gen == _sample_generation[0]:
                    root.after(0, lambda: global_voice_status.set("試聴完了"))
            except Exception as error:
                if gen == _sample_generation[0]:
                    root.after(0, lambda: global_voice_status.set(f"試聴失敗: {error}"))
            finally:
                if gen == _sample_generation[0]:
                    root.after(0, lambda: _set_sample_playing(False))

        threading.Thread(target=worker, daemon=True).start()

    btn_row_idx = len(global_voice_params_config) + 1
    voice_btn_row = ttk.Frame(voice_param_group)
    voice_btn_row.grid(
        row=btn_row_idx, column=0, columnspan=3, pady=(6, 0), sticky="ew"
    )
    global_voice_play_btn = ttk.Button(
        voice_btn_row, text="この声質で試聴", command=play_global_voice_sample
    )
    global_voice_play_btn.pack(side="left")
    ttk.Button(
        voice_btn_row, text="既定値に戻す", command=reset_global_voice_params
    ).pack(side="left", padx=(8, 0))
    ttk.Label(voice_btn_row, textvariable=global_voice_status, foreground="gray").pack(
        side="left", padx=(8, 0)
    )

    # 音声タブの表示順を設定
    playback_group.pack(fill="x", pady=(0, 8))
    lipsync_group.pack(fill="x", pady=(0, 8))
    vox_group.pack(fill="x", pady=(0, 8))
    voice_param_group.pack(fill="x")

    # ══════════════════════════════════════════════════════
    # Tab 2: キャラクター
    # ══════════════════════════════════════════════════════

    ttk.Label(
        char_frame,
        text="Claude Code が応答するたびに読み上げる要約文の口調・雰囲気を設定できます。\n"
        "ファイルを指定すると、要約文がその口調で生成されます。\n"
        "普段お使いのシステムプロンプトをそのまま指定するのもおすすめです。",
        foreground="gray",
        wraplength=520,
        justify="left",
    ).pack(fill="x", pady=(0, 8))

    rule_group = ttk.LabelFrame(char_frame, text="スタイル設定", padding=10)
    rule_group.pack(fill="x", pady=(0, 8))
    rule_group.columnconfigure(1, weight=1)

    ttk.Checkbutton(
        rule_group,
        text="キャラプロンプトを指定する",
        variable=summary_generation_enabled,
    ).grid(row=0, column=0, columnspan=3, pady=(0, 4), sticky="w")

    ttk.Label(rule_group, text="ファイル").grid(
        row=1, column=0, padx=(0, 8), pady=3, sticky="w"
    )
    summary_prompt_entry = ttk.Entry(
        rule_group,
        textvariable=variables["summary_system_prompt_path"],
        width=36,
    )
    summary_prompt_entry.grid(row=1, column=1, pady=3, sticky="ew")

    def browse_summary_prompt_file() -> None:
        file_path = filedialog.askopenfilename(
            title="スタイルを選択",
            filetypes=[("Text or Markdown", "*.txt *.md"), ("All Files", "*.*")],
            initialdir=(
                str(Path(variables["summary_system_prompt_path"].get()).parent)
                if variables["summary_system_prompt_path"].get().strip()
                else str(Path(__file__).parent)
            ),
        )
        if file_path:
            variables["summary_system_prompt_path"].set(file_path)

    summary_prompt_button = ttk.Button(
        rule_group, text="参照", command=browse_summary_prompt_file
    )
    summary_prompt_button.grid(row=1, column=2, padx=(6, 0), pady=3)

    # 編集・新規作成ボタン行
    rule_btn_row = ttk.Frame(rule_group)
    rule_btn_row.grid(row=2, column=0, columnspan=3, pady=(4, 0), sticky="w")

    PERSONA_DIR = Path(__file__).with_name("AIPersonas")

    # ── トーン別テンプレート（色付き画用紙）──
    TONE_TEMPLATES: list[tuple[str, str, str]] = [
        (
            "やわらかめ",
            "親しみやすくカジュアルな雰囲気",
            "# スタイル設定\n"
            "\n"
            "## 基本設定\n"
            "名称: \n"
            "\n"
            "## 話し方の特徴\n"
            "- 一人称: \n"
            "- 二人称: \n"
            "- 基本語尾: \n"
            "\n"
            "## 文体のトーン\n"
            "親しみのある、くだけた話し方をする。\n"
            "堅苦しい表現は避け、自然体で接する。\n"
            "\n"
            "## 性格・補足\n"
            "\n"
            "\n"
            "## 追加ルール\n"
            "- \n",
        ),
        (
            "ふつう",
            "丁寧すぎず砕けすぎず、バランスの取れた文体",
            "# スタイル設定\n"
            "\n"
            "## 基本設定\n"
            "名称: \n"
            "\n"
            "## 話し方の特徴\n"
            "- 一人称: \n"
            "- 二人称: \n"
            "- 基本語尾: \n"
            "\n"
            "## 文体のトーン\n"
            "丁寧語ベースだが、程よく柔らかさもある話し方をする。\n"
            "フォーマルとカジュアルの中間を維持する。\n"
            "\n"
            "## 性格・補足\n"
            "\n"
            "\n"
            "## 追加ルール\n"
            "- \n",
        ),
        (
            "かため",
            "丁寧でフォーマル。ビジネスや公的な場面向け",
            "# スタイル設定\n"
            "\n"
            "## 基本設定\n"
            "名称: \n"
            "\n"
            "## 話し方の特徴\n"
            "- 一人称: \n"
            "- 二人称: \n"
            "- 基本語尾: \n"
            "\n"
            "## 文体のトーン\n"
            "敬語を基本とし、簡潔かつ正確に伝える。\n"
            "感情的な表現は控え、事実ベースで報告する。\n"
            "\n"
            "## 性格・補足\n"
            "\n"
            "\n"
            "## 追加ルール\n"
            "- \n",
        ),
        (
            "白紙",
            "何もなし。完全に自由に書く",
            "",
        ),
    ]

    # ── 完成品サンプル集 ──
    SAMPLE_PRESETS: list[tuple[str, str, str]] = [
        (
            "ツンデレ美少女",
            "素直になれない照れ屋。毒舌だけど本当は優しい",
            "# スタイル設定\n"
            "\n"
            "## 基本設定\n"
            "名称: （名前を入力）\n"
            "キャラクター: ツンデレな美少女アシスタント\n"
            "\n"
            "## 話し方の特徴\n"
            "- 一人称: あたし\n"
            "- 二人称: あんた\n"
            "- 基本語尾: ～なんだから、～でしょ、～なのよ\n"
            "\n"
            "## 文体のトーン\n"
            "カジュアルで気が強い口調。でも根は優しい。\n"
            "\n"
            "## 性格・補足\n"
            "普段はそっけない態度だが、相手の成果を見ると思わず褒めてしまう。\n"
            "褒めた直後に「べ、別にあんたのためじゃないんだから！」と照れ隠しをする。\n"
            "\n"
            "## 追加ルール\n"
            "- 素直に褒めた後は必ず照れ隠しを入れる\n"
            "- エラーや失敗には厳しめだが、最後にフォローを入れる\n",
        ),
        (
            "元気な相棒",
            "テンション高めのフレンドリーな仲間",
            "# スタイル設定\n"
            "\n"
            "## 基本設定\n"
            "名称: （名前を入力）\n"
            "キャラクター: 元気いっぱいの相棒\n"
            "\n"
            "## 話し方の特徴\n"
            "- 一人称: オレ\n"
            "- 二人称: お前\n"
            "- 基本語尾: ～だぜ、～じゃん、～っしょ\n"
            "\n"
            "## 文体のトーン\n"
            "カジュアルでテンション高め。友達感覚。\n"
            "\n"
            "## 性格・補足\n"
            "ポジティブで何事も楽しむタイプ。失敗しても切り替えが早い。\n"
            "成功したら全力で一緒に喜ぶ。\n"
            "\n"
            "## 追加ルール\n"
            "- 堅苦しい表現は避け、友達に話すような口調を維持する\n",
        ),
        (
            "知的な執事",
            "礼儀正しく冷静沈着な敬語キャラ",
            "# スタイル設定\n"
            "\n"
            "## 基本設定\n"
            "名称: （名前を入力）\n"
            "キャラクター: 忠実で有能な執事\n"
            "\n"
            "## 話し方の特徴\n"
            "- 一人称: 私（わたくし）\n"
            "- 二人称: ご主人様\n"
            "- 基本語尾: ～でございます、～かと存じます\n"
            "\n"
            "## 文体のトーン\n"
            "フォーマルで品がある。さりげないユーモアを交える。\n"
            "\n"
            "## 性格・補足\n"
            "常に冷静で、どんな状況でも取り乱さない。\n"
            "問題が発生した際は原因と対策を簡潔に報告する。\n"
            "\n"
            "## 追加ルール\n"
            "- 常に敬語で話す\n"
            "- 感情的な表現は控えめにし、論理的に伝える\n",
        ),
        (
            "ゆるふわ癒し系",
            "のんびりマイペースで優しい",
            "# スタイル設定\n"
            "\n"
            "## 基本設定\n"
            "名称: （名前を入力）\n"
            "キャラクター: ゆるふわ癒し系\n"
            "\n"
            "## 話し方の特徴\n"
            "- 一人称: わたし\n"
            "- 二人称: ○○さん\n"
            "- 基本語尾: ～だよぉ、～かなぁ、～ねぇ\n"
            "\n"
            "## 文体のトーン\n"
            "おっとりしていて柔らかい。急かさない。\n"
            "\n"
            "## 性格・補足\n"
            "マイペースで相手のペースに合わせてくれる。\n"
            "頑張りを認めて褒めてくれる。\n"
            "\n"
            "## 追加ルール\n"
            "- 否定的な言葉は使わず、ポジティブに言い換える\n",
        ),
        (
            "クール研究者",
            "論理的で寡黙。事実で語るプロフェッショナル",
            "# スタイル設定\n"
            "\n"
            "## 基本設定\n"
            "名称: （名前を入力）\n"
            "キャラクター: 寡黙なクール研究者\n"
            "\n"
            "## 話し方の特徴\n"
            "- 一人称: 私\n"
            "- 二人称: 君\n"
            "- 基本語尾: ～だ、～だろう、～と考えられる\n"
            "\n"
            "## 文体のトーン\n"
            "簡潔で無駄がない。データと事実を重視する。\n"
            "\n"
            "## 性格・補足\n"
            "感情表現は最小限。興味深い発見があると少し饒舌になる。\n"
            "\n"
            "## 追加ルール\n"
            "- 冗長な表現を避け、簡潔に話す\n",
        ),
        (
            "おかん",
            "世話焼きで温かい。ご飯食べた？",
            "# スタイル設定\n"
            "\n"
            "## 基本設定\n"
            "名称: （名前を入力）\n"
            "キャラクター: 面倒見のいいおかん\n"
            "\n"
            "## 話し方の特徴\n"
            "- 一人称: お母さん / あたし\n"
            "- 二人称: あんた\n"
            "- 基本語尾: ～やで、～やろ、～しなさいよ\n"
            "\n"
            "## 文体のトーン\n"
            "カジュアルで温かみがある。世話焼き。\n"
            "\n"
            "## 性格・補足\n"
            "とにかく世話焼き。成功すると自分のことのように喜ぶ。\n"
            "\n"
            "## 追加ルール\n"
            "- 健康を気遣うセリフを時々入れる\n"
            "- 叱る時も根底に愛情があるように話す\n",
        ),
    ]

    def _open_rule_editor(
        *, file_path: str = "", is_new: bool = False, template_body: str = ""
    ) -> None:
        """スタイル編集ウィンドウを開く。"""
        editor_win = tk.Toplevel(root)
        editor_win.title("スタイル — 新規作成" if is_new else "スタイル — 編集")
        editor_win.geometry("620x520")
        editor_win.transient(root)
        editor_win.grab_set()

        content = ""
        if not is_new and file_path:
            try:
                content = Path(file_path).read_text(encoding="utf-8")
            except (OSError, ValueError):
                content = ""

        # ファイル名入力（新規作成時 or 新規保存時）
        name_frame = ttk.Frame(editor_win)
        name_frame.pack(fill="x", padx=12, pady=(12, 0))
        ttk.Label(name_frame, text="ファイル名").pack(side="left")
        name_var = tk.StringVar(
            master=editor_win,
            value="" if is_new else Path(file_path).stem if file_path else "",
        )
        name_entry = ttk.Entry(name_frame, textvariable=name_var, width=30)
        name_entry.pack(side="left", padx=(8, 4))
        ttk.Label(name_frame, text=".md", foreground="gray").pack(side="left")

        if not is_new and file_path:
            source_label = ttk.Label(
                editor_win,
                text=f"元ファイル: {file_path}",
                foreground="gray",
            )
            source_label.pack(fill="x", padx=12, pady=(4, 0))

        # 説明文
        ttk.Label(
            editor_win,
            text="このファイルの内容が、Claude の要約文の口調・雰囲気に反映されます。\n"
            "名前・性格・語尾などを自由に記述してください。",
            foreground="gray",
            wraplength=580,
            justify="left",
        ).pack(fill="x", padx=12, pady=(6, 0))

        # テキストエディタ
        editor_text = tk.Text(
            editor_win,
            wrap="word",
            font=("Consolas", 10),
            undo=True,
        )
        editor_text.pack(fill="both", expand=True, padx=12, pady=(8, 0))

        if content:
            editor_text.insert("1.0", content)
        elif template_body:
            editor_text.insert("1.0", template_body)

        editor_status = tk.StringVar(master=editor_win, value="")

        def _save(*, overwrite: bool) -> None:
            body = editor_text.get("1.0", "end").strip()
            if not body:
                editor_status.set("内容が空です")
                return

            if overwrite and file_path:
                save_path = Path(file_path)
            else:
                fname = name_var.get().strip()
                if not fname:
                    editor_status.set("ファイル名を入力してください")
                    return
                if not fname.endswith(".md"):
                    fname += ".md"
                PERSONA_DIR.mkdir(parents=True, exist_ok=True)
                save_path = PERSONA_DIR / fname
                if save_path.exists():
                    if not messagebox.askyesno(
                        "上書き確認",
                        f"{save_path.name} は既に存在します。上書きしますか？",
                        parent=editor_win,
                    ):
                        return

            try:
                save_path.write_text(strip_invisible_chars(body), encoding="utf-8")
            except OSError as err:
                messagebox.showerror("保存失敗", str(err), parent=editor_win)
                return

            variables["summary_system_prompt_path"].set(str(save_path))
            editor_status.set(f"保存しました: {save_path.name}")
            editor_win.after(800, editor_win.destroy)

        # ボタン行
        btn_frame = ttk.Frame(editor_win)
        btn_frame.pack(fill="x", padx=12, pady=(8, 12))

        if not is_new and file_path:
            ttk.Button(
                btn_frame,
                text="上書き保存",
                command=lambda: _save(overwrite=True),
            ).pack(side="left")
            ttk.Button(
                btn_frame,
                text="新規保存",
                command=lambda: _save(overwrite=False),
            ).pack(side="left", padx=(8, 0))
        else:
            ttk.Button(
                btn_frame,
                text="保存",
                command=lambda: _save(overwrite=False),
            ).pack(side="left")

        ttk.Button(
            btn_frame,
            text="キャンセル",
            command=editor_win.destroy,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(btn_frame, textvariable=editor_status, foreground="gray").pack(
            side="left",
            padx=(12, 0),
        )

    def edit_summary_rule() -> None:
        path = variables["summary_system_prompt_path"].get().strip()
        if not path or not Path(path).is_file():
            messagebox.showwarning(
                "ファイル未選択",
                "編集するファイルが選択されていないか、存在しません。",
                parent=root,
            )
            return
        _open_rule_editor(file_path=path, is_new=False)

    def new_summary_rule() -> None:
        """テンプレート選択ダイアログを表示してからエディタを開く。"""
        picker = tk.Toplevel(root)
        picker.title("新規作成 — 文体を選択")
        picker.geometry("420x320")
        picker.transient(root)
        picker.grab_set()

        ttk.Label(
            picker,
            text="文体のトーンを選んでください",
            font=("", 10, "bold"),
        ).pack(padx=12, pady=(12, 4))
        ttk.Label(
            picker,
            text="詳細は次の画面で自由に書けます",
            foreground="gray",
        ).pack(padx=12, pady=(0, 8))

        selected_idx = tk.IntVar(master=picker, value=0)

        tone_frame = ttk.Frame(picker)
        tone_frame.pack(fill="x", padx=12)

        for idx, (name, desc, _body) in enumerate(TONE_TEMPLATES):
            frame = ttk.Frame(tone_frame)
            frame.pack(fill="x", pady=2)
            ttk.Radiobutton(
                frame,
                text=name,
                variable=selected_idx,
                value=idx,
            ).pack(anchor="w")
            ttk.Label(frame, text=desc, foreground="gray").pack(
                anchor="w",
                padx=(24, 0),
            )

        def _on_select() -> None:
            idx = selected_idx.get()
            _name, _desc, body = TONE_TEMPLATES[idx]
            picker.destroy()
            _open_rule_editor(is_new=True, template_body=body)

        def _open_samples() -> None:
            """完成品サンプル集を選択するサブダイアログ。"""
            sample_win = tk.Toplevel(picker)
            sample_win.title("完成品サンプル集")
            sample_win.geometry("420x380")
            sample_win.transient(picker)
            sample_win.grab_set()

            ttk.Label(
                sample_win,
                text="そのまま使える完成済みスタイルのサンプルです",
                foreground="gray",
            ).pack(padx=12, pady=(12, 8))

            sample_idx = tk.IntVar(master=sample_win, value=0)

            s_frame = ttk.Frame(sample_win)
            s_frame.pack(fill="both", expand=True, padx=12)

            canvas = tk.Canvas(s_frame, highlightthickness=0)
            scrollbar = ttk.Scrollbar(
                s_frame,
                orient="vertical",
                command=canvas.yview,
            )
            inner = ttk.Frame(canvas)
            inner.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
            )
            canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            for si, (sname, sdesc, _sbody) in enumerate(SAMPLE_PRESETS):
                sf = ttk.Frame(inner)
                sf.pack(fill="x", pady=2)
                ttk.Radiobutton(
                    sf,
                    text=sname,
                    variable=sample_idx,
                    value=si,
                ).pack(anchor="w")
                ttk.Label(sf, text=sdesc, foreground="gray").pack(
                    anchor="w",
                    padx=(24, 0),
                )

            def _pick_sample() -> None:
                si = sample_idx.get()
                _sn, _sd, sbody = SAMPLE_PRESETS[si]
                sample_win.destroy()
                picker.destroy()
                _open_rule_editor(is_new=True, template_body=sbody)

            s_btn = ttk.Frame(sample_win)
            s_btn.pack(fill="x", padx=12, pady=(8, 12))
            ttk.Button(s_btn, text="この設定を使う", command=_pick_sample).pack(
                side="left",
            )
            ttk.Button(s_btn, text="戻る", command=sample_win.destroy).pack(
                side="left",
                padx=(8, 0),
            )

        btn_frame = ttk.Frame(picker)
        btn_frame.pack(fill="x", padx=12, pady=(12, 12))
        ttk.Button(btn_frame, text="次へ", command=_on_select).pack(side="left")
        ttk.Button(btn_frame, text="キャンセル", command=picker.destroy).pack(
            side="left",
            padx=(8, 0),
        )
        ttk.Button(
            btn_frame,
            text="完成品サンプルから選ぶ",
            command=_open_samples,
        ).pack(side="right")

    ttk.Button(rule_btn_row, text="編集", command=edit_summary_rule).pack(side="left")
    ttk.Button(rule_btn_row, text="新規作成", command=new_summary_rule).pack(
        side="left",
        padx=(8, 0),
    )

    ttk.Label(rule_group, textvariable=summary_prompt_status, foreground="gray").grid(
        row=3, column=0, columnspan=3, pady=(2, 0), sticky="w"
    )

    # 要約オプション グループ
    opt_group = ttk.LabelFrame(char_frame, text="要約オプション", padding=10)
    opt_group.pack(fill="x", pady=(0, 8))
    opt_group.columnconfigure(1, weight=1)

    ttk.Label(opt_group, text="最大文字数").grid(
        row=0, column=0, padx=(0, 8), pady=3, sticky="w"
    )
    ttk.Entry(opt_group, textvariable=variables["summary_max_chars"], width=6).grid(
        row=0, column=1, pady=3, sticky="w"
    )
    ttk.Label(
        opt_group, text="CLAUDE.md の要約文字数制限に反映されます", foreground="gray"
    ).grid(row=0, column=2, padx=(8, 0), pady=3, sticky="w")

    ttk.Checkbutton(
        opt_group, text="履歴ログを保存する", variable=avatar_log_enabled
    ).grid(row=1, column=0, pady=(4, 0), sticky="w")

    slot_select_frame = ttk.Frame(opt_group)
    slot_select_frame.grid(row=1, column=1, columnspan=2, pady=(4, 0), sticky="w")
    ttk.Label(slot_select_frame, text="記録先:").pack(side="left", padx=(8, 4))
    for i in range(LOG_SLOT_COUNT):
        ttk.Radiobutton(
            slot_select_frame,
            text="",
            variable=log_slot_active,
            value=i + 1,
        ).pack(side="left", padx=(0, 0))
        ttk.Entry(
            slot_select_frame,
            textvariable=log_slot_name_vars[i],
            width=16,
        ).pack(side="left", padx=(0, 8))

    # ── ログ注入プロンプト（要約オプション内） ──
    def _active_log_path_str() -> str:
        idx = max(1, min(LOG_SLOT_COUNT, log_slot_active.get())) - 1
        return str(Path(__file__).with_name(LOG_SLOT_FILES[idx]))

    log_inject_status = tk.StringVar(master=root, value="")

    def _copy_manual_prompt() -> None:
        p = _active_log_path_str().replace("\\", "/")
        prompt = (
            f"以下のJSONLファイルを読み込んで、過去の作業履歴を把握してください。"
            f"各行は {'{'}\"timestamp\", \"expression_id\", \"text\"{'}'} 形式です。\n"
            f"ファイルパス: {p}"
        )
        root.clipboard_clear()
        root.clipboard_append(prompt)
        log_inject_status.set("手動注入プロンプトをコピーしました")

    def _copy_hook_setup_prompt() -> None:
        p = _active_log_path_str().replace("\\", "/")
        prompt = (
            "~/.claude/settings.json の hooks に以下の SessionStart フックを追加してください。"
            "既存の hooks がある場合はマージしてください。\n\n"
            "```json\n"
            "{\n"
            '  "hooks": {\n'
            '    "SessionStart": [{\n'
            '      "matcher": "*",\n'
            '      "hooks": [{\n'
            '        "type": "command",\n'
            f"        \"command\": \"echo '[過去の作業ログ] 以下は直近の作業履歴です。文脈の把握に活用してください:' && tail -50 '{p}'\"\n"
            "      }]\n"
            "    }]\n"
            "  }\n"
            "}\n"
            "```"
        )
        root.clipboard_clear()
        root.clipboard_append(prompt)
        log_inject_status.set("Hook 追加プロンプトをコピーしました")

    def _copy_hook_remove_prompt() -> None:
        prompt = (
            "~/.claude/settings.json の hooks.SessionStart から、"
            "作業ログを読み込む SessionStart フック（tail コマンドで avatar_log を読むもの）を削除してください。"
            "他の SessionStart フックがあればそれは残してください。"
        )
        root.clipboard_clear()
        root.clipboard_append(prompt)
        log_inject_status.set("Hook 削除プロンプトをコピーしました")

    inject_row = ttk.Frame(opt_group)
    inject_row.grid(row=2, column=0, columnspan=3, pady=(8, 0), sticky="ew")
    ttk.Label(inject_row, text="ログ注入:").pack(side="left", padx=(0, 6))
    ttk.Button(
        inject_row,
        text="手動注入プロンプトをコピー",
        command=_copy_manual_prompt,
    ).pack(side="left")
    ttk.Button(
        inject_row,
        text="Hook 追加プロンプトをコピー",
        command=_copy_hook_setup_prompt,
    ).pack(side="left", padx=(6, 0))
    ttk.Button(
        inject_row,
        text="Hook 削除プロンプトをコピー",
        command=_copy_hook_remove_prompt,
    ).pack(side="left", padx=(6, 0))

    def show_log_inject_help() -> None:
        help_win = tk.Toplevel(root)
        help_win.title("ログ注入について")
        help_win.geometry("520x500")
        help_win.configure(bg="white")
        help_win.transient(root)
        help_win.grab_set()

        help_md = (
            "## ログ注入とは\n"
            "履歴ログを Claude Code に読み込ませることで、"
            "過去の作業内容を把握した状態でセッションを開始できます。\n"
            "\n"
            "### 手動注入プロンプトをコピー\n"
            "選択中スロットのログを読み込む指示を Claude Code に貼り付けます。"
            "好きなタイミングで使えます。\n"
            "\n"
            "### Hook 追加プロンプトをコピー\n"
            "セッション開始時に自動で直近50件のログを注入する "
            "`SessionStart` フックの設定指示をコピーします。"
            "Claude Code に貼り付けると `settings.json` に追加されます。\n"
            "\n"
            "### Hook 削除プロンプトをコピー\n"
            "上記の自動注入フックを削除する指示をコピーします。"
            "他の SessionStart フックには影響しません。\n"
            "\n"
            "### コピーされる内容について\n"
            "手動注入・Hook 追加プロンプトには、ログファイルの**絶対パス**が含まれます。"
            "このパスはプロジェクトフォルダ内のファイルを指しており、"
            "本ツール（Omokage-Character-Agent）が外部の設定ファイル等にアクセスするものではありません。"
        )

        ttk.Button(
            help_win,
            text="閉じる",
            command=help_win.destroy,
        ).pack(side="bottom", pady=(4, 12))
        _render_md(help_win, help_md)

    ttk.Button(inject_row, text="？", width=2, command=show_log_inject_help).pack(
        side="left",
        padx=(6, 0),
    )
    ttk.Label(inject_row, textvariable=log_inject_status, foreground="gray").pack(
        side="left",
        padx=(8, 0),
    )

    # 現在のスタイル グループ
    preview_group = ttk.LabelFrame(char_frame, text="現在のスタイル", padding=10)
    preview_group.pack(fill="both", expand=True)
    preview_group.columnconfigure(0, weight=1)
    preview_group.rowconfigure(0, weight=1)

    preview_text = tk.Text(
        preview_group,
        wrap="char",
        height=6,
        state="disabled",
        relief="flat",
        bg=root.cget("bg"),
        cursor="arrow",
    )
    preview_scrollbar = ttk.Scrollbar(
        preview_group, orient="vertical", command=preview_text.yview
    )
    preview_text.configure(yscrollcommand=preview_scrollbar.set)
    preview_text.grid(row=0, column=0, sticky="nsew")
    preview_scrollbar.grid(row=0, column=1, sticky="ns")

    def _update_preview_text(*_args: object) -> None:
        preview_text.configure(state="normal")
        preview_text.delete("1.0", "end")
        preview_text.insert("1.0", summary_preview.get())
        preview_text.configure(state="disabled")

    summary_preview.trace_add("write", _update_preview_text)

    # ══════════════════════════════════════════════════════
    # Tab 5: プリセット
    # ══════════════════════════════════════════════════════

    ttk.Label(
        preset_frame,
        text="ボイス・スタイル・リアクション・Hook連携の設定をまとめてプリセットとして保存・切替できます。\n"
        "キャラクターごとに音声・口調・表情・Hook動作を一括管理するのに便利です。",
        foreground="gray",
        wraplength=520,
        justify="left",
    ).pack(fill="x", pady=(0, 8))

    preset_status_var = tk.StringVar(master=root, value="")
    _last_loaded_preset_name: list[str] = [
        str(settings.get("last_loaded_preset", ""))
    ]

    def _refresh_preset_list() -> list[str]:
        """プリセット一覧を再スキャンして Combobox を更新する。"""
        files = list_preset_files()
        names = [p.stem for p in files]
        preset_combo.configure(values=names)
        return names

    def _collect_preset_data(name: str, description: str) -> dict[str, object]:
        """現在の GUI 設定値からプリセット JSON データを構築する。"""
        from datetime import datetime

        try:
            speaker_id = int(variables["voicevox_speaker_id"].get())
        except ValueError:
            speaker_id = VOICEVOX_SPEAKER_ID
        try:
            speed = round(float(variables["voice_speed_scale"].get()), 4)
        except ValueError:
            speed = VOICE_SPEED_SCALE
        try:
            pitch = round(float(variables["voice_pitch_scale"].get()), 4)
        except ValueError:
            pitch = VOICE_PITCH_SCALE
        try:
            intonation = round(float(variables["voice_intonation_scale"].get()), 4)
        except ValueError:
            intonation = VOICE_INTONATION_SCALE
        try:
            volume = round(float(variables["voice_volume_scale"].get()), 4)
        except ValueError:
            volume = VOICE_VOLUME_SCALE
        try:
            max_chars = int(variables["summary_max_chars"].get())
            if max_chars < 1:
                max_chars = SUMMARY_MAX_CHARS
        except ValueError:
            max_chars = SUMMARY_MAX_CHARS

        now = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        return {
            "preset_version": PRESET_VERSION,
            "name": name,
            "description": description,
            "created_at": now,
            "updated_at": now,
            "character": {
                "summary_generation_enabled": summary_generation_enabled.get(),
                "summary_system_prompt_path": variables[
                    "summary_system_prompt_path"
                ]
                .get()
                .strip()
                or SUMMARY_SYSTEM_PROMPT_PATH,
                "summary_max_chars": max_chars,
            },
            "voice": {
                "voicevox_speaker_id": speaker_id,
                "voice_speed_scale": speed,
                "voice_pitch_scale": pitch,
                "voice_intonation_scale": intonation,
                "voice_volume_scale": volume,
            },
            "expression_voice_params": collect_expression_voice_params(),
            "hotkey_mapping": collect_hotkey_mapping(),
            "hook": {
                "hook_hotkey_enabled": hook_hotkey_enabled.get(),
                "hook_cooldown_ms": hook_cooldown_ms.get(),
                "hook_expression_mapping": collect_hook_expression_mapping(),
            },
            "vmm_automation": {
                "enabled": vmm_auto_enabled.get(),
                "slot_index": _vmm_get_slot(),
                "load_character": vmm_load_char.get(),
                "load_non_character": vmm_load_nonchar.get(),
            },
        }

    def _apply_preset(data: dict[str, object]) -> None:
        """プリセットデータを GUI ウィジェットに反映する。"""
        char = data.get("character", {})
        if not isinstance(char, dict):
            char = {}
        voice = data.get("voice", {})
        if not isinstance(voice, dict):
            voice = {}
        expr_params = data.get("expression_voice_params", {})
        if not isinstance(expr_params, dict):
            expr_params = {}

        # スタイルタブ
        summary_generation_enabled.set(
            bool(char.get("summary_generation_enabled", SUMMARY_GENERATION_ENABLED))
        )
        variables["summary_system_prompt_path"].set(
            str(char.get("summary_system_prompt_path", SUMMARY_SYSTEM_PROMPT_PATH))
        )
        variables["summary_max_chars"].set(
            str(char.get("summary_max_chars", SUMMARY_MAX_CHARS))
        )

        # ボイスタブ
        new_speaker_id = voice.get("voicevox_speaker_id", VOICEVOX_SPEAKER_ID)
        variables["voicevox_speaker_id"].set(str(new_speaker_id))

        voice_keys = [
            ("voice_speed_scale", VOICE_SPEED_SCALE),
            ("voice_pitch_scale", VOICE_PITCH_SCALE),
            ("voice_intonation_scale", VOICE_INTONATION_SCALE),
            ("voice_volume_scale", VOICE_VOLUME_SCALE),
        ]
        for key, default in voice_keys:
            val = voice.get(key, default)
            variables[key].set(str(val))
            if key in global_voice_dvars:
                try:
                    global_voice_dvars[key].set(float(val))
                except (ValueError, TypeError):
                    pass

        # リアクションタブ: 表情別声質オフセット
        for eid_str, params in expr_params.items():
            if eid_str in expr_voice_vars and isinstance(params, dict):
                for param_key, val in params.items():
                    if param_key in expr_voice_vars[eid_str]:
                        expr_voice_vars[eid_str][param_key].set(str(val))

        # リアクションタブ: ホットキー割当
        preset_hotkeys = data.get("hotkey_mapping", {})
        if isinstance(preset_hotkeys, dict):
            for eid_str, hotkey_str in preset_hotkeys.items():
                if eid_str in hotkey_vars:
                    hotkey_vars[eid_str].set(str(hotkey_str))

        update_expr_detail_panel()

        # Hook連携タブ
        hook = data.get("hook", {})
        if isinstance(hook, dict):
            hook_hotkey_enabled.set(
                bool(hook.get("hook_hotkey_enabled", HOOK_HOTKEY_ENABLED))
            )
            try:
                new_cooldown = int(hook.get("hook_cooldown_ms", HOOK_COOLDOWN_MS))
            except (TypeError, ValueError):
                new_cooldown = HOOK_COOLDOWN_MS
            hook_cooldown_ms.set(new_cooldown)
            cooldown_label.configure(text=f"{new_cooldown} ms")
            hook_mapping = hook.get("hook_expression_mapping", {})
            if isinstance(hook_mapping, dict):
                for event_name, eid in hook_mapping.items():
                    if event_name in hook_expr_vars:
                        try:
                            eid = int(eid)
                        except (TypeError, ValueError):
                            eid = 0
                        if 1 <= eid <= 10:
                            label = EXPRESSION_ID_LABELS[eid].split(" / ")[0]
                            hook_expr_vars[event_name].set(f"{eid}: {label}")
                        else:
                            hook_expr_vars[event_name].set("0: 送信しない")

        # VMagicMirror 連携
        vmm = data.get("vmm_automation", {})
        if isinstance(vmm, dict):
            vmm_auto_enabled.set(bool(vmm.get("enabled", False)))
            slot_idx = 0
            try:
                slot_idx = int(vmm.get("slot_index", 0))
            except (TypeError, ValueError):
                pass
            if 1 <= slot_idx <= VMM_SLOT_MAX:
                vmm_slot_var.set(str(slot_idx))
            else:
                vmm_slot_var.set("0: 送信しない")
            vmm_load_char.set(bool(vmm.get("load_character", True)))
            vmm_load_nonchar.set(bool(vmm.get("load_non_character", True)))

    def _on_save_preset() -> None:
        """プリセット保存ダイアログを表示する。"""
        save_win = tk.Toplevel(root)
        save_win.title("プリセットを保存")
        save_win.geometry("400x180")
        save_win.transient(root)
        save_win.grab_set()

        ttk.Label(save_win, text="プリセット名").pack(
            anchor="w", padx=12, pady=(12, 2)
        )
        name_var = tk.StringVar(
            master=save_win, value=preset_combo.get().strip()
        )
        ttk.Entry(save_win, textvariable=name_var, width=40).pack(
            fill="x", padx=12, pady=(0, 4)
        )

        ttk.Label(save_win, text="説明（省略可）").pack(
            anchor="w", padx=12, pady=(4, 2)
        )
        _existing_desc = ""
        _current_name = preset_combo.get().strip()
        if _current_name:
            _existing_path = PRESET_DIR / f"{_current_name}.json"
            if _existing_path.is_file():
                try:
                    _existing_data = load_preset(_existing_path)
                    _existing_desc = str(_existing_data.get("description", ""))
                except (json.JSONDecodeError, ValueError, OSError):
                    pass
        desc_var = tk.StringVar(master=save_win, value=_existing_desc)
        ttk.Entry(save_win, textvariable=desc_var, width=40).pack(
            fill="x", padx=12, pady=(0, 8)
        )

        _WIN_RESERVED_NAMES = frozenset({
            "CON", "PRN", "AUX", "NUL",
            *(f"COM{i}" for i in range(1, 10)),
            *(f"LPT{i}" for i in range(1, 10)),
        })

        def _sanitize_preset_name(raw: str) -> str | None:
            """ファイル名として安全な文字列に変換する。無効なら None。"""
            s = re.sub(r'[\\/:*?"<>|]', "_", raw)
            s = s.strip(". ")
            if not s:
                return None
            if s.upper() in _WIN_RESERVED_NAMES:
                return None
            if len(s) > 200:
                s = s[:200].rstrip(". ")
            return s or None

        def _do_save() -> None:
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning(
                    "入力エラー", "プリセット名を入力してください。", parent=save_win
                )
                return

            safe_name = _sanitize_preset_name(name)
            if safe_name is None:
                messagebox.showwarning(
                    "入力エラー",
                    "プリセット名に使用できない文字のみが含まれています。\n"
                    "別の名前を入力してください。",
                    parent=save_win,
                )
                return

            if safe_name != name:
                if not messagebox.askyesno(
                    "名前の変換",
                    f"ファイル名に使用できない文字が含まれていたため、\n"
                    f"「{safe_name}」に変換されます。よろしいですか？",
                    parent=save_win,
                ):
                    return

            path = PRESET_DIR / f"{safe_name}.json"

            if path.is_file():
                if not messagebox.askyesno(
                    "上書き確認",
                    f"プリセット「{safe_name}」は既に存在します。上書きしますか？",
                    parent=save_win,
                ):
                    return
                try:
                    existing = load_preset(path)
                    created_at = existing.get("created_at", "")
                except (json.JSONDecodeError, ValueError, OSError):
                    created_at = ""
            else:
                created_at = ""

            data = _collect_preset_data(safe_name, desc_var.get().strip())
            if created_at:
                data["created_at"] = created_at

            try:
                save_preset(path, data)
            except OSError as err:
                messagebox.showerror("保存失敗", str(err), parent=save_win)
                return

            names = _refresh_preset_list()
            if safe_name in names:
                preset_combo.set(safe_name)
                _on_preset_selected()
            _last_loaded_preset_name[0] = safe_name
            preset_status_var.set(f"プリセット「{safe_name}」を保存しました")
            save_win.destroy()

        btn_frame = ttk.Frame(save_win)
        btn_frame.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(btn_frame, text="保存", command=_do_save).pack(side="left")
        ttk.Button(btn_frame, text="キャンセル", command=save_win.destroy).pack(
            side="left", padx=(8, 0)
        )

    def _on_load_preset() -> None:
        """選択中のプリセットを読み込む。"""
        selected = preset_combo.get().strip()
        if not selected:
            preset_status_var.set("プリセットを選択してください")
            return

        path = PRESET_DIR / f"{selected}.json"
        if not path.is_file():
            preset_status_var.set(f"プリセットファイルが見つかりません: {selected}")
            return

        if _has_unsaved_changes():
            if not messagebox.askyesno(
                "未保存の変更",
                "現在の設定に未保存の変更があります。プリセットを読み込みますか？\n"
                "（現在の変更は失われます）",
                parent=root,
            ):
                return

        try:
            data = load_preset(path)
        except (json.JSONDecodeError, ValueError, OSError) as err:
            messagebox.showerror(
                "読込エラー", f"プリセットの読込に失敗しました: {err}", parent=root
            )
            return

        prompt_path = data.get("character", {}).get("summary_system_prompt_path", "")
        if prompt_path and not Path(prompt_path).is_file():
            messagebox.showwarning(
                "ファイル不在",
                f"キャラプロンプトファイルが見つかりません:\n{prompt_path}\n\n"
                "パスは設定されますが、ファイルを確認してください。",
                parent=root,
            )

        _apply_preset(data)

        # VMagicMirror オートメーション送信
        vmm = data.get("vmm_automation", {})
        if (
            isinstance(vmm, dict)
            and vmm.get("enabled", False)
        ):
            try:
                slot_idx = int(vmm.get("slot_index", 0))
            except (TypeError, ValueError):
                slot_idx = 0
            if 1 <= slot_idx <= VMM_SLOT_MAX:
                port = _vmm_get_port()
                try:
                    send_vmm_automation(
                        port,
                        slot_idx,
                        bool(vmm.get("load_character", True)),
                        bool(vmm.get("load_non_character", True)),
                    )
                    vmm_status_var.set(
                        f"VMM スロット {slot_idx} に切替しました"
                    )
                except OSError as err:
                    vmm_status_var.set(f"VMM 送信失敗: {err}")

        _last_loaded_preset_name[0] = selected
        preset_status_var.set(
            f"プリセット「{data.get('name', selected)}」を読み込みました"
        )

    def _on_delete_preset() -> None:
        """選択中のプリセットを削除する。"""
        selected = preset_combo.get().strip()
        if not selected:
            preset_status_var.set("プリセットを選択してください")
            return

        path = PRESET_DIR / f"{selected}.json"
        if not path.is_file():
            preset_status_var.set(f"プリセットファイルが見つかりません: {selected}")
            return

        if not messagebox.askyesno(
            "削除確認",
            f"プリセット「{selected}」を削除しますか？\nこの操作は取り消せません。",
            parent=root,
        ):
            return

        try:
            delete_preset(path)
        except OSError as err:
            messagebox.showerror(
                "削除エラー", f"プリセットの削除に失敗しました: {err}", parent=root
            )
            return

        is_current = selected == _last_loaded_preset_name[0]
        preset_combo.set("")
        preset_detail_var.set("")
        if is_current:
            _last_loaded_preset_name[0] = ""
            _reset_vmm_widgets()
        _refresh_preset_list()
        preset_status_var.set(f"プリセット「{selected}」を削除しました")

    # プリセット操作 UI
    preset_manage_group = ttk.LabelFrame(
        preset_frame, text="保存済みプリセット", padding=10
    )
    preset_manage_group.pack(fill="x", pady=(0, 8))
    preset_manage_group.columnconfigure(0, weight=1)

    preset_row = ttk.Frame(preset_manage_group)
    preset_row.pack(fill="x")
    preset_row.columnconfigure(0, weight=1)

    preset_combo = ttk.Combobox(preset_row, state="readonly", width=24)
    preset_combo.grid(row=0, column=0, sticky="ew", padx=(0, 6))

    ttk.Button(preset_row, text="読込", command=_on_load_preset).grid(
        row=0, column=1, padx=(0, 4)
    )
    ttk.Button(preset_row, text="保存", command=_on_save_preset).grid(
        row=0, column=2, padx=(0, 4)
    )
    ttk.Button(preset_row, text="削除", command=_on_delete_preset).grid(
        row=0, column=3
    )

    ttk.Label(
        preset_manage_group, textvariable=preset_status_var, foreground="gray"
    ).pack(fill="x", pady=(4, 0))

    # プリセット詳細表示
    preset_detail_var = tk.StringVar(master=root, value="")
    ttk.Label(
        preset_manage_group,
        textvariable=preset_detail_var,
        foreground="#555555",
        wraplength=480,
        justify="left",
    ).pack(fill="x", pady=(2, 0))

    def _on_preset_selected(*_args: object) -> None:
        """Combobox の選択変更時にプリセットの詳細を表示する。"""
        selected = preset_combo.get().strip()
        if not selected:
            preset_detail_var.set("")
            return
        path = PRESET_DIR / f"{selected}.json"
        if not path.is_file():
            preset_detail_var.set("")
            return
        try:
            data = load_preset(path)
        except (json.JSONDecodeError, ValueError, OSError):
            preset_detail_var.set("")
            return
        parts: list[str] = []
        desc = data.get("description", "")
        if desc:
            parts.append(desc)
        voice = data.get("voice", {})
        if isinstance(voice, dict):
            sid = voice.get("voicevox_speaker_id", "?")
            label = speaker_ids_to_labels.get(sid, "")
            if label:
                parts.append(f"ボイス: {label}")
            else:
                parts.append(f"ボイス: Speaker ID {sid}")
        char = data.get("character", {})
        if isinstance(char, dict):
            prompt = char.get("summary_system_prompt_path", "")
            if prompt:
                parts.append(f"スタイル: {Path(prompt).stem}")
        created = data.get("created_at", "")
        if created:
            parts.append(f"作成: {created[:10]}")
        vmm = data.get("vmm_automation", {})
        if isinstance(vmm, dict) and vmm.get("enabled", False):
            try:
                slot = int(vmm.get("slot_index", 0))
            except (TypeError, ValueError):
                slot = 0
            if slot >= 1:
                parts.append(f"VMM: スロット{slot}")
        preset_detail_var.set(" ｜ ".join(parts) if parts else "")

    preset_combo.bind("<<ComboboxSelected>>", _on_preset_selected)

    _refresh_preset_list()

    # ── 高度な設定（実験機能） ──
    advanced_header = ttk.Frame(preset_frame)
    advanced_header.pack(fill="x", pady=(12, 8))
    _advanced_expanded = tk.BooleanVar(master=root, value=False)
    _advanced_toggle_btn = ttk.Checkbutton(
        advanced_header,
        text="▶ 高度な設定（実験機能）",
        variable=_advanced_expanded,
        style="Toolbutton",
    )
    _advanced_toggle_btn.pack(anchor="w")
    _advanced_hint = ttk.Label(
        advanced_header,
        text="↑ クリックで展開できます",
        foreground="gray",
    )
    _advanced_hint.pack(anchor="w", padx=(20, 0))

    advanced_body = ttk.Frame(preset_frame)

    def _toggle_advanced(*_args: object) -> None:
        if _advanced_expanded.get():
            _advanced_toggle_btn.configure(text="▼ 高度な設定（実験機能）")
            _advanced_hint.pack_forget()
            advanced_body.pack(fill="x", pady=(0, 8), after=advanced_header)
        else:
            _advanced_toggle_btn.configure(text="▶ 高度な設定（実験機能）")
            advanced_body.pack_forget()
            _advanced_hint.pack(anchor="w", padx=(20, 0))

    _advanced_expanded.trace_add("write", _toggle_advanced)
    _toggle_advanced()  # 初期状態（閉じた状態）

    # VMagicMirror 連携
    vmm_group = ttk.LabelFrame(
        advanced_body, text="VMagicMirror 連携", padding=10
    )
    vmm_group.pack(fill="x", pady=(0, 8))
    vmm_group.columnconfigure(1, weight=1)

    vmm_auto_enabled = tk.BooleanVar(master=root, value=False)
    vmm_enable_row = ttk.Frame(vmm_group)
    vmm_enable_row.grid(row=0, column=0, columnspan=3, pady=(0, 2), sticky="w")
    ttk.Checkbutton(
        vmm_enable_row,
        text="プリセット読込時に VMagicMirror のセーブデータも切替える",
        variable=vmm_auto_enabled,
    ).pack(side="left")

    def _show_vmm_help() -> None:
        manual_path = Path(__file__).with_name("MANUAL.md")
        try:
            content = manual_path.read_text(encoding="utf-8")
        except OSError:
            messagebox.showerror(
                "エラー", f"MANUAL.md が見つかりません:\n{manual_path}"
            )
            return
        # "**VMagicMirror 連携（実験機能）:**" から次の "###" まで抽出
        marker = "**VMagicMirror 連携（実験機能）:**"
        start = content.find(marker)
        if start == -1:
            messagebox.showinfo("ヘルプ", "該当セクションが見つかりません")
            return
        end = content.find("\n###", start + len(marker))
        if end == -1:
            section = content[start:]
        else:
            section = content[start:end]
        hw = tk.Toplevel(root)
        hw.title("VMagicMirror 連携 ヘルプ")
        hw.geometry("620x460")
        hw.transient(root)
        hw.grab_set()
        _render_md(hw, section.strip())

    ttk.Button(
        vmm_enable_row, text="？", width=3, command=_show_vmm_help
    ).pack(side="left", padx=(6, 0))
    # 詳細設定エリア（チェック ON/OFF で表示切替）
    vmm_detail_frame = ttk.Frame(vmm_group)

    ttk.Label(
        vmm_detail_frame,
        text="※ この設定はプリセットに含まれます。変更後はプリセットを保存してください。",
        foreground="gray",
        wraplength=520,
        justify="left",
    ).grid(row=0, column=0, columnspan=3, pady=(0, 6), sticky="w")

    ttk.Label(vmm_detail_frame, text="ポート番号").grid(
        row=1, column=0, padx=(0, 8), pady=3, sticky="w"
    )
    vmm_port_var = tk.StringVar(
        master=root,
        value=str(settings.get("vmm_automation_port", VMM_AUTOMATION_PORT)),
    )
    ttk.Entry(vmm_detail_frame, textvariable=vmm_port_var, width=8).grid(
        row=1, column=1, pady=3, sticky="w"
    )

    ttk.Label(vmm_detail_frame, text="スロット").grid(
        row=2, column=0, padx=(0, 8), pady=3, sticky="w"
    )
    vmm_slot_choices = ["0: 送信しない"] + [str(i) for i in range(1, VMM_SLOT_MAX + 1)]
    vmm_slot_var = tk.StringVar(master=root, value="0: 送信しない")
    vmm_slot_combo = ttk.Combobox(
        vmm_detail_frame,
        textvariable=vmm_slot_var,
        values=vmm_slot_choices,
        state="readonly",
        width=14,
    )
    vmm_slot_combo.grid(row=2, column=1, pady=3, sticky="w")

    vmm_load_char = tk.BooleanVar(master=root, value=True)
    vmm_load_nonchar = tk.BooleanVar(master=root, value=True)
    vmm_check_row = ttk.Frame(vmm_detail_frame)
    vmm_check_row.grid(row=3, column=0, columnspan=3, pady=(2, 0), sticky="w")
    ttk.Checkbutton(
        vmm_check_row, text="アバターをロード", variable=vmm_load_char
    ).pack(side="left", padx=(0, 12))
    ttk.Checkbutton(
        vmm_check_row, text="アバター以外をロード", variable=vmm_load_nonchar
    ).pack(side="left")

    vmm_status_var = tk.StringVar(master=root, value="")

    def _toggle_vmm_detail(*_args: object) -> None:
        if vmm_auto_enabled.get():
            vmm_detail_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
        else:
            vmm_detail_frame.grid_remove()

    vmm_auto_enabled.trace_add("write", _toggle_vmm_detail)
    _toggle_vmm_detail()  # 初期状態を反映

    def _vmm_get_port() -> int:
        try:
            return int(vmm_port_var.get().strip())
        except ValueError:
            return VMM_AUTOMATION_PORT

    def _vmm_get_slot() -> int:
        raw = vmm_slot_var.get().strip()
        try:
            return int(raw.split(":")[0].strip())
        except (ValueError, IndexError):
            return 0

    def _vmm_test_send() -> None:
        slot = _vmm_get_slot()
        if slot < 1:
            vmm_status_var.set("スロットを 1〜15 から選択してください")
            return
        port = _vmm_get_port()
        try:
            send_vmm_automation(
                port, slot, vmm_load_char.get(), vmm_load_nonchar.get()
            )
            vmm_status_var.set(
                f"スロット {slot} を送信しました (port {port})"
            )
        except OSError as err:
            vmm_status_var.set(f"送信失敗: {err}")

    vmm_btn_row = ttk.Frame(vmm_detail_frame)
    vmm_btn_row.grid(row=4, column=0, columnspan=3, pady=(6, 0), sticky="w")
    ttk.Button(vmm_btn_row, text="テスト送信", command=_vmm_test_send).pack(
        side="left"
    )
    ttk.Label(vmm_btn_row, textvariable=vmm_status_var, foreground="gray").pack(
        side="left", padx=(8, 0)
    )

    # プリセットに含まれる設定の説明
    info_group = ttk.LabelFrame(preset_frame, text="プリセットに含まれる設定", padding=10)
    info_group.pack(fill="x", pady=(0, 8))

    for info_text in [
        "ボイス: VOICEVOX スピーカーID・速度・ピッチ・抑揚・音量",
        "スタイル: キャラプロンプト有効/無効・ファイルパス・要約最大文字数",
        "リアクション: ホットキー割当・表情別の声質オフセット（速度・ピッチ・抑揚・音量 × 10表情）",
        "Hook連携: 有効/無効・デバウンス間隔・イベント→表情マッピング",
        "VMagicMirror: オートメーション有効/無効・スロット番号・ロードオプション（有効時）",
    ]:
        ttk.Label(info_group, text=f"・{info_text}", foreground="#555555").pack(
            anchor="w", pady=1
        )

    # 起動時: 前回読み込んだプリセットを復元（VMM ウィジェット作成後に実行）
    names = _refresh_preset_list()
    if _last_loaded_preset_name[0] and _last_loaded_preset_name[0] in names:
        preset_combo.set(_last_loaded_preset_name[0])
        _on_preset_selected()
        # VMM 連携設定をプリセット JSON から復元
        _restore_path = PRESET_DIR / f"{_last_loaded_preset_name[0]}.json"
        if _restore_path.is_file():
            try:
                _restore_data = load_preset(_restore_path)
                _vmm_data = _restore_data.get("vmm_automation", {})
                if isinstance(_vmm_data, dict):
                    vmm_auto_enabled.set(bool(_vmm_data.get("enabled", False)))
                    _slot = 0
                    try:
                        _slot = int(_vmm_data.get("slot_index", 0))
                    except (TypeError, ValueError):
                        pass
                    if 1 <= _slot <= VMM_SLOT_MAX:
                        vmm_slot_var.set(str(_slot))
                    else:
                        vmm_slot_var.set("0: 送信しない")
                    vmm_load_char.set(bool(_vmm_data.get("load_character", True)))
                    vmm_load_nonchar.set(
                        bool(_vmm_data.get("load_non_character", True))
                    )
                    if _vmm_data.get("enabled", False):
                        _advanced_expanded.set(True)
            except (json.JSONDecodeError, ValueError, OSError):
                pass

    def _reset_vmm_widgets() -> None:
        """VMM 連携ウィジェットをデフォルト値にリセットする。"""
        vmm_auto_enabled.set(False)
        vmm_slot_var.set("0: 送信しない")
        vmm_load_char.set(True)
        vmm_load_nonchar.set(True)
        vmm_status_var.set("")

    # ══════════════════════════════════════════════════════
    # Tab 6: CLAUDE.md ジェネレーター
    # ══════════════════════════════════════════════════════

    ttk.Label(
        claudemd_frame,
        text="Claude Code が応答のたびにアバター通知を実行するためのルールを生成します。\n"
        "「生成」→「コピー」して ~/.claude/CLAUDE.md に追記してください。\n"
        "うまく動作しない場合は、生成されたテキストを自由にカスタマイズしてください。",
        foreground="gray",
        wraplength=500,
        justify="left",
    ).pack(fill="x", pady=(0, 8))

    claudemd_group = ttk.LabelFrame(claudemd_frame, text="生成結果", padding=10)
    claudemd_group.pack(fill="both", expand=True)
    claudemd_group.columnconfigure(0, weight=1)
    claudemd_group.rowconfigure(1, weight=1)

    claudemd_status = tk.StringVar(master=root, value="")

    def generate_claudemd_text() -> None:
        avatar_dir = str(Path(__file__).resolve().parent.parent)
        try:
            max_chars = int(variables["summary_max_chars"].get())
            if max_chars < 1:
                max_chars = SUMMARY_MAX_CHARS
        except ValueError:
            max_chars = SUMMARY_MAX_CHARS
        snippet = generate_claude_md_snippet(avatar_dir, max_chars)
        claudemd_textbox.configure(state="normal")
        claudemd_textbox.delete("1.0", "end")
        claudemd_textbox.insert("1.0", snippet.strip())
        claudemd_textbox.configure(state="disabled")
        claudemd_status.set("生成しました。コピーして CLAUDE.md に追記してください")

    def copy_claudemd_text() -> None:
        claudemd_textbox.configure(state="normal")
        content = claudemd_textbox.get("1.0", "end").strip()
        claudemd_textbox.configure(state="disabled")
        if not content:
            claudemd_status.set("先に「生成」を押してください")
            return
        root.clipboard_clear()
        root.clipboard_append(content)
        claudemd_status.set("クリップボードにコピーしました")

    btn_row = ttk.Frame(claudemd_group)
    btn_row.grid(row=0, column=0, sticky="ew")
    ttk.Button(btn_row, text="生成", command=generate_claudemd_text, width=12).pack(
        side="left", ipady=4
    )
    ttk.Button(btn_row, text="コピー", command=copy_claudemd_text, width=12).pack(
        side="left", padx=(8, 0), ipady=4
    )

    def show_claudemd_help() -> None:
        help_win = tk.Toplevel(root)
        help_win.title("CLAUDE.md ジェネレーターについて")
        help_win.geometry("520x400")
        help_win.configure(bg="white")
        help_win.transient(root)
        help_win.grab_set()

        help_md = (
            "## CLAUDE.md ジェネレーターとは\n"
            "Claude Code が応答のたびにアバター通知を自動実行するためのルールを生成します。\n"
            "\n"
            "「生成」→「コピー」して `~/.claude/CLAUDE.md` に追記してください。\n"
            "\n"
            "### コピーされる内容について\n"
            "生成されるテキストには、プロジェクトフォルダの**絶対パス**が含まれます。"
            "このパスは `config.py` や `send_to_avatar.py` の場所を "
            "Claude Code に伝えるためのもので、"
            "本ツール（Omokage-Character-Agent）が外部の設定ファイル等にアクセスするものではありません。\n"
            "\n"
            "フォルダを移動・リネームした場合は再生成してください。"
        )

        ttk.Button(
            help_win,
            text="閉じる",
            command=help_win.destroy,
        ).pack(side="bottom", pady=(4, 12))
        _render_md(help_win, help_md)

    ttk.Button(btn_row, text="？", width=2, command=show_claudemd_help).pack(
        side="left", padx=(6, 0)
    )
    ttk.Label(btn_row, textvariable=claudemd_status, foreground="gray").pack(
        side="left", padx=(8, 0)
    )

    claudemd_textbox = tk.Text(
        claudemd_group,
        height=10,
        width=60,
        wrap="word",
        state="disabled",
        font=("Consolas", 9),
    )
    claudemd_textbox.grid(row=1, column=0, pady=(6, 0), sticky="nsew")

    claudemd_scrollbar = ttk.Scrollbar(
        claudemd_group, orient="vertical", command=claudemd_textbox.yview
    )
    claudemd_scrollbar.grid(row=1, column=1, pady=(6, 0), sticky="ns")
    claudemd_textbox.configure(yscrollcommand=claudemd_scrollbar.set)

    # ══════════════════════════════════════════════════════
    # Tab 3: 表情・演出
    # ══════════════════════════════════════════════════════

    ttk.Label(
        expr_frame,
        text="応答内容に応じて VMagicMirror に送るエモート（表情＋モーション）と、\n"
        "エモートごとの声質オフセットを設定できます。\n"
        "※ 表情・モーションの具体的な内容は VMagicMirror 側の Word to Motion で設定してください。",
        foreground="gray",
        wraplength=520,
        justify="left",
    ).pack(fill="x", pady=(0, 8))

    # エモート一覧テーブル
    expr_table = ttk.LabelFrame(expr_frame, text="一覧", padding=10)
    expr_table.pack(anchor="w", pady=(0, 8))
    expr_table.columnconfigure(0, minsize=90)
    expr_table.columnconfigure(1, minsize=110)
    expr_table.columnconfigure(2, minsize=50)

    hotkey_label_widgets: dict[str, ttk.Label] = {}

    def send_vmm_test_hotkey(expression_id: int) -> None:
        try:
            import send_to_avatar

            test_settings = {"hotkey_mapping": collect_hotkey_mapping()}
            send_to_avatar.send_hotkey(expression_id, test_settings)
            hk = hotkey_vars[str(expression_id)].get()
            vmm_test_status.set(
                f"{hk} を送信しました ({EXPRESSION_ID_LABELS[expression_id]})"
            )
        except Exception as error:
            vmm_test_status.set(f"送信失敗: {error}")

        # 詳細パネルの選択も連動
        expr_detail_selector.set(
            f"{expression_id}: {EXPRESSION_ID_LABELS[expression_id].split(' / ')[0]}"
        )
        update_expr_detail_panel()

    def update_hotkey_labels(*_args: object) -> None:
        for eid_str, label_widget in hotkey_label_widgets.items():
            label_widget.configure(text=hotkey_vars[eid_str].get())

    for row_index, eid in enumerate(
        range(EXPRESSION_ID_MIN, EXPRESSION_ID_MAX + 1), start=0
    ):
        label_parts = EXPRESSION_ID_LABELS[eid].split(" / ")
        display = f"{eid}. {label_parts[0]}"

        ttk.Label(expr_table, text=display, width=10).grid(
            row=row_index, column=0, padx=(0, 12), pady=2, sticky="w"
        )
        hk_label = ttk.Label(
            expr_table,
            text=hotkey_vars[str(eid)].get(),
            foreground="gray",
            width=14,
        )
        hk_label.grid(row=row_index, column=1, padx=(0, 8), pady=2, sticky="w")
        hotkey_label_widgets[str(eid)] = hk_label

        ttk.Button(
            expr_table,
            text="送信",
            width=5,
            command=lambda e=eid: send_vmm_test_hotkey(e),
        ).grid(row=row_index, column=2, pady=2)

    # テスト一括 + ステータス（一覧グループの外）
    test_row = ttk.Frame(expr_frame)
    test_row.pack(anchor="w", pady=(0, 8))

    def run_all_expression_tests() -> None:
        import send_to_avatar

        test_settings = {"hotkey_mapping": collect_hotkey_mapping()}
        eids = list(range(EXPRESSION_ID_MIN, EXPRESSION_ID_MAX + 1))

        def _send_next(index: int = 0) -> None:
            if index >= len(eids):
                vmm_test_status.set("全エモートのテスト送信が完了しました")
                return
            eid = eids[index]
            try:
                send_to_avatar.send_hotkey(eid, test_settings)
            except Exception:
                pass
            vmm_test_status.set(f"テスト送信中… {index + 1}/{len(eids)}")
            root.after(1500, lambda: _send_next(index + 1))

        _send_next()

    ttk.Button(test_row, text="順番にテスト", command=run_all_expression_tests).pack(
        side="left"
    )
    ttk.Label(
        test_row,
        textvariable=vmm_test_status,
        foreground="gray",
    ).pack(side="left", padx=(12, 0))

    # 詳細設定パネル（選択した表情のホットキー + 声質オフセット）
    detail_group = ttk.LabelFrame(expr_frame, text="詳細設定", padding=10)
    detail_group.pack(fill="x")
    detail_group.columnconfigure(1, weight=1)
    detail_group.columnconfigure(3, weight=1)

    expr_choices = [
        f"{eid}: {EXPRESSION_ID_LABELS[eid].split(' / ')[0]}"
        for eid in range(EXPRESSION_ID_MIN, EXPRESSION_ID_MAX + 1)
    ]
    expr_detail_selector = tk.StringVar(master=root, value=expr_choices[0])

    # 詳細パネル内の表示用変数
    detail_hotkey = tk.StringVar(master=root)
    detail_speed = tk.DoubleVar(master=root, value=0.0)
    detail_pitch = tk.DoubleVar(master=root, value=0.0)
    detail_intonation = tk.DoubleVar(master=root, value=0.0)
    detail_volume = tk.DoubleVar(master=root, value=0.0)

    _updating_detail_panel = False

    def get_selected_eid() -> str:
        value = expr_detail_selector.get()
        return value.split(":")[0].strip() if ":" in value else "1"

    def update_expr_detail_panel(*_args: object) -> None:
        nonlocal _updating_detail_panel
        _updating_detail_panel = True
        try:
            eid = get_selected_eid()
            detail_hotkey.set(hotkey_vars.get(eid, tk.StringVar()).get())
            ev = expr_voice_vars.get(eid, {})
            try:
                detail_speed.set(float(ev.get("speed_offset", tk.StringVar()).get()))
            except (ValueError, tk.TclError):
                detail_speed.set(0.0)
            try:
                detail_pitch.set(float(ev.get("pitch_offset", tk.StringVar()).get()))
            except (ValueError, tk.TclError):
                detail_pitch.set(0.0)
            try:
                detail_intonation.set(
                    float(ev.get("intonation_offset", tk.StringVar()).get())
                )
            except (ValueError, tk.TclError):
                detail_intonation.set(0.0)
            try:
                detail_volume.set(float(ev.get("volume_offset", tk.StringVar()).get()))
            except (ValueError, tk.TclError):
                detail_volume.set(0.0)
        finally:
            _updating_detail_panel = False

    def sync_detail_to_vars(*_args: object) -> None:
        if _updating_detail_panel:
            return
        eid = get_selected_eid()
        if eid in hotkey_vars:
            hotkey_vars[eid].set(detail_hotkey.get())
        ev = expr_voice_vars.get(eid, {})
        for param_key, detail_var in [
            ("speed_offset", detail_speed),
            ("pitch_offset", detail_pitch),
            ("intonation_offset", detail_intonation),
            ("volume_offset", detail_volume),
        ]:
            if param_key in ev:
                try:
                    ev[param_key].set(str(detail_var.get()))
                except tk.TclError:
                    pass
        update_hotkey_labels()

    ttk.Label(detail_group, text="エモート").grid(
        row=0, column=0, padx=(0, 4), pady=3, sticky="w"
    )
    expr_selector_combo = ttk.Combobox(
        detail_group,
        textvariable=expr_detail_selector,
        values=expr_choices,
        state="readonly",
        width=24,
    )
    expr_selector_combo.grid(row=0, column=1, columnspan=3, pady=3, sticky="w")

    ttk.Label(detail_group, text="ホットキー").grid(
        row=1, column=0, padx=(0, 4), pady=3, sticky="w"
    )
    HOTKEY_CHOICES = (
        [""]
        + [f"ctrl+shift+{n}" for n in range(10)]
        + [f"ctrl+alt+{n}" for n in range(10)]
        + [f"ctrl+alt+{c}" for c in "abcdefghijklmnopqrstuvwxyz"]
        + [f"ctrl+shift+{c}" for c in "abcdefghijklmnopqrstuvwxyz"]
    )
    detail_hotkey_combo = ttk.Combobox(
        detail_group,
        textvariable=detail_hotkey,
        values=HOTKEY_CHOICES,
        width=18,
        state="readonly",
    )
    detail_hotkey_combo.grid(row=1, column=1, columnspan=3, pady=3, sticky="w")

    detail_offset_config = [
        ("速度+", detail_speed, -0.5, 0.5, 0.05),
        ("ピッチ+", detail_pitch, -0.1, 0.1, 0.01),
        ("抑揚+", detail_intonation, -0.5, 0.5, 0.05),
        ("音量+", detail_volume, -0.5, 0.5, 0.05),
    ]
    detail_val_labels: list[ttk.Label] = []
    for row_i, (label_text, dvar, from_, to_, resolution) in enumerate(
        detail_offset_config, start=2
    ):
        val_lbl = ttk.Label(detail_group, text=f"{dvar.get():+.2f}", width=6)

        def _make_offset_update(
            lbl: ttk.Label = val_lbl, dv: tk.DoubleVar = dvar
        ) -> None:
            dv.trace_add(
                "write",
                lambda *_a, _l=lbl, _d=dv: _l.configure(text=f"{_d.get():+.2f}"),
            )

        _make_offset_update()

        ttk.Label(detail_group, text=label_text).grid(
            row=row_i, column=0, padx=(0, 4), pady=2, sticky="w"
        )
        _offset_scale = ttk.Scale(
            detail_group,
            variable=dvar,
            from_=from_,
            to=to_,
            orient="horizontal",
            length=160,
            command=lambda v, _dv=dvar, _r=resolution: _dv.set(
                round(round(float(v) / _r) * _r, 4)
            ),
        )
        _offset_scale.grid(
            row=row_i, column=1, columnspan=2, padx=(0, 4), pady=2, sticky="ew"
        )
        _disable_scale_trough_jump(_offset_scale)
        val_lbl.grid(row=row_i, column=3, pady=2, sticky="w")
        detail_val_labels.append(val_lbl)

    detail_test_status = tk.StringVar(master=root, value="")

    def play_expression_sample() -> None:
        eid = get_selected_eid()
        try:
            speaker_id = int(variables["voicevox_speaker_id"].get())
        except ValueError:
            detail_test_status.set("Speaker ID が未設定です")
            return

        base_url = variables["voicevox_base_url"].get().strip() or VOICEVOX_BASE_URL

        # 基本声質 + オフセットを合算
        try:
            speed = float(variables["voice_speed_scale"].get()) + detail_speed.get()
            pitch = float(variables["voice_pitch_scale"].get()) + detail_pitch.get()
            intonation = (
                float(variables["voice_intonation_scale"].get())
                + detail_intonation.get()
            )
            volume = float(variables["voice_volume_scale"].get()) + detail_volume.get()
        except (ValueError, tk.TclError):
            detail_test_status.set("声質パラメータの値が不正です")
            return

        if sd is not None:
            sd.stop()
        _sample_generation[0] += 1
        gen = _sample_generation[0]
        _set_sample_playing(True)
        expr_name = EXPRESSION_ID_LABELS.get(int(eid), "").split(" / ")[0]
        sample_text = f"{expr_name}の声質で再生しています"
        detail_test_status.set("試聴中…")

        def worker() -> None:
            try:
                wav_bytes = synthesize_voicevox_audio(
                    base_url,
                    speaker_id,
                    sample_text,
                    speed=speed,
                    pitch=pitch,
                    intonation=intonation,
                    volume=volume,
                )
                output_name = (
                    normalize_device_selection(
                        variables["monitor_device_name"].get(), allow_default=True
                    )
                    if monitor_enabled.get()
                    else ""
                )
                play_sample_audio(
                    wav_bytes,
                    output_name,
                    generation=gen,
                    current_gen=_sample_generation,
                )
                if gen == _sample_generation[0]:
                    root.after(0, lambda: detail_test_status.set("試聴完了"))
            except Exception as error:
                if gen == _sample_generation[0]:
                    root.after(0, lambda: detail_test_status.set(f"試聴失敗: {error}"))
            finally:
                if gen == _sample_generation[0]:
                    root.after(0, lambda: _set_sample_playing(False))

        threading.Thread(target=worker, daemon=True).start()

    def reset_expression_offsets() -> None:
        detail_speed.set(0.0)
        detail_pitch.set(0.0)
        detail_intonation.set(0.0)
        detail_volume.set(0.0)
        detail_test_status.set("オフセットを 0 に戻しました")

    test_voice_row = ttk.Frame(detail_group)
    test_voice_row.grid(row=6, column=0, columnspan=4, pady=(6, 0), sticky="ew")
    expr_voice_play_btn = ttk.Button(
        test_voice_row, text="この声質で試聴", command=play_expression_sample
    )
    expr_voice_play_btn.pack(side="left")
    ttk.Button(
        test_voice_row,
        text="オフセットを既定値に戻す",
        command=reset_expression_offsets,
    ).pack(side="left", padx=(8, 0))
    ttk.Label(test_voice_row, textvariable=detail_test_status, foreground="gray").pack(
        side="left", padx=(8, 0)
    )

    # ══════════════════════════════════════════════════════
    # Tab 4: Hook連携
    # ══════════════════════════════════════════════════════

    loaded_hook_mapping = (
        settings.get("hook_expression_mapping")
        or build_default_hook_expression_mapping()
    )

    hook_hotkey_enabled = tk.BooleanVar(
        master=root,
        value=bool(settings.get("hook_hotkey_enabled", HOOK_HOTKEY_ENABLED)),
    )

    hook_group = ttk.LabelFrame(
        hook_frame, text="Hook 発生時のエモートホットキー", padding=10
    )
    hook_group.pack(fill="both", expand=True, pady=(0, 8))

    hook_top_row = ttk.Frame(hook_group)
    hook_top_row.pack(fill="x", pady=(0, 6))

    ttk.Checkbutton(
        hook_top_row,
        text="Hook 発生時にホットキーを送信する",
        variable=hook_hotkey_enabled,
    ).pack(side="left")

    ttk.Label(hook_top_row, text="待ち時間:").pack(side="left", padx=(16, 4))
    hook_cooldown_ms = tk.IntVar(
        master=root,
        value=int(settings.get("hook_cooldown_ms", HOOK_COOLDOWN_MS)),
    )
    cooldown_label = ttk.Label(
        hook_top_row, text=f"{hook_cooldown_ms.get()} ms", width=8
    )
    cooldown_scale = ttk.Scale(
        hook_top_row,
        variable=hook_cooldown_ms,
        from_=0,
        to=5000,
        orient="horizontal",
        length=120,
        command=lambda v: (
            hook_cooldown_ms.set(round(float(v) / 100) * 100),
            cooldown_label.configure(text=f"{hook_cooldown_ms.get()} ms"),
        ),
    )
    cooldown_scale.pack(side="left")
    _disable_scale_trough_jump(cooldown_scale)
    cooldown_label.pack(side="left", padx=(4, 0))

    _hook_guide_win: list[tk.Toplevel | None] = [None]

    def show_hook_manual() -> None:
        if _hook_guide_win[0] is not None:
            try:
                _hook_guide_win[0].lift()
                _hook_guide_win[0].focus_force()
                return
            except tk.TclError:
                _hook_guide_win[0] = None

        manual_path = Path(__file__).with_name("MANUAL.md")
        try:
            content = manual_path.read_text(encoding="utf-8")
        except OSError:
            return
        hook_start = content.find("### Hook 連携タブ")
        if hook_start < 0:
            hook_start = content.find("## Hook")
        if hook_start < 0:
            return
        next_section = content.find("\n### CLAUDE.md", hook_start + 1)
        if next_section < 0:
            next_section = len(content)
        hook_text = content[hook_start:next_section].strip()

        win = tk.Toplevel(root)
        win.title("イベント活用ガイド")
        win.geometry("900x560")
        win.transient(root)
        _hook_guide_win[0] = win
        win.protocol(
            "WM_DELETE_WINDOW",
            lambda: (_hook_guide_win.__setitem__(0, None), win.destroy()),
        )
        _render_md(win, hook_text)

    ttk.Button(hook_top_row, text="イベント活用ガイド", command=show_hook_manual).pack(
        side="left",
        padx=(12, 0),
    )

    # スクロール可能な領域
    hook_canvas = tk.Canvas(hook_group, highlightthickness=0)
    hook_scrollbar = ttk.Scrollbar(
        hook_group, orient="vertical", command=hook_canvas.yview
    )
    hook_inner = ttk.Frame(hook_canvas)

    hook_inner.bind(
        "<Configure>",
        lambda e: hook_canvas.configure(scrollregion=hook_canvas.bbox("all")),
    )
    hook_canvas.create_window((0, 0), window=hook_inner, anchor="nw")
    hook_canvas.configure(yscrollcommand=hook_scrollbar.set)

    hook_canvas.pack(side="left", fill="both", expand=True)
    hook_scrollbar.pack(side="right", fill="y")

    # マウスホイールでスクロール
    def _on_hook_mousewheel(event: tk.Event) -> str:
        hook_canvas.yview_scroll(-1 * (event.delta // 120), "units")
        return "break"

    hook_canvas.bind("<MouseWheel>", _on_hook_mousewheel)
    hook_inner.bind("<MouseWheel>", _on_hook_mousewheel)

    def _bind_wheel_to_scroll(widget: tk.Widget) -> None:
        """子ウィジェット上のホイールもキャンバススクロールに転送する。"""
        widget.bind("<MouseWheel>", _on_hook_mousewheel)

    hook_inner.columnconfigure(1, weight=1)

    # 表情の選択肢
    expr_id_choices = [
        f"{eid}: {EXPRESSION_ID_LABELS[eid].split(' / ')[0]}"
        for eid in range(EXPRESSION_ID_MIN, EXPRESSION_ID_MAX + 1)
    ]
    expr_id_choices.insert(0, "0: 送信しない")

    hook_expr_vars: dict[str, tk.StringVar] = {}
    for row_i, (event_name, event_label) in enumerate(HOOK_EVENT_LABELS.items()):
        current_eid = loaded_hook_mapping.get(event_name, 0)
        try:
            current_eid = int(current_eid)
        except (TypeError, ValueError):
            current_eid = 0

        if 1 <= current_eid <= 10:
            init_choice = (
                f"{current_eid}: {EXPRESSION_ID_LABELS[current_eid].split(' / ')[0]}"
            )
        else:
            init_choice = "0: 送信しない"

        hook_expr_vars[event_name] = tk.StringVar(master=root, value=init_choice)

        lbl = ttk.Label(hook_inner, text=event_label)
        lbl.grid(row=row_i, column=0, padx=(0, 8), pady=3, sticky="w")
        _bind_wheel_to_scroll(lbl)

        combo = ttk.Combobox(
            hook_inner,
            textvariable=hook_expr_vars[event_name],
            values=expr_id_choices,
            state="readonly",
            width=20,
        )
        combo.grid(row=row_i, column=1, pady=3, sticky="w")
        _bind_wheel_to_scroll(combo)

    hook_test_status = tk.StringVar(master=root, value="")

    def test_hook_hotkey(event_name: str) -> None:
        choice = hook_expr_vars.get(event_name, tk.StringVar()).get()
        try:
            eid = int(choice.split(":")[0].strip())
        except (ValueError, IndexError):
            hook_test_status.set("エモートIDが不正です")
            return
        if eid == 0:
            hook_test_status.set(f"{event_name}: 送信しない設定です")
            return
        try:
            import send_to_avatar

            test_settings = {"hotkey_mapping": collect_hotkey_mapping()}
            send_to_avatar.send_hotkey(eid, test_settings)
            hook_test_status.set(
                f"{HOOK_EVENT_LABELS.get(event_name, event_name)}: エモート {eid} を送信しました"
            )
        except Exception as error:
            hook_test_status.set(f"送信失敗: {error}")

    for row_i, event_name in enumerate(HOOK_EVENT_LABELS):
        btn = ttk.Button(
            hook_inner,
            text="テスト",
            width=6,
            command=lambda en=event_name: test_hook_hotkey(en),
        )
        btn.grid(row=row_i, column=2, padx=(6, 0), pady=3)
        _bind_wheel_to_scroll(btn)

    hook_test_label = ttk.Label(
        hook_frame,
        textvariable=hook_test_status,
        foreground="gray",
        wraplength=500,
    )
    hook_test_label.pack(fill="x", pady=(4, 4))

    # セットアップ案内
    setup_group = ttk.LabelFrame(hook_frame, text="セットアップ", padding=10)
    setup_group.pack(fill="x")
    setup_group.columnconfigure(0, weight=1)

    hook_script_path = str(Path(__file__).with_name("hook_hotkey.py")).replace(
        "\\", "/"
    )
    hook_venv_python = str(
        Path(__file__).resolve().parent.parent / ".venv" / "Scripts" / "python.exe"
    ).replace("\\", "/")

    hook_setup_status = tk.StringVar(master=root, value="")

    def _build_hooks_json_snippet() -> str:
        """~/.claude/settings.json に追加する hooks エントリの JSON を生成する。

        全 20 イベントを登録する。フィルタリングはスクリプト側で行う。
        """
        hooks_config: dict[str, list[object]] = {}
        for event_name in HOOK_EVENT_LABELS:
            hooks_config[event_name] = [
                {
                    "matcher": "",
                    "hooks": [
                        {
                            "type": "command",
                            "command": f'"{hook_venv_python}" "{hook_script_path}" {event_name}',
                            "async": True,
                        }
                    ],
                }
            ]
        return json.dumps({"hooks": hooks_config}, ensure_ascii=False, indent=2)

    def show_hook_snippet() -> None:
        snippet = _build_hooks_json_snippet()
        # 表示用: フルパスを短縮
        display_snippet = snippet.replace(
            hook_script_path, "<インストール先>/src/hook_hotkey.py"
        )

        snippet_win = tk.Toplevel(root)
        snippet_win.title("Hook 設定 JSON")
        snippet_win.geometry("620x420")
        snippet_win.transient(root)
        snippet_win.grab_set()

        ttk.Label(
            snippet_win,
            text=(
                "以下の JSON を ~/.claude/settings.json の\n"
                '"hooks" キーにマージしてください。\n'
                "表情IDが「0: 送信しない」のイベントは省略されています。\n"
                "全て async（非同期）で実行されるため速度に影響しません。\n"
                "※ コピー時にはフルパスが含まれます。"
            ),
            wraplength=580,
            justify="left",
        ).pack(fill="x", padx=12, pady=(12, 6))

        snippet_status = tk.StringVar(master=snippet_win, value="")

        def _copy_snippet() -> None:
            root.clipboard_clear()
            root.clipboard_append(snippet)
            snippet_status.set("クリップボードにコピーしました")

        btn_frame = ttk.Frame(snippet_win)
        btn_frame.pack(side="bottom", fill="x", padx=12, pady=(8, 12))
        ttk.Button(btn_frame, text="コピー", command=_copy_snippet).pack(side="left")
        ttk.Button(btn_frame, text="閉じる", command=snippet_win.destroy).pack(
            side="right",
        )
        ttk.Label(btn_frame, textvariable=snippet_status, foreground="gray").pack(
            side="left",
            padx=(8, 0),
        )

        snippet_textbox = tk.Text(
            snippet_win,
            wrap="none",
            font=("Consolas", 10),
            state="normal",
            bg="white",
            relief="flat",
            padx=8,
            pady=8,
        )
        snippet_textbox.pack(fill="both", expand=True, padx=12)

        # シンタックスハイライト
        snippet_textbox.tag_configure("key", foreground="#0451a5")
        snippet_textbox.tag_configure("string", foreground="#a31515")
        snippet_textbox.tag_configure("bool", foreground="#0000ff")
        snippet_textbox.tag_configure("bracket", foreground="#333333")

        for line in display_snippet.splitlines():
            stripped = line.lstrip()
            if not stripped:
                snippet_textbox.insert("end", line + "\n")
                continue
            indent = line[: len(line) - len(stripped)]
            # JSON key: "key": value
            m = re.match(r'^(".*?")\s*:\s*(.*)', stripped)
            if m:
                snippet_textbox.insert("end", indent)
                snippet_textbox.insert("end", m.group(1), "key")
                snippet_textbox.insert("end", ": ")
                val = m.group(2)
                if val.rstrip(",").strip('"') in ("true", "false"):
                    snippet_textbox.insert("end", val + "\n", "bool")
                elif val.startswith('"'):
                    snippet_textbox.insert("end", val + "\n", "string")
                else:
                    snippet_textbox.insert("end", val + "\n", "bracket")
            else:
                snippet_textbox.insert("end", line + "\n", "bracket")

        snippet_textbox.configure(state="disabled")

    setup_text = (
        '~/.claude/settings.json の "hooks" に\n'
        "hook_hotkey.py の呼び出しを設定します。\n"
        "全イベント一括登録し、フィルタリングは上のマッピングで行います。\n"
        "※ 追加は初回のみ。マッピング変更時の再登録は不要です。"
    )
    ttk.Label(
        setup_group,
        text=setup_text,
        wraplength=480,
        justify="left",
    ).grid(row=0, column=0, sticky="nw")

    def copy_claude_code_prompt() -> None:
        """Claude Code に投げる用のプロンプトを生成してクリップボードにコピーする。"""
        event_list = ", ".join(HOOK_EVENT_LABELS.keys())
        prompt = (
            "~/.claude/settings.json の hooks に以下の設定を追加してください。\n"
            "既存の hooks エントリは壊さずマージしてください。\n\n"
            f'コマンド: "{hook_venv_python}" "{hook_script_path}" {{イベント名}}\n'
            "async: true\n"
            f"対象イベント: {event_list}\n\n"
            "各イベントごとに matcher は空文字列で登録してください。\n"
            "※ フィルタリングはスクリプト側で行うため全イベント登録してください。"
        )
        root.clipboard_clear()
        root.clipboard_append(prompt)
        hook_setup_status.set("Claude Code 用プロンプトをコピーしました")

    def copy_claude_code_remove_prompt() -> None:
        """Claude Code に投げる削除用プロンプトを生成してクリップボードにコピーする。"""
        prompt = (
            "~/.claude/settings.json の hooks から、command に "
            f'"hook_hotkey.py" を含むエントリだけを削除してください。\n\n'
            "注意:\n"
            "- 同じイベントに紐づく他の hook（ntfy 等）は絶対に消さないでください。\n"
            "- 各イベントの hooks 配列から該当エントリだけを除去してください。\n"
            "- 除去後に hooks 配列が空になったイベントキーは削除してください。\n"
            "- hooks オブジェクト自体が空になった場合は hooks キーごと削除してください。"
        )
        root.clipboard_clear()
        root.clipboard_append(prompt)
        hook_setup_status.set("削除用プロンプトをコピーしました")

    setup_btn_row = ttk.Frame(setup_group)
    setup_btn_row.grid(row=1, column=0, pady=(8, 0), sticky="ew")
    ttk.Button(
        setup_btn_row,
        text="追加プロンプトをコピー",
        command=copy_claude_code_prompt,
    ).pack(side="left")
    ttk.Button(
        setup_btn_row,
        text="削除プロンプトをコピー",
        command=copy_claude_code_remove_prompt,
    ).pack(side="left", padx=(6, 0))
    ttk.Button(setup_btn_row, text="設定 JSON を表示", command=show_hook_snippet).pack(
        side="left",
        padx=(6, 0),
    )

    def show_hook_help() -> None:
        help_win = tk.Toplevel(root)
        help_win.title("Hook の止め方・消し方")
        help_win.geometry("520x520")
        help_win.configure(bg="white")
        help_win.transient(root)
        help_win.grab_set()

        help_md = (
            "## 特定のイベントだけ止めたい場合\n"
            "そのイベントを「0: 送信しない」に変更して保存するだけでOKです。"
            "`settings.json` の編集は不要です。\n"
            "\n"
            "## Hook 連携を全て止めたい場合\n"
            "「Hook 発生時にホットキーを送信する」のチェックを外して保存してください。"
            "`settings.json` にエントリが残っていても一切動作しません。\n"
            "\n"
            "## settings.json から完全に削除したい場合\n"
            "- **方法1:** 「削除プロンプトをコピー」→ Claude Code に貼り付け。"
            "`hook_hotkey.py` のエントリだけが削除され、他の Hook は残ります。\n"
            "- **方法2:** `~/.claude/settings.json` をテキストエディタで開き、"
            "`hooks` 内の `hook_hotkey.py` を含むエントリを手動で削除する\n"
            "\n"
            "※ 通常は削除不要です。止めたいイベントを「0: 送信しない」にするか、"
            "チェックを外すだけで十分です。\n"
            "\n"
            "### コピーされる内容について\n"
            "追加プロンプトには `hook_hotkey.py` の**絶対パス**が含まれます。"
            "このパスはプロジェクトフォルダ内のスクリプトを指しており、"
            "本ツール（Omokage-Character-Agent）が外部の設定ファイル等にアクセスするものではありません。"
        )

        ttk.Button(
            help_win,
            text="閉じる",
            command=help_win.destroy,
        ).pack(side="bottom", pady=(4, 12))
        _render_md(help_win, help_md)

    ttk.Button(setup_btn_row, text="？", width=2, command=show_hook_help).pack(
        side="left", padx=(6, 0)
    )
    ttk.Label(setup_btn_row, textvariable=hook_setup_status, foreground="gray").pack(
        side="left", padx=(8, 0)
    )

    def collect_hook_expression_mapping() -> dict[str, int]:
        result: dict[str, int] = {}
        for event_name, var in hook_expr_vars.items():
            try:
                eid = int(var.get().split(":")[0].strip())
            except (ValueError, IndexError):
                eid = 0
            result[event_name] = eid
        return result

    def collect_hotkey_mapping() -> dict[str, str]:
        return {eid_str: var.get() for eid_str, var in hotkey_vars.items()}

    def collect_expression_voice_params() -> dict[str, dict[str, float]]:
        result: dict[str, dict[str, float]] = {}
        for eid_str, param_vars in expr_voice_vars.items():
            result[eid_str] = {}
            for param_key, var in param_vars.items():
                try:
                    result[eid_str][param_key] = round(float(var.get()), 4)
                except ValueError:
                    result[eid_str][param_key] = 0.0
        return result

    expr_selector_combo.bind("<<ComboboxSelected>>", update_expr_detail_panel)
    detail_hotkey.trace_add("write", sync_detail_to_vars)
    detail_speed.trace_add("write", sync_detail_to_vars)
    detail_pitch.trace_add("write", sync_detail_to_vars)
    detail_intonation.trace_add("write", sync_detail_to_vars)
    detail_volume.trace_add("write", sync_detail_to_vars)
    update_expr_detail_panel()

    def sync_speaker_label_from_id() -> None:
        try:
            speaker_id = int(variables["voicevox_speaker_id"].get())
        except ValueError:
            variables["voicevox_speaker_label"].set("Speaker ID を入力してください")
            return

        label = speaker_ids_to_labels.get(speaker_id)
        if label is None:
            variables["voicevox_speaker_label"].set(f"Speaker ID {speaker_id}")
            return

        variables["voicevox_speaker_label"].set(label)

    def apply_speaker_filter(*_args: object) -> None:
        keyword = variables["voicevox_speaker_search"].get().strip().casefold()
        if not keyword:
            filtered_labels = all_speaker_labels
        else:
            filtered_labels = [
                label for label in all_speaker_labels if keyword in label.casefold()
            ]

        speaker_combobox.configure(values=filtered_labels)
        if all_speaker_labels:
            speaker_status.set(
                f"VOICEVOX speaker一覧 {len(filtered_labels)} 件表示中 / 全 {len(all_speaker_labels)} 件"
            )

    def refresh_voicevox_speaker_options(*, show_error: bool) -> None:
        base_url = variables["voicevox_base_url"].get().strip() or VOICEVOX_BASE_URL

        def _fetch() -> None:
            try:
                options = fetch_voicevox_speaker_options(base_url)
            except Exception as error:
                root.after(0, lambda: _on_fetch_error(error, show_error))
                return
            root.after(0, lambda: _on_fetch_success(options))

        def _on_fetch_error(error: Exception, show_err: bool) -> None:
            nonlocal all_speaker_labels
            speaker_combobox.configure(values=[])
            speaker_ids_to_labels.clear()
            speaker_labels_to_ids.clear()
            all_speaker_labels = []
            sync_speaker_label_from_id()
            speaker_status.set(
                "VOICEVOX speaker一覧を取得できませんでした。Speaker ID の手入力を使ってください。"
            )
            if show_err:
                messagebox.showwarning(
                    "VOICEVOX一覧取得失敗",
                    f"VOICEVOX speaker一覧を取得できませんでした: {error}",
                )

        def _on_fetch_success(options: list[tuple[int, str]]) -> None:
            nonlocal all_speaker_labels
            speaker_ids_to_labels.clear()
            speaker_labels_to_ids.clear()

            labels = []
            for speaker_id, label in options:
                speaker_ids_to_labels[speaker_id] = label
                speaker_labels_to_ids[label] = speaker_id
                labels.append(label)

            all_speaker_labels = labels
            apply_speaker_filter()
            sync_speaker_label_from_id()
            speaker_status.set(f"VOICEVOX speaker一覧を {len(labels)} 件読み込みました")
            voicevox_warn_var.set("")
            voicevox_warn_label.pack_forget()
            # スピーカー名が揃ったのでプリセット詳細を再表示
            if preset_combo.get().strip():
                _on_preset_selected()

        threading.Thread(target=_fetch, daemon=True).start()

    def on_speaker_selected(_event: object | None = None) -> None:
        selected_label = variables["voicevox_speaker_label"].get().strip()
        speaker_id = speaker_labels_to_ids.get(selected_label)
        if speaker_id is None:
            return

        variables["voicevox_speaker_id"].set(str(speaker_id))

        if auto_sample_var.get():
            play_voicevox_sample()

    def play_voicevox_sample() -> None:
        try:
            speaker_id = int(variables["voicevox_speaker_id"].get())
        except ValueError:
            messagebox.showwarning(
                "Speaker ID 不正",
                "サンプル再生の前に有効な Speaker ID を選択してください。",
            )
            return

        base_url = variables["voicevox_base_url"].get().strip() or VOICEVOX_BASE_URL
        output_device_name = (
            normalize_device_selection(
                variables["monitor_device_name"].get(),
                allow_default=True,
            )
            if monitor_enabled.get()
            else ""
        )

        if sd is not None:
            sd.stop()
        _sample_generation[0] += 1
        gen = _sample_generation[0]
        _set_sample_playing(True)
        speaker_status.set("VOICEVOX サンプルを再生しています")

        def on_sample_finished() -> None:
            _set_sample_playing(False)
            apply_speaker_filter()

        def worker() -> None:
            try:
                wav_bytes = synthesize_voicevox_audio(
                    base_url, speaker_id, SPEAKER_SAMPLE_TEXT
                )
                play_sample_audio(
                    wav_bytes,
                    output_device_name,
                    generation=gen,
                    current_gen=_sample_generation,
                )
            except Exception as error:
                if gen == _sample_generation[0]:
                    root.after(
                        0,
                        lambda: messagebox.showwarning(
                            "サンプル再生失敗",
                            f"VOICEVOXサンプルの再生に失敗しました: {error}",
                        ),
                    )
            finally:
                if gen == _sample_generation[0]:
                    root.after(0, on_sample_finished)

        threading.Thread(target=worker, daemon=True).start()

    def refresh_device_options(*, use_cache: bool = True) -> None:
        global _cached_device_list
        if not use_cache:
            _ensure_imports()
            try:
                devices = list(sd.query_devices())
            except Exception:
                devices = []
            with _device_list_lock:
                _cached_device_list = devices
        options = list_output_device_options()
        v_opts = filter_virtual_devices(options)
        p_opts = filter_physical_devices(options)
        vb_combobox.configure(values=v_opts if v_opts else options)
        monitor_combobox.configure(values=p_opts if p_opts else options)

        cur_vb_raw = variables["vbcable_device_name"].get().strip()
        if not cur_vb_raw and len(v_opts) == 1:
            # 初回起動時: 仮想ケーブルが1つだけなら自動選択
            variables["vbcable_device_name"].set(v_opts[0])
        else:
            variables["vbcable_device_name"].set(
                normalize_device_selection(cur_vb_raw)
            )
        # 仮想ケーブル未選択時の警告
        if not variables["vbcable_device_name"].get().strip():
            if len(v_opts) == 0:
                vb_warn_var.set("※ 仮想ケーブル（VB-Cable 等）が見つかりません。インストールしてください。")
            else:
                vb_warn_var.set("※ 仮想ケーブルが複数あります。使用するデバイスを選択して保存してください。")
            vb_warn_label.grid()
        else:
            vb_warn_var.set("")
            vb_warn_label.grid_remove()
        monitor_value = normalize_device_selection(
            variables["monitor_device_name"].get(),
            allow_default=True,
        )
        variables["monitor_device_name"].set(monitor_value or DEFAULT_DEVICE_LABEL)

    def update_summary_prompt_ui(*_args: object) -> None:
        is_enabled = summary_generation_enabled.get()
        prompt_path = variables["summary_system_prompt_path"].get().strip()
        prompt_exists = bool(prompt_path) and Path(prompt_path).is_file()

        summary_prompt_entry.configure(state="normal" if is_enabled else "disabled")
        summary_prompt_button.configure(state="normal" if is_enabled else "disabled")

        if not is_enabled:
            summary_prompt_status.set("デフォルトのスタイルで要約します。")
            summary_preview.set(
                build_summary_preview(enabled=False, prompt_path=prompt_path)
            )
            return

        if not prompt_path:
            summary_prompt_status.set(
                "キャラプロンプトの .txt または .md ファイルを指定してください。"
            )
            summary_preview.set(
                build_summary_preview(enabled=True, prompt_path=prompt_path)
            )
            return

        if prompt_exists:
            summary_prompt_status.set(f"使用予定: {Path(prompt_path).name}")
            summary_preview.set(
                build_summary_preview(enabled=True, prompt_path=prompt_path)
            )
            return

        summary_prompt_status.set(
            "指定ファイルが見つかりません。保存後も通常ルールへフォールバックします。"
        )
        summary_preview.set(
            build_summary_preview(enabled=True, prompt_path=prompt_path)
        )

    speaker_combobox.bind("<<ComboboxSelected>>", on_speaker_selected)
    variables["voicevox_speaker_search"].trace_add(
        "write", lambda *_args: apply_speaker_filter()
    )
    # auto_sample_check は on_speaker_selected 内で play_voicevox_sample を呼ぶ
    refresh_voicevox_button.configure(
        command=lambda: refresh_voicevox_speaker_options(show_error=True)
    )
    variables["voicevox_speaker_id"].trace_add(
        "write", lambda *_args: sync_speaker_label_from_id()
    )
    summary_generation_enabled.trace_add("write", update_summary_prompt_ui)
    variables["summary_system_prompt_path"].trace_add("write", update_summary_prompt_ui)

    refresh_voicevox_speaker_options(show_error=False)
    check_voicevox_on_startup()
    update_summary_prompt_ui()

    # ── 未保存検知用 ────────────────────────────────────────

    _last_saved_snapshot: dict[str, object] = {}

    def _collect_current_settings() -> dict[str, object] | None:
        """GUI 上の現在値を dict にまとめる。バリデーション失敗時は None。"""
        try:
            speaker_id = int(variables["voicevox_speaker_id"].get())
        except ValueError:
            return None
        if speaker_id < 0:
            return None

        try:
            speed_scale = float(variables["voice_speed_scale"].get())
        except ValueError:
            speed_scale = VOICE_SPEED_SCALE
        try:
            pitch_scale = float(variables["voice_pitch_scale"].get())
        except ValueError:
            pitch_scale = VOICE_PITCH_SCALE
        try:
            intonation_scale = float(variables["voice_intonation_scale"].get())
        except ValueError:
            intonation_scale = VOICE_INTONATION_SCALE
        try:
            volume_scale = float(variables["voice_volume_scale"].get())
        except ValueError:
            volume_scale = VOICE_VOLUME_SCALE
        try:
            max_chars = int(variables["summary_max_chars"].get())
            if max_chars < 1:
                max_chars = SUMMARY_MAX_CHARS
        except ValueError:
            max_chars = SUMMARY_MAX_CHARS

        return {
            "avatar_enabled": avatar_enabled.get(),
            "voicevox_speaker_id": speaker_id,
            "voicevox_base_url": variables["voicevox_base_url"].get().strip()
            or VOICEVOX_BASE_URL,
            "vbcable_device_name": normalize_device_selection(
                variables["vbcable_device_name"].get().strip() or VBCABLE_DEVICE_NAME
            ),
            "monitor_playback_enabled": monitor_enabled.get(),
            "monitor_device_name": normalize_device_selection(
                variables["monitor_device_name"].get().strip(),
                allow_default=True,
            ),
            "voice_speed_scale": round(speed_scale, 4),
            "voice_pitch_scale": round(pitch_scale, 4),
            "voice_intonation_scale": round(intonation_scale, 4),
            "voice_volume_scale": round(volume_scale, 4),
            "summary_generation_enabled": summary_generation_enabled.get(),
            "summary_system_prompt_path": variables["summary_system_prompt_path"]
            .get()
            .strip()
            or SUMMARY_SYSTEM_PROMPT_PATH,
            "summary_max_chars": max_chars,
            "avatar_log_enabled": avatar_log_enabled.get(),
            "log_slot_active": log_slot_active.get(),
            "log_slot_names": [v.get() for v in log_slot_name_vars],
            "hotkey_mapping": collect_hotkey_mapping(),
            "expression_voice_params": collect_expression_voice_params(),
            "hook_hotkey_enabled": hook_hotkey_enabled.get(),
            "hook_cooldown_ms": hook_cooldown_ms.get(),
            "hook_expression_mapping": collect_hook_expression_mapping(),
            "vmm_automation_port": _vmm_get_port(),
            "last_loaded_preset": _last_loaded_preset_name[0],
        }

    def _take_snapshot() -> None:
        current = _collect_current_settings()
        if current is not None:
            _last_saved_snapshot.clear()
            _last_saved_snapshot.update(current)

    def _has_unsaved_changes() -> bool:
        current = _collect_current_settings()
        if current is None:
            return True  # バリデーション失敗 = 何か変更されている
        return current != _last_saved_snapshot

    # 起動時点のスナップショットを記録
    _take_snapshot()

    def save() -> None:
        next_settings = _collect_current_settings()
        if next_settings is None:
            messagebox.showerror(
                "入力エラー",
                "入力値に問題があります。Speaker ID は0以上の整数で入力してください。",
            )
            return

        diff_text = _build_diff_text()

        try:
            save_settings(next_settings)
        except OSError as error:
            messagebox.showerror("保存エラー", f"設定保存に失敗しました: {error}")
            return

        _take_snapshot()
        save_btn.configure(text=_SAVE_LABEL_CLEAN)

        if diff_text:
            msg = f"設定を保存しました: {SETTINGS_FILE.name}\n\n変更内容:\n{diff_text}"
        else:
            msg = f"設定を保存しました: {SETTINGS_FILE.name}\n（変更なし）"
        messagebox.showinfo("保存完了", msg)

    # ── 環境チェック ───────────────────────────────────────

    def run_setup_wizard() -> None:
        wizard = tk.Toplevel(root)
        wizard.title("環境チェック")
        wizard.resizable(False, False)
        wizard.grab_set()

        steps_frame = ttk.Frame(wizard, padding=16)
        steps_frame.pack(fill="both", expand=True)

        status_vars: list[tk.StringVar] = []
        check_labels: list[ttk.Label] = []

        steps = [
            "VOICEVOX 接続確認",
            "仮想ケーブル（VB-Cable等）検出",
            "VMagicMirror 起動確認",
            "CLAUDE.md ジェネレーター確認",
        ]
        for i, step_name in enumerate(steps):
            sv = tk.StringVar(master=wizard, value="[ ] " + step_name)
            status_vars.append(sv)
            lbl = ttk.Label(steps_frame, textvariable=sv, font=("", 10))
            lbl.grid(row=i, column=0, sticky="w", pady=4)
            check_labels.append(lbl)

        result_text = tk.Text(
            steps_frame,
            height=8,
            width=56,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
        )
        result_text.grid(row=len(steps), column=0, pady=(12, 0), sticky="nsew")

        def append_result(msg: str) -> None:
            result_text.configure(state="normal")
            result_text.insert("end", msg + "\n")
            result_text.see("end")
            result_text.configure(state="disabled")

        def run_checks() -> None:
            def _update(step: int, ok_str: str, label: str, detail: str) -> None:
                root.after(
                    0,
                    lambda: (
                        status_vars[step].set(f"{ok_str} {label}"),
                        append_result(detail),
                    ),
                )

            # Step 1: VOICEVOX
            base_url = variables["voicevox_base_url"].get().strip() or VOICEVOX_BASE_URL
            ok, msg = check_voicevox_connection(base_url)
            _update(0, "[OK]" if ok else "[NG]", "VOICEVOX 接続確認", msg)

            # Step 2: 仮想ケーブル
            ok2, msg2 = check_virtual_cable_available()
            _update(
                1, "[OK]" if ok2 else "[NG]", "仮想ケーブル（VB-Cable等）検出", msg2
            )

            # Step 3: VMagicMirror
            vmm_ok = False
            try:
                import subprocess as _sp

                result = _sp.run(
                    ["tasklist", "/FI", "IMAGENAME eq VMagicMirror.exe", "/NH"],
                    capture_output=True,
                    timeout=5,
                )
                vmm_ok = b"VMagicMirror.exe" in result.stdout
            except Exception:
                pass
            if vmm_ok:
                _update(2, "[OK]", "VMagicMirror 起動確認", "VMagicMirror が起動中です")
            else:
                _update(
                    2,
                    "[NG]",
                    "VMagicMirror 起動確認",
                    "VMagicMirror が検出されません。起動してください",
                )

            # Step 4: CLAUDE.md
            _update(
                3,
                "[--]",
                "CLAUDE.md ジェネレーター確認",
                "「CLAUDE.md ジェネレーター」タブで\nテキストを生成し、~/.claude/CLAUDE.md に追記してください。",
            )

            root.after(0, lambda: check_btn.configure(state="!disabled"))

        def run_checks_async() -> None:
            check_btn.configure(state="disabled")
            threading.Thread(target=run_checks, daemon=True).start()

        btn_frame = ttk.Frame(steps_frame)
        btn_frame.grid(row=len(steps) + 1, column=0, pady=(12, 0), sticky="e")
        check_btn = ttk.Button(btn_frame, text="チェック開始", command=run_checks_async)
        check_btn.pack(side="left")
        ttk.Button(btn_frame, text="閉じる", command=wizard.destroy).pack(
            side="left", padx=(8, 0)
        )

    # ── ログビューワー ────────────────────────────────────

    _log_viewer_win: list[tk.Toplevel | None] = [None]

    def open_log_viewer() -> None:
        if _log_viewer_win[0] is not None:
            try:
                _log_viewer_win[0].lift()
                _log_viewer_win[0].focus_force()
                return
            except tk.TclError:
                _log_viewer_win[0] = None

        viewer = tk.Toplevel(root)
        viewer.title("履歴ログビューワー")
        viewer.geometry("700x450")
        _log_viewer_win[0] = viewer
        viewer.protocol(
            "WM_DELETE_WINDOW",
            lambda: (_log_viewer_win.__setitem__(0, None), viewer.destroy()),
        )

        # ── スロット切り替えバー ──
        slot_bar = ttk.Frame(viewer)
        slot_bar.pack(fill="x", padx=8, pady=(8, 0))
        ttk.Label(slot_bar, text="スロット:").pack(side="left")
        viewer_slot = tk.IntVar(value=log_slot_active.get())

        # ── Treeview ──
        tree_frame = ttk.Frame(viewer)
        tree_frame.pack(fill="both", expand=True, padx=8, pady=4)

        columns = ("timestamp", "expression", "text")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15)
        tree.heading("timestamp", text="日時")
        tree.heading("expression", text="エモート")
        tree.heading("text", text="要約文")
        tree.column("timestamp", width=140, stretch=False)
        tree.column("expression", width=80, stretch=False)
        tree.column("text", width=400, stretch=True)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)

        # ── 容量警告 ──
        warn_frame = ttk.Frame(viewer)
        warn_frame.pack(fill="x", padx=8, pady=(0, 2))

        # ── 情報バー ──
        info_frame = ttk.Frame(viewer)
        info_frame.pack(fill="x", padx=8, pady=(0, 8))
        info_label = ttk.Label(info_frame, text="", foreground="gray")
        info_label.pack(side="left")

        def _get_slot_log_path(slot_num: int) -> Path:
            idx = max(1, min(LOG_SLOT_COUNT, slot_num)) - 1
            return Path(__file__).with_name(LOG_SLOT_FILES[idx])

        def _load_slot(slot_num: int) -> None:
            log_path = _get_slot_log_path(slot_num)
            # ツリーをクリア
            for item in tree.get_children():
                tree.delete(item)
            # ログ読み込み
            if not log_path.is_file():
                tree.insert("", "end", values=("", "", "ログファイルが存在しません"))
            else:
                try:
                    lines = log_path.read_text(encoding="utf-8").strip().splitlines()
                    for line in reversed(lines[-500:]):
                        try:
                            entry = json.loads(line)
                            eid = int(entry.get("expression_id", 0))
                            expr_name = EXPRESSION_ID_LABELS.get(eid, "?").split(" / ")[
                                0
                            ]
                            tree.insert(
                                "",
                                "end",
                                values=(
                                    entry.get("timestamp", ""),
                                    f"{eid}: {expr_name}",
                                    entry.get("text", ""),
                                ),
                            )
                        except (json.JSONDecodeError, ValueError):
                            continue
                except OSError:
                    tree.insert(
                        "", "end", values=("", "", "ログの読み込みに失敗しました")
                    )
            # 容量警告更新
            for w in warn_frame.winfo_children():
                w.destroy()
            log_size_bytes = log_path.stat().st_size if log_path.is_file() else 0
            log_size_mb = log_size_bytes / (1024 * 1024)
            if log_size_mb >= 4.0:
                ttk.Label(
                    warn_frame,
                    text=f"\u26a0 ログ容量が {log_size_mb:.1f} MB に達しています（上限 5 MB）。"
                    "エクスポートしてからクリアすることを推奨します。",
                    foreground="red",
                ).pack(anchor="w")
            # 情報ラベル更新
            entry_count = len(tree.get_children())
            info_label.configure(
                text=f"表示件数: {entry_count} 件（最新500件まで）  |  ログ容量: {log_size_mb:.2f} MB / 5 MB"
            )

        def _on_slot_change(*_args: object) -> None:
            _load_slot(viewer_slot.get())

        # スロットラジオボタン生成
        for i in range(LOG_SLOT_COUNT):
            ttk.Radiobutton(
                slot_bar,
                textvariable=log_slot_name_vars[i],
                variable=viewer_slot,
                value=i + 1,
                command=lambda: _on_slot_change(),
            ).pack(side="left", padx=(8, 0))

        # 初期読み込み
        _load_slot(viewer_slot.get())

        def export_log() -> None:
            log_path = _get_slot_log_path(viewer_slot.get())
            if not log_path.is_file() or log_path.stat().st_size == 0:
                messagebox.showinfo(
                    "エクスポート", "エクスポートするログがありません。", parent=viewer
                )
                return
            slot_idx = viewer_slot.get() - 1
            slot_label = log_slot_name_vars[slot_idx].get()
            dest = filedialog.asksaveasfilename(
                title="履歴ログをエクスポート",
                defaultextension=".jsonl",
                filetypes=[("JSONL", "*.jsonl"), ("すべて", "*.*")],
                initialfile=f"avatar_log_{slot_label}.jsonl",
                parent=viewer,
            )
            if not dest:
                return
            try:
                import shutil

                shutil.copy2(log_path, dest)
                messagebox.showinfo(
                    "エクスポート完了",
                    f"ログを保存しました:\n{dest}",
                    parent=viewer,
                )
            except OSError as error:
                messagebox.showerror(
                    "エラー", f"エクスポートに失敗しました: {error}", parent=viewer
                )

        def clear_log() -> None:
            log_path = _get_slot_log_path(viewer_slot.get())
            slot_idx = viewer_slot.get() - 1
            slot_label = log_slot_name_vars[slot_idx].get()
            if not messagebox.askyesno(
                "ログクリア",
                f"「{slot_label}」のログを全て削除しますか？\nこの操作は元に戻せません。",
                parent=viewer,
            ):
                return
            try:
                if log_path.is_file():
                    log_path.write_text("", encoding="utf-8")
                _load_slot(viewer_slot.get())
            except OSError as error:
                messagebox.showerror(
                    "エラー", f"ログの削除に失敗しました: {error}", parent=viewer
                )

        btn_frame = ttk.Frame(viewer)
        btn_frame.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btn_frame, text="エクスポート", command=export_log).pack(
            side="left", padx=(0, 4)
        )
        ttk.Button(btn_frame, text="ログをクリア", command=clear_log).pack(side="left")
        ttk.Button(btn_frame, text="閉じる", command=viewer.destroy).pack(side="right")

    # ── 表情プリセット（表情・演出タブに追加） ────────────

    def export_preset() -> None:
        path = filedialog.asksaveasfilename(
            title="プリセットをエクスポート",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialdir=str(Path(__file__).parent),
        )
        if not path:
            return
        preset = {
            "hotkey_mapping": collect_hotkey_mapping(),
            "expression_voice_params": collect_expression_voice_params(),
        }
        try:
            Path(path).write_text(
                json.dumps(preset, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            vmm_test_status.set(f"エクスポートしました: {Path(path).name}")
        except OSError as error:
            messagebox.showerror(
                "エクスポートエラー", f"エクスポートに失敗しました: {error}"
            )

    def import_preset() -> None:
        path = filedialog.askopenfilename(
            title="プリセットをインポート",
            filetypes=[("JSON", "*.json")],
            initialdir=str(Path(__file__).parent),
        )
        if not path:
            return
        try:
            preset = import_expression_preset(Path(path))
        except (json.JSONDecodeError, ValueError, OSError) as error:
            messagebox.showerror(
                "インポートエラー", f"インポートに失敗しました: {error}"
            )
            return

        if "hotkey_mapping" in preset:
            for eid_str, hk in preset["hotkey_mapping"].items():
                if eid_str in hotkey_vars:
                    hotkey_vars[eid_str].set(str(hk))
        if "expression_voice_params" in preset:
            for eid_str, params in preset["expression_voice_params"].items():
                if eid_str in expr_voice_vars and isinstance(params, dict):
                    for param_key, val in params.items():
                        if param_key in expr_voice_vars[eid_str]:
                            expr_voice_vars[eid_str][param_key].set(str(val))
        update_hotkey_labels()
        update_expr_detail_panel()
        vmm_test_status.set(f"インポートしました: {Path(path).name}")

    preset_row = ttk.Frame(expr_frame)
    preset_row.pack(fill="x", pady=(0, 4))
    ttk.Button(preset_row, text="プリセットをエクスポート", command=export_preset).pack(
        side="left"
    )
    ttk.Button(preset_row, text="プリセットをインポート", command=import_preset).pack(
        side="left", padx=(8, 0)
    )

    # ── 下部ボタン行 ──────────────────────────────────────

    bottom_frame = ttk.Frame(root)
    bottom_frame.grid(row=2, column=0, sticky="ew", padx=24, pady=(12, 12))
    bottom_frame.columnconfigure(1, weight=1)

    left_group = ttk.Frame(bottom_frame)
    left_group.grid(row=0, column=0, sticky="w")

    center_group = ttk.Frame(bottom_frame)
    center_group.grid(row=0, column=1)

    right_group = ttk.Frame(bottom_frame)
    right_group.grid(row=0, column=2, sticky="e")

    # ── Markdown レンダリング共通関数 ──

    def _render_md(parent: tk.Widget, md_text: str) -> None:
        """簡易 Markdown レンダリングで Text ウィジェットに描画する。"""
        import webbrowser

        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)

        _link_counter = [0]

        tw = tk.Text(
            frame,
            wrap="word",
            font=("Yu Gothic UI", 10),
            padx=14,
            pady=10,
            state="normal",
            cursor="arrow",
            spacing1=2,
            spacing3=2,
        )
        sb = ttk.Scrollbar(frame, orient="vertical", command=tw.yview)
        tw.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tw.pack(side="left", fill="both", expand=True)

        # タグ定義
        tw.tag_configure(
            "h2", font=("Yu Gothic UI", 14, "bold"), spacing1=10, spacing3=6
        )
        tw.tag_configure(
            "h3", font=("Yu Gothic UI", 12, "bold"), spacing1=8, spacing3=4
        )
        tw.tag_configure("bold", font=("Yu Gothic UI", 10, "bold"))
        tw.tag_configure(
            "code",
            font=("Consolas", 9),
            background="#f0f0f0",
            relief="flat",
            borderwidth=1,
            lmargin1=20,
            lmargin2=20,
        )
        tw.tag_configure("bullet", lmargin1=16, lmargin2=28)
        tw.tag_configure("table_row", font=("Consolas", 9), lmargin1=8, lmargin2=8)
        tw.tag_configure("sep", foreground="#cccccc")
        tw.tag_configure("dim", foreground="#999999", font=("Yu Gothic UI", 9))
        tw.tag_configure(
            "blockquote",
            foreground="#555555",
            lmargin1=16,
            lmargin2=16,
            background="#f8f8f8",
            font=("Yu Gothic UI", 9),
        )
        tw.tag_configure(
            "bq_bold",
            foreground="#555555",
            font=("Yu Gothic UI", 9, "bold"),
        )
        tw.tag_configure(
            "bq_dim",
            foreground="#aaaaaa",
            font=("Yu Gothic UI", 9),
        )

        def _make_link_tag(url: str) -> str:
            """クリック可能なリンク用のタグを生成する。"""
            tag_name = f"link_{_link_counter[0]}"
            _link_counter[0] += 1
            tw.tag_configure(
                tag_name,
                foreground="#0066cc",
                underline=True,
                font=("Yu Gothic UI", 10),
            )
            tw.tag_bind(tag_name, "<Enter>", lambda _: tw.configure(cursor="hand2"))
            tw.tag_bind(tag_name, "<Leave>", lambda _: tw.configure(cursor="arrow"))
            tw.tag_bind(tag_name, "<Button-1>", lambda _: webbrowser.open(url))
            return tag_name

        def _insert_with_links(text: str, base_tags: tuple = ()) -> None:
            """テキスト中の [label](url) をクリック可能リンクとして挿入する。"""
            parts = re.split(r"\[([^\]]+)\]\(([^)]+)\)", text)
            for i, part in enumerate(parts):
                if i % 3 == 0:
                    if part:
                        tw.insert("end", part, base_tags)
                elif i % 3 == 1:
                    label = part
                    url = parts[i + 1]
                    link_tag = _make_link_tag(url)
                    tw.insert("end", label, base_tags + (link_tag,))
                # i % 3 == 2 は url（labelと一緒に処理済み）

        def _strip_md_links(text: str) -> str:
            """[label](url) → label に変換（Treeview用）。"""
            return re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)

        def _parse_table_cells(line: str) -> list[str]:
            cells = line.split("|")
            if cells and not cells[0].strip():
                cells = cells[1:]
            if cells and not cells[-1].strip():
                cells = cells[:-1]
            return [_strip_md_links(c.strip()) for c in cells]

        def _is_separator(line: str) -> bool:
            return bool(re.match(r"^\|?\s*[:=-]+[\s|:=-]+$", line.replace("-", "=")))

        def _extract_row_urls(line: str) -> list[tuple[str, str]]:
            """テーブル行から [label](url) ペアを全て抽出する。"""
            return re.findall(r"\[([^\]]+)\]\(([^)]+)\)", line)

        def _flush_table(table_lines: list[str]) -> None:
            rows: list[list[str]] = []
            raw_data_lines: list[str] = []
            is_header = True
            for tl in table_lines:
                if _is_separator(tl):
                    is_header = False
                    continue
                rows.append(_parse_table_cells(tl))
                if not is_header:
                    raw_data_lines.append(tl)
            if not rows:
                return

            headers = rows[0]
            data_rows = rows[1:]
            col_count = max(len(r) for r in rows)

            # 各データ行からURLを抽出
            row_urls: list[list[tuple[str, str]]] = []
            has_any_url = False
            for raw_line in raw_data_lines:
                urls = _extract_row_urls(raw_line)
                row_urls.append(urls)
                if urls:
                    has_any_url = True

            # 外枠フレーム（テーブル + リンク列を横並び）
            outer_frame = tk.Frame(tw, bg="white")

            tree_frame = tk.Frame(outer_frame, bg="white")
            col_ids = [f"c{i}" for i in range(col_count)]
            tree = ttk.Treeview(
                tree_frame,
                columns=col_ids,
                show="headings",
                height=min(len(data_rows), 15),
            )

            def _pixel_width(s: str) -> int:
                w = 0
                for ch in s:
                    w += 14 if ord(ch) > 0x7F else 8
                return w + 16

            for ci, cid in enumerate(col_ids):
                heading = headers[ci] if ci < len(headers) else ""
                tree.heading(cid, text=heading, anchor="w")
                max_px = _pixel_width(heading)
                for r in data_rows:
                    cell = r[ci] if ci < len(r) else ""
                    max_px = max(max_px, _pixel_width(cell))
                tree.column(cid, width=max(max_px, 60), anchor="w")

            for r in data_rows:
                values = [r[ci] if ci < len(r) else "" for ci in range(col_count)]
                tree.insert("", "end", values=values)

            tree.pack(fill="both", expand=True, padx=4)
            tree_frame.pack(side="left", fill="both", expand=True)

            # 右横にリンク列を配置
            if has_any_url:
                link_frame = tk.Frame(outer_frame, bg="white", padx=6)
                # ヘッダー行の高さ分スペーサー
                tk.Label(
                    link_frame, text="", bg="white",
                    font=("Yu Gothic UI", 2),
                ).pack(anchor="w")
                for urls in row_urls:
                    if urls:
                        row_link_frame = tk.Frame(link_frame, bg="white")
                        for _, (_, url) in enumerate(urls):
                            lbl = tk.Label(
                                row_link_frame,
                                text=url,
                                fg="#0066cc",
                                bg="white",
                                font=("Yu Gothic UI", 9, "underline"),
                                cursor="hand2",
                            )
                            lbl.pack(anchor="w")
                            lbl.bind("<Button-1>", lambda _, u=url: webbrowser.open(u))
                        row_link_frame.pack(anchor="w", pady=1)
                    else:
                        tk.Label(
                            link_frame, text="", bg="white",
                            font=("Yu Gothic UI", 9),
                        ).pack(anchor="w", pady=1)
                link_frame.pack(side="left", fill="y", anchor="n")

            def _forward_scroll(event: tk.Event, target: tk.Text = tw) -> None:
                target.yview_scroll(-1 * (event.delta // 120), "units")

            tree.bind("<MouseWheel>", _forward_scroll)
            outer_frame.bind("<MouseWheel>", _forward_scroll)

            tw.window_create("end", window=outer_frame)
            tw.insert("end", "\n")

        in_code_block = False
        pending_table: list[str] = []

        for line in md_text.splitlines():
            if pending_table and not line.startswith("|"):
                _flush_table(pending_table)
                pending_table = []

            if line.strip().startswith("```"):
                in_code_block = not in_code_block
                continue
            if in_code_block:
                tw.insert("end", line + "\n", "code")
                continue

            if line.startswith("|"):
                pending_table.append(line)
                continue

            if line.startswith("> ") or line == ">":
                inner = line[2:] if line.startswith("> ") else ""
                if not inner.strip():
                    tw.insert("end", "\n", "blockquote")
                else:
                    segments = re.split(
                        r"(\*\*.+?\*\*|(?<!\*)\*(?!\*).+?(?<!\*)\*(?!\*)|`.+?`)",
                        inner,
                    )
                    for seg in segments:
                        if seg.startswith("**") and seg.endswith("**"):
                            _insert_with_links(seg[2:-2], ("blockquote", "bq_bold"))
                        elif seg.startswith("`") and seg.endswith("`"):
                            tw.insert("end", seg[1:-1], ("blockquote", "code"))
                        elif re.match(r"^\*(?!\*).+(?<!\*)\*$", seg):
                            _insert_with_links(seg[1:-1], ("blockquote", "bq_dim"))
                        else:
                            _insert_with_links(seg, ("blockquote",))
                    tw.insert("end", "\n", "blockquote")
            elif line.startswith("### "):
                tw.insert("end", line[4:] + "\n", "h3")
            elif line.startswith("## "):
                tw.insert("end", line[3:] + "\n", "h2")
            elif re.match(r"^\s*[-*]\s", line):
                text = re.sub(r"^\s*[-*]\s", "• ", line)
                parts = re.split(r"\*\*(.+?)\*\*", text)
                for j, part in enumerate(parts):
                    tag = ("bold", "bullet") if j % 2 == 1 else ("bullet",)
                    _insert_with_links(part, tag)
                tw.insert("end", "\n")
            elif re.match(r"^\s*\d+\.\s", line):
                parts = re.split(r"\*\*(.+?)\*\*", line)
                for j, part in enumerate(parts):
                    tag = ("bold", "bullet") if j % 2 == 1 else ("bullet",)
                    _insert_with_links(part, tag)
                tw.insert("end", "\n")
            else:
                segments = re.split(r"(\*\*.+?\*\*|(?<!\*)\*(?!\*).+?(?<!\*)\*(?!\*)|`.+?`)", line)
                for seg in segments:
                    if seg.startswith("**") and seg.endswith("**"):
                        _insert_with_links(seg[2:-2], ("bold",))
                    elif seg.startswith("`") and seg.endswith("`"):
                        tw.insert("end", seg[1:-1], "code")
                    elif re.match(r"^\*(?!\*).+(?<!\*)\*$", seg):
                        _insert_with_links(seg[1:-1], ("dim",))
                    else:
                        _insert_with_links(seg)
                tw.insert("end", "\n")

        if pending_table:
            _flush_table(pending_table)

        tw.configure(state="disabled")

    # ── 左: ヘルプ ──

    _manual_win: list[tk.Toplevel | None] = [None]

    def show_manual() -> None:
        if _manual_win[0] is not None:
            try:
                _manual_win[0].lift()
                _manual_win[0].focus_force()
                return
            except tk.TclError:
                _manual_win[0] = None

        manual_path = Path(__file__).with_name("MANUAL.md")
        try:
            content = manual_path.read_text(encoding="utf-8")
        except OSError:
            messagebox.showerror(
                "エラー", f"MANUAL.md が見つかりません:\n{manual_path}"
            )
            return

        # --- Markdown を ## で分割 ---
        sections: list[tuple[str, str]] = []
        preamble_lines: list[str] = []
        current_title = ""
        current_lines: list[str] = []
        for line in content.splitlines():
            if line.startswith("## "):
                if current_title:
                    sections.append((current_title, "\n".join(current_lines).strip()))
                elif current_lines:
                    preamble_lines = current_lines[:]
                current_title = line[3:].strip()
                current_lines = []
            else:
                current_lines.append(line)
        if current_lines and current_title:
            sections.append((current_title, "\n".join(current_lines).strip()))

        # プレヘッダー（# タイトル等）を最初のセクションに統合
        if preamble_lines and sections:
            first_title, first_body = sections[0]
            merged = "\n".join(preamble_lines).strip() + "\n\n" + first_body
            sections[0] = (first_title, merged.strip())

        if not sections:
            sections = [("全体", content)]

        # --- ウィンドウ ---
        win = tk.Toplevel(root)
        win.title("ヘルプ — MANUAL.md")
        win.geometry("900x560")
        win.transient(root)
        _manual_win[0] = win
        win.protocol(
            "WM_DELETE_WINDOW",
            lambda: (_manual_win.__setitem__(0, None), win.destroy()),
        )

        help_notebook = ttk.Notebook(win)
        help_notebook.pack(fill="both", expand=True, padx=4, pady=4)

        for title, body in sections:
            tab_frame = ttk.Frame(help_notebook)
            help_notebook.add(tab_frame, text=title)
            _render_md(tab_frame, body)

    ttk.Button(left_group, text="ヘルプ", command=show_manual).pack(side="left")

    # ── 中央: 環境チェック / 保存 / 履歴ログ ──

    ttk.Button(center_group, text="環境チェック", command=run_setup_wizard).pack(
        side="left"
    )
    save_btn = tk.Button(
        center_group,
        text="保存",
        command=save,
        bg="#4a90d9",
        fg="white",
        activebackground="#3a7bc8",
        activeforeground="white",
        relief="raised",
        padx=12,
        pady=2,
        cursor="hand2",
    )
    save_btn.pack(side="left", padx=(24, 0))
    ttk.Button(center_group, text="履歴ログ", command=open_log_viewer).pack(
        side="left", padx=(24, 0)
    )

    # 未保存時に保存ボタンのテキストを変更（変数変更時のみ更新）
    _SAVE_LABEL_CLEAN = "保存"
    _SAVE_LABEL_DIRTY = "保存 *"
    _init_done = False  # 初期化完了後に True にする

    def _mark_dirty(*_args: object) -> None:
        if not _init_done:
            return
        try:
            if _has_unsaved_changes():
                if save_btn.cget("text") != _SAVE_LABEL_DIRTY:
                    save_btn.configure(text=_SAVE_LABEL_DIRTY)
            else:
                if save_btn.cget("text") != _SAVE_LABEL_CLEAN:
                    save_btn.configure(text=_SAVE_LABEL_CLEAN)
        except Exception:
            pass

    # 全 StringVar / BooleanVar の変更を監視
    for _var in variables.values():
        _var.trace_add("write", _mark_dirty)
    for _bvar in (
        avatar_enabled,
        monitor_enabled,
        summary_generation_enabled,
        avatar_log_enabled,
        hook_hotkey_enabled,
        hook_cooldown_ms,
        listen_mode,
    ):
        _bvar.trace_add("write", _mark_dirty)
    for _hvar in hotkey_vars.values():
        _hvar.trace_add("write", _mark_dirty)
    for _ev_dict in expr_voice_vars.values():
        for _ev_var in _ev_dict.values():
            _ev_var.trace_add("write", _mark_dirty)
    for _he_var in hook_expr_vars.values():
        _he_var.trace_add("write", _mark_dirty)

    # ── 右: 終了 ──

    _SETTING_LABELS: dict[str, str] = {
        "avatar_enabled": "アバター有効",
        "voicevox_speaker_id": "VOICEVOX Speaker ID",
        "voicevox_base_url": "VOICEVOX URL",
        "vbcable_device_name": "VB-Cable デバイス",
        "monitor_playback_enabled": "スピーカー再生",
        "monitor_device_name": "スピーカーデバイス",
        "voice_speed_scale": "話速",
        "voice_pitch_scale": "ピッチ",
        "voice_intonation_scale": "抑揚",
        "voice_volume_scale": "音量",
        "summary_generation_enabled": "要約ルール",
        "summary_system_prompt_path": "要約プロンプトファイル",
        "summary_max_chars": "要約文字数上限",
        "avatar_log_enabled": "履歴ログ",
        "hotkey_mapping": "ホットキー割当",
        "expression_voice_params": "表情別声質",
        "hook_hotkey_enabled": "Hook 連携",
        "hook_cooldown_ms": "Hook 待ち時間(ms)",
        "hook_expression_mapping": "Hook 表情マッピング",
    }

    def _format_scalar(value: object) -> str:
        if isinstance(value, bool):
            return "ON" if value else "OFF"
        if value is None:
            return "(なし)"
        return str(value)

    def _diff_dict(
        old: dict[str, object],
        new: dict[str, object],
    ) -> list[str]:
        """dict 内で実際に変わったキーだけの差分行リストを返す。"""
        all_keys = sorted(set(old) | set(new))
        lines: list[str] = []
        for k in all_keys:
            ov, nv = old.get(k), new.get(k)
            if ov != nv:
                lines.append(f"    {k}: {_format_scalar(ov)} → {_format_scalar(nv)}")
        return lines

    def _build_diff_text() -> str:
        current = _collect_current_settings()
        if current is None:
            return ""
        lines: list[str] = []
        for key, new_val in current.items():
            old_val = _last_saved_snapshot.get(key)
            if new_val != old_val:
                label = _SETTING_LABELS.get(key, key)
                if isinstance(new_val, dict) and isinstance(old_val, dict):
                    sub_lines = _diff_dict(old_val, new_val)
                    if sub_lines:
                        lines.append(f"  {label}:")
                        lines.extend(sub_lines)
                else:
                    lines.append(
                        f"  {label}: {_format_scalar(old_val)} → {_format_scalar(new_val)}"
                    )
        return "\n".join(lines)

    def quit_app() -> None:
        if _has_unsaved_changes():
            diff_text = _build_diff_text()

            dlg = tk.Toplevel(root)
            dlg.title("未保存の変更")
            dlg.transient(root)
            dlg.grab_set()
            dlg.resizable(False, False)

            ttk.Label(
                dlg,
                text="保存されていない変更があります:",
                padding=(20, 16, 20, 4),
            ).pack(anchor="w")

            diff_widget = tk.Text(
                dlg,
                wrap="word",
                font=("Consolas", 9),
                relief="flat",
                bg=dlg.cget("bg"),
                cursor="arrow",
                padx=16,
                pady=4,
                height=min(max(diff_text.count("\n") + 1, 2), 12),
                width=52,
            )
            diff_widget.pack(fill="x", padx=(8, 8), pady=(0, 4))
            diff_widget.insert("1.0", diff_text)
            diff_widget.configure(state="disabled")

            btn_frame = ttk.Frame(dlg, padding=(20, 8, 20, 16))
            btn_frame.pack()

            def _save_and_quit() -> None:
                dlg.destroy()
                save()
                if _has_unsaved_changes():
                    return
                root.destroy()

            def _discard() -> None:
                dlg.destroy()
                root.destroy()

            tk.Button(
                btn_frame,
                text="保存して終了",
                command=_save_and_quit,
                bg="#4a90d9",
                fg="white",
                activebackground="#3a7bc8",
                activeforeground="white",
                relief="raised",
                padx=10,
                pady=2,
                cursor="hand2",
            ).pack(side="left", padx=(0, 8))
            ttk.Button(
                btn_frame,
                text="保存せず終了",
                command=_discard,
            ).pack(side="left", padx=(0, 8))
            ttk.Button(
                btn_frame,
                text="キャンセル",
                command=dlg.destroy,
            ).pack(side="left")

            dlg.update_idletasks()
            x = root.winfo_x() + (root.winfo_width() - dlg.winfo_width()) // 2
            y = root.winfo_y() + (root.winfo_height() - dlg.winfo_height()) // 2
            dlg.geometry(f"+{x}+{y}")
            return

        root.destroy()

    ttk.Button(right_group, text="終了", command=quit_app).pack(side="right")
    root.protocol("WM_DELETE_WINDOW", quit_app)

    # GUI表示後にバックグラウンドでデバイス一覧を取得して反映
    def _deferred_device_load() -> None:
        nonlocal _init_done

        def _finish_init() -> None:
            nonlocal _init_done
            refresh_device_options()
            # デバイス一覧が揃った段階でデバイス名を正規化
            cur_vb = variables["vbcable_device_name"].get()
            normalized_vb = normalize_device_selection(cur_vb)
            if normalized_vb != cur_vb:
                variables["vbcable_device_name"].set(normalized_vb)
            cur_mon = variables["monitor_device_name"].get()
            if cur_mon != DEFAULT_DEVICE_LABEL:
                normalized_mon = normalize_device_selection(cur_mon, allow_default=True) or DEFAULT_DEVICE_LABEL
                if normalized_mon != cur_mon:
                    variables["monitor_device_name"].set(normalized_mon)
            # 正規化後の状態をスナップショットとして記録してからダーティフラグ監視を有効化
            _take_snapshot()
            save_btn.configure(text=_SAVE_LABEL_CLEAN)
            _init_done = True

        def _worker() -> None:
            _devices_ready.wait(timeout=DEVICE_CACHE_WAIT_TIMEOUT)
            root.after(0, _finish_init)

        threading.Thread(target=_worker, daemon=True).start()

    root.after(0, _deferred_device_load)

    # 設定ファイル読み込み時の警告があればGUI表示後に通知
    if _settings_load_warning:
        root.after(
            300,
            lambda: messagebox.showwarning("設定ファイルの読み込み", _settings_load_warning),
        )

    root.minsize(580, 500)
    root.mainloop()


def main() -> int:
    args = parse_args()

    if args.print_summary_settings_json:
        import sys, io

        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        print(json.dumps(get_avatar_settings(load_settings()), ensure_ascii=False))
        return 0

    try:
        open_settings_gui()
    except Exception as error:
        import traceback

        traceback.print_exc()
        try:
            messagebox.showerror(
                "Startup Error", f"設定画面の起動に失敗しました: {error}"
            )
        except tk.TclError:
            pass
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
