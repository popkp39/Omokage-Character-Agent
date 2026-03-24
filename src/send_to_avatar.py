from __future__ import annotations

import argparse
import contextlib
import ctypes
import datetime
import io
import json
import platform
import os
import sys
import tempfile
import threading
import time
import wave
from pathlib import Path

import numpy as np
import requests
import sounddevice as sd

import config

IS_WINDOWS = platform.system() == "Windows"

# ── Win32 keybd_event 定義 ─────────────────────────────────
KEYEVENTF_KEYUP = 0x0002

if IS_WINDOWS:
    user32 = ctypes.windll.user32
else:
    user32 = None

VK_MODIFIER_MAP: dict[str, int] = {
    "ctrl": 0x11,
    "shift": 0x10,
    "alt": 0x12,
}

VK_KEY_MAP: dict[str, int] = {
    "0": 0x30, "1": 0x31, "2": 0x32, "3": 0x33, "4": 0x34,
    "5": 0x35, "6": 0x36, "7": 0x37, "8": 0x38, "9": 0x39,
    "a": 0x41, "b": 0x42, "c": 0x43, "d": 0x44, "e": 0x45,
    "f": 0x46, "g": 0x47, "h": 0x48, "i": 0x49, "j": 0x4A,
    "k": 0x4B, "l": 0x4C, "m": 0x4D, "n": 0x4E, "o": 0x4F,
    "p": 0x50, "q": 0x51, "r": 0x52, "s": 0x53, "t": 0x54,
    "u": 0x55, "v": 0x56, "w": 0x57, "x": 0x58, "y": 0x59, "z": 0x5A,
    "f1": 0x70, "f2": 0x71, "f3": 0x72, "f4": 0x73,
    "f5": 0x74, "f6": 0x75, "f7": 0x76, "f8": 0x77,
    "f9": 0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
}


def parse_hotkey_string(hotkey_str: str) -> tuple[list[int], int | None]:
    """'ctrl+shift+1' を ([modifier_vks], key_vk) にパースする。"""
    parts = [p.strip().lower() for p in hotkey_str.split("+")]
    modifiers: list[int] = []
    key_vk: int | None = None

    for part in parts:
        if part in VK_MODIFIER_MAP:
            modifiers.append(VK_MODIFIER_MAP[part])
        elif part in VK_KEY_MAP:
            key_vk = VK_KEY_MAP[part]

    return modifiers, key_vk


def send_hotkey(expression_id: int, settings: dict[str, object] | None = None) -> None:
    """設定に基づいたホットキーを送信して VMagicMirror の Word to Motion を発火する。"""
    if not IS_WINDOWS or user32 is None:
        print(
            "ホットキー送信は Windows でのみ動作します。",
            file=sys.stderr,
        )
        return

    if not config.EXPRESSION_ID_MIN <= expression_id <= config.EXPRESSION_ID_MAX:
        return

    hotkey_str = ""
    if settings:
        mapping = settings.get("hotkey_mapping")
        if isinstance(mapping, dict):
            hotkey_str = str(mapping.get(str(expression_id), ""))

    if not hotkey_str.strip():
        hotkey_str = f"ctrl+shift+{expression_id % 10}"

    modifiers, key_vk = parse_hotkey_string(hotkey_str)
    if key_vk is None:
        print(
            f"無効なホットキー設定です（expression_id={expression_id}, "
            f"hotkey='{hotkey_str}'）",
            file=sys.stderr,
        )
        return

    for mod in modifiers:
        user32.keybd_event(mod, 0, 0, 0)
    user32.keybd_event(key_vk, 0, 0, 0)
    time.sleep(0.05)
    user32.keybd_event(key_vk, 0, KEYEVENTF_KEYUP, 0)
    for mod in reversed(modifiers):
        user32.keybd_event(mod, 0, KEYEVENTF_KEYUP, 0)


# ── 本命送信タイムスタンプ（Hook 抑制用） ────────────────────
_AVATAR_SENT_FILE = Path(__file__).with_name(".avatar_sent")


def _mark_avatar_sent() -> None:
    """本命の表情送信時刻を記録する。hook_hotkey.py がこれを参照して遠慮する。"""
    try:
        _AVATAR_SENT_FILE.write_text(str(time.time()), encoding="utf-8")
    except OSError:
        pass


def get_avatar_sent_time() -> float:
    """最後に send_to_avatar.py が表情を送信した時刻を返す。未記録時は 0。"""
    try:
        return float(_AVATAR_SENT_FILE.read_text(encoding="utf-8").strip())
    except (OSError, ValueError):
        return 0.0


# ── 表情送信 ──────────────────────────────────────────────

def send_expression(expression_id: int, settings: dict[str, object]) -> None:
    """ホットキーで VMagicMirror に表情を送信する。"""
    try:
        send_hotkey(expression_id, settings)
        _mark_avatar_sent()
    except Exception as error:
        print(
            f"VMagicMirror へのホットキー送信をスキップしました: {error}",
            file=sys.stderr,
        )


# ── 音声合成・再生 ────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="VOICEVOXで音声を再生しつつ、VMagicMirrorへホットキーで表情を送信します。",
    )
    parser.add_argument("text", help="読み上げるテキスト")
    parser.add_argument("expression_id", type=int, help="送信する表情ID")
    return parser.parse_args()


def _safe_float(value: object, default: float) -> float:
    """値を float に変換する。失敗時は default を返す。"""
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def build_voice_params(
    settings: dict[str, object], expression_id: int
) -> dict[str, float]:
    """グローバル声質 + 表情別オフセットをマージして返す。"""
    global_speed = _safe_float(settings.get("voice_speed_scale"), config.VOICE_SPEED_SCALE)
    global_pitch = _safe_float(settings.get("voice_pitch_scale"), config.VOICE_PITCH_SCALE)
    global_intonation = _safe_float(settings.get("voice_intonation_scale"), config.VOICE_INTONATION_SCALE)
    global_volume = _safe_float(settings.get("voice_volume_scale"), config.VOICE_VOLUME_SCALE)

    expr_params = {}
    all_expr_params = settings.get("expression_voice_params")
    if isinstance(all_expr_params, dict):
        entry = all_expr_params.get(str(expression_id))
        if isinstance(entry, dict):
            expr_params = entry

    return {
        "speedScale": global_speed + _safe_float(expr_params.get("speed_offset"), 0.0),
        "pitchScale": global_pitch + _safe_float(expr_params.get("pitch_offset"), 0.0),
        "intonationScale": global_intonation
        + _safe_float(expr_params.get("intonation_offset"), 0.0),
        "volumeScale": global_volume + _safe_float(expr_params.get("volume_offset"), 0.0),
    }


def synthesize_voice(
    text: str,
    settings: dict[str, object],
    expression_id: int = 0,
) -> bytes:
    base_url = config._validate_voicevox_url(
        str(settings.get("voicevox_base_url", config.VOICEVOX_BASE_URL))
    )
    try:
        speaker_id = int(settings.get("voicevox_speaker_id", 1))
    except (TypeError, ValueError):
        speaker_id = 1

    query_response = requests.post(
        f"{base_url}/audio_query",
        params={"text": text, "speaker": speaker_id},
        timeout=config.VOICEVOX_AUDIO_QUERY_TIMEOUT,
    )
    query_response.raise_for_status()
    query_data = query_response.json()

    if expression_id is not None:
        voice_params = build_voice_params(settings, expression_id)
        query_data["speedScale"] = voice_params["speedScale"]
        query_data["pitchScale"] = voice_params["pitchScale"]
        query_data["intonationScale"] = voice_params["intonationScale"]
        query_data["volumeScale"] = voice_params["volumeScale"]

    synthesis_response = requests.post(
        f"{base_url}/synthesis",
        params={"speaker": speaker_id},
        json=query_data,
        timeout=config.VOICEVOX_SYNTHESIS_TIMEOUT,
    )
    synthesis_response.raise_for_status()
    return synthesis_response.content


def decode_wav_bytes(wav_bytes: bytes) -> tuple[np.ndarray, int]:
    with wave.open(io.BytesIO(wav_bytes), "rb") as wav_file:
        channels = wav_file.getnchannels()
        sample_width = wav_file.getsampwidth()
        sample_rate = wav_file.getframerate()
        frames = wav_file.readframes(wav_file.getnframes())

    if sample_width == 1:
        audio = np.frombuffer(frames, dtype=np.uint8).astype(np.float32)
        audio = (audio - 128.0) / 128.0
    elif sample_width == 2:
        audio = np.frombuffer(frames, dtype=np.int16).astype(np.float32) / 32768.0
    elif sample_width == 4:
        audio = np.frombuffer(frames, dtype=np.int32).astype(np.float32) / 2147483648.0
    else:
        raise ValueError(f"未対応のサンプル幅です: {sample_width}")

    if channels > 1:
        audio = audio.reshape(-1, channels)

    return audio, sample_rate


def find_output_device(device_name: str) -> int:
    normalized = device_name.casefold()
    devices = sd.query_devices()

    exact_match = None
    partial_match = None

    for index, device in enumerate(devices):
        if int(device.get("max_output_channels", 0)) < 1:
            continue

        current_name = str(device.get("name", ""))
        current_normalized = current_name.casefold()

        if current_normalized == normalized:
            exact_match = index
            break

        if normalized in current_normalized and partial_match is None:
            partial_match = index

    if exact_match is not None:
        return exact_match

    if partial_match is not None:
        return partial_match

    output_names = [
        str(device.get("name", ""))
        for device in devices
        if int(device.get("max_output_channels", 0)) >= 1
    ]
    raise ValueError(
        "指定した出力デバイスが見つかりません: "
        f"{device_name}. 利用可能な出力デバイス: {', '.join(output_names)}"
    )


def get_default_output_device() -> int:
    default_device = sd.default.device

    try:
        output_device = default_device[1]
    except (TypeError, IndexError, KeyError):
        output_device = default_device

    if output_device is None:
        raise ValueError("既定の出力デバイスが設定されていません。")

    output_device_index = int(output_device)

    if output_device_index < 0:
        raise ValueError("既定の出力デバイスが設定されていません。")

    return output_device_index


def resolve_output_device(device_name: str) -> int:
    if not device_name.strip():
        return get_default_output_device()

    matched = config.DEVICE_SELECTOR_PATTERN.match(device_name)
    if matched:
        return int(matched.group(1))

    return find_output_device(device_name)


def prepare_audio_for_device(audio: np.ndarray, device_index: int) -> np.ndarray:
    device_info = sd.query_devices(device_index)
    max_output_channels = int(device_info.get("max_output_channels", 0))

    if audio.ndim == 1:
        if max_output_channels >= 2:
            return np.column_stack((audio, audio))
        return audio.reshape(-1, 1)

    if audio.shape[1] <= max_output_channels:
        return audio

    return audio.mean(axis=1, keepdims=True)


def play_audio_stream(audio: np.ndarray, sample_rate: int, device_index: int) -> None:
    frame_count = audio.shape[0]
    channels = 1 if audio.ndim == 1 else audio.shape[1]
    next_frame = 0
    finished = threading.Event()

    if audio.ndim == 1:
        stream_audio = audio.reshape(-1, 1)
    else:
        stream_audio = audio

    def callback(
        outdata: np.ndarray, frames: int, _time: object, status: sd.CallbackFlags
    ) -> None:
        nonlocal next_frame

        if status:
            print(f"音声再生ステータス: {status}", file=sys.stderr)

        chunk = stream_audio[next_frame : next_frame + frames]
        outdata.fill(0)
        outdata[: len(chunk)] = chunk
        next_frame += len(chunk)

        if next_frame >= frame_count:
            raise sd.CallbackStop()

    with sd.OutputStream(
        samplerate=sample_rate,
        device=device_index,
        channels=channels,
        dtype="float32",
        callback=callback,
        finished_callback=finished.set,
    ):
        finished.wait()


def play_wav_bytes(wav_bytes: bytes, settings: dict[str, object]) -> None:
    audio, sample_rate = decode_wav_bytes(wav_bytes)
    device_targets: list[tuple[str, int]] = []

    # VB-Cable（リップシンク用 — デバイス名が設定されている場合のみ送信）
    vb_cable_name = str(settings.get("vbcable_device_name", "")).strip()
    if vb_cable_name:
        try:
            vb_cable_device_index = resolve_output_device(vb_cable_name)
            device_targets.append(("VB-Cable", vb_cable_device_index))
        except ValueError as error:
            vb_cable_device_index = -1
            print(f"VB-Cable が見つからないためリップシンク送信をスキップしました: {error}", file=sys.stderr)
    else:
        vb_cable_device_index = -1

    # スピーカー / ヘッドホン（ユーザーが声を聞く用）
    if bool(settings.get("monitor_playback_enabled", True)):
        try:
            monitor_device_index = resolve_output_device(
                str(settings.get("monitor_device_name", ""))
            )
            if monitor_device_index != vb_cable_device_index:
                device_targets.append(("スピーカー", monitor_device_index))
        except ValueError as error:
            print(f"スピーカー再生をスキップしました: {error}", file=sys.stderr)

    threads: list[threading.Thread] = []
    errors: list[str] = []
    errors_lock = threading.Lock()

    for device_label, device_index in device_targets:
        prepared_audio = prepare_audio_for_device(audio, device_index)

        def play_target(
            target_audio: np.ndarray = prepared_audio,
            target_rate: int = sample_rate,
            target_index: int = device_index,
            target_label: str = device_label,
        ) -> None:
            try:
                play_audio_stream(target_audio, target_rate, target_index)
            except Exception as error:
                with errors_lock:
                    errors.append(f"{target_label}: {error}")

        thread = threading.Thread(target=play_target, daemon=True)
        thread.start()
        threads.append(thread)

    for thread in threads:
        thread.join()

    if errors:
        raise RuntimeError(" / ".join(errors))


# ── 履歴ログ ──────────────────────────────────────────────

_LOG_MAX_BYTES = 5 * 1024 * 1024  # 5 MB


def append_log(text: str, expression_id: int, settings: dict[str, object]) -> None:
    """要約文と表情IDを JSONL ファイルに追記する。上限超過時は古い半分を切り捨てる。"""
    if not bool(settings.get("avatar_log_enabled", False)):
        return

    # 旧ファイル → スロット1 マイグレーション
    legacy = Path(__file__).with_name(config.LEGACY_LOG_FILE)
    slot1 = Path(__file__).with_name(config.LOG_SLOT_FILES[0])
    if legacy.is_file() and not slot1.is_file():
        try:
            legacy.rename(slot1)
        except OSError:
            pass

    log_path = config.get_active_log_path(settings)
    lock_path = log_path.with_suffix(".lock")

    try:
        _append_log_locked(log_path, lock_path, text, expression_id)
    except OSError as error:
        print(f"ログ書き込みに失敗しました: {error}", file=sys.stderr)


def _append_log_locked(
    log_path: Path, lock_path: Path, text: str, expression_id: int
) -> None:
    """ロックファイルで直列化し、重複チェック→書き込みをアトミックに行う。"""
    import msvcrt

    lock_fd = os.open(str(lock_path), os.O_CREAT | os.O_RDWR)
    try:
        # ノンブロッキングで最大2秒リトライしてロック取得
        _lock_deadline = time.monotonic() + 2.0
        while True:
            try:
                msvcrt.locking(lock_fd, msvcrt.LK_NBLCK, 1)
                break
            except OSError:
                if time.monotonic() >= _lock_deadline:
                    raise
                time.sleep(0.05)
        try:
            # 直前エントリと同一内容なら重複として弾く
            if log_path.exists():
                try:
                    with open(log_path, "rb") as f:
                        f.seek(0, 2)
                        pos = f.tell()
                        buf = b""
                        while pos > 0 and b"\n" not in buf.lstrip(b"\n"):
                            chunk = min(pos, 512)
                            pos -= chunk
                            f.seek(pos)
                            buf = f.read(chunk) + buf
                        last_line = buf.rstrip(b"\n").split(b"\n")[-1]
                        if last_line:
                            prev = json.loads(last_line)
                            if (prev.get("text") == text
                                    and prev.get("expression_id") == expression_id):
                                return
                except (json.JSONDecodeError, KeyError):
                    pass

            # サイズ上限チェック — 超過時は後半だけ残す
            if log_path.exists() and log_path.stat().st_size > _LOG_MAX_BYTES:
                try:
                    lines = log_path.read_text(encoding="utf-8").splitlines()
                    half = len(lines) // 2
                    tmp_fd, tmp_path = tempfile.mkstemp(
                        dir=log_path.parent, suffix=".tmp"
                    )
                    try:
                        with os.fdopen(tmp_fd, "w", encoding="utf-8") as tmp_f:
                            tmp_f.write("\n".join(lines[half:]) + "\n")
                        os.replace(tmp_path, log_path)
                    except BaseException:
                        with contextlib.suppress(OSError):
                            os.unlink(tmp_path)
                        raise
                except OSError as trim_err:
                    print(f"ログ切り詰めに失敗しました: {trim_err}", file=sys.stderr)

            # 追記
            entry = {
                "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
                "expression_id": expression_id,
                "text": text,
            }
            line = json.dumps(entry, ensure_ascii=False) + "\n"
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(line)
        finally:
            os.lseek(lock_fd, 0, os.SEEK_SET)
            msvcrt.locking(lock_fd, msvcrt.LK_UNLCK, 1)
    finally:
        os.close(lock_fd)


# ── メイン ────────────────────────────────────────────────

def main() -> int:
    args = parse_args()
    settings = config.load_settings()

    try:
        wav_bytes = synthesize_voice(args.text, settings, args.expression_id)
    except (requests.RequestException, ValueError) as error:
        print(
            f"VOICEVOX音声合成をスキップし、ホットキー送信のみ実行します: {error}",
            file=sys.stderr,
        )
        send_expression(args.expression_id, settings)
        append_log(args.text, args.expression_id, settings)
        return 0

    send_expression(args.expression_id, settings)

    try:
        play_wav_bytes(wav_bytes, settings)
    except Exception as error:
        print(f"音声再生に失敗しました: {error}", file=sys.stderr)
        append_log(args.text, args.expression_id, settings)
        return 1

    append_log(args.text, args.expression_id, settings)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
