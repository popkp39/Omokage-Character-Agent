"""旧バージョンからのデータ移行スクリプト。

使い方:
    python src/_migrate_data.py
    （フォルダ選択ダイアログで旧バージョンのフォルダを指定）
"""

from __future__ import annotations

import ctypes
import json
import os
import shutil
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

VENV_DIR_NAME = ".venv"
SETTINGS_FILE_NAME = "avatar_settings.json"
PRESET_DIR_NAME = "CharacterPresets"
PERSONAS_DIR_NAME = "AIPersonas"
LOG_FILE_PATTERN = "avatar_log_*.jsonl"
LEGACY_LOG_FILE = "avatar_log.jsonl"
SLOT1_LOG_FILE = "avatar_log_1.jsonl"
_PERSONA_EXTENSIONS = {".txt", ".md"}

def _is_exact_assignment(line: str, name: str) -> bool:
    """行が `name = ...` または `name: type = ...` の代入文か判定する。

    BUG10修正: `PRESET_VERSION_NEXT` のような前方一致の誤マッチを防ぐ。
    """
    if not line.startswith(name):
        return False
    rest = line[len(name):]
    if not rest:
        return False
    # name の直後が空白, =, : のいずれかなら代入文
    return rest[0] in (" ", "\t", "=", ":")


def _strip_inline_comment(value: str) -> str:
    """値文字列からインラインコメントを除去する。

    BUG11修正: `1  # comment` → `1`
    クォート内の # は無視する。
    """
    in_quote: str | None = None
    for i, ch in enumerate(value):
        if in_quote:
            if ch == in_quote:
                in_quote = None
        elif ch in ("'", '"'):
            in_quote = ch
        elif ch == "#":
            return value[:i].rstrip()
    return value


def _read_config_int(name: str, default: int, config_path: Path | None = None) -> int:
    """config.py のソースから整数定数を読み取る（インポート副作用回避）。

    BUG-P修正: = を含まない型アノテーション専用行の IndexError を防止。
    BUG-Q修正: Python と同じく最後のマッチを採用する。
    """
    if config_path is None:
        config_path = Path(__file__).with_name("config.py")
    result = default
    try:
        for line in config_path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if _is_exact_assignment(stripped, name) and "=" in stripped:
                value = stripped.split("=", 1)[1].strip()
                value = _strip_inline_comment(value)
                result = int(value)
    except (OSError, ValueError):
        pass
    return result


def _read_config_str(name: str, default: str, config_path: Path | None = None) -> str:
    """config.py のソースから文字列定数を読み取る（引用符を除去）。

    BUG-P修正: = を含まない型アノテーション専用行の IndexError を防止。
    BUG-Q修正: Python と同じく最後のマッチを採用する。
    """
    if config_path is None:
        config_path = Path(__file__).with_name("config.py")
    result = default
    try:
        for line in config_path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if _is_exact_assignment(stripped, name) and "=" in stripped:
                value = stripped.split("=", 1)[1].strip()
                value = _strip_inline_comment(value)
                result = value.strip("\"'")
    except (OSError, ValueError):
        pass
    return result


SETTINGS_FILE_VERSION = _read_config_int("SETTINGS_FILE_VERSION", 1)
PRESET_VERSION = _read_config_int("PRESET_VERSION", 1)
APP_VERSION = _read_config_str("APP_VERSION", "不明")

# 設定ファイル内の書き換え対象パスキー
_PATH_KEYS_IN_SETTINGS = ("summary_system_prompt_path",)
# プリセットJSON内の書き換え対象パスキー（character セクション内）
_PATH_KEYS_IN_PRESET = ("summary_system_prompt_path",)

# ── 不可視文字検出（ASCII Smuggling / プロンプトインジェクション対策） ──
# 許可する制御文字: TAB(0x09), LF(0x0A), CR(0x0D)
_SAFE_CONTROL = {"\t", "\n", "\r"}

# 危険な不可視文字の範囲（検出対象）
def _is_suspicious_char(ch: str) -> bool:
    """AI に対するプロンプトインジェクションに悪用されうる不可視文字を検出する。"""
    cp = ord(ch)
    # C0制御文字（TAB/LF/CR以外）
    if cp <= 0x1F and ch not in _SAFE_CONTROL:
        return True
    # DEL
    if cp == 0x7F:
        return True
    # C1制御文字
    if 0x80 <= cp <= 0x9F:
        return True
    # Soft hyphen（表示されないがテキストに存在）
    if cp == 0xAD:
        return True
    # Zero-width / 結合制御 / Bidi marks
    if cp in (0x200B, 0x200C, 0x200D, 0x200E, 0x200F, 0x2060, 0xFEFF):
        return True
    # Bidi override / embedding
    if 0x202A <= cp <= 0x202E or 0x2066 <= cp <= 0x2069:
        return True
    # Unicode Tag Characters（ASCII Smuggling の本体）
    if 0xE0001 <= cp <= 0xE007F:
        return True
    # Variation Selectors Supplement（テキストに不要）
    if 0xE0100 <= cp <= 0xE01EF:
        return True
    return False


_SUSPICIOUS_NAMES: dict[int, str] = {
    0x00: "NULL",
    0x08: "BACKSPACE",
    0x7F: "DEL",
    0xAD: "SOFT HYPHEN",
    0x200B: "ZWSP",
    0x200C: "ZWNJ",
    0x200D: "ZWJ",
    0x200E: "LRM",
    0x200F: "RLM",
    0x202A: "LRE",
    0x202B: "RLE",
    0x202C: "PDF",
    0x202D: "LRO",
    0x202E: "RLO",
    0x2060: "WJ",
    0x2066: "LRI",
    0x2067: "RLI",
    0x2068: "FSI",
    0x2069: "PDI",
    0xFEFF: "BOM/ZWNBSP",
}


def _scan_suspicious_chars(text: str, max_report: int = 5) -> list[str]:
    """テキスト内の危険な不可視文字を検出し、説明リストを返す。"""
    found: dict[int, list[int]] = {}  # codepoint → [line_numbers]
    for line_no, line in enumerate(text.splitlines(), 1):
        for ch in line:
            if _is_suspicious_char(ch):
                cp = ord(ch)
                found.setdefault(cp, []).append(line_no)
    if not found:
        return []
    results: list[str] = []
    for cp, lines in sorted(found.items()):
        name = _SUSPICIOUS_NAMES.get(cp, "")
        if not name:
            if 0xE0001 <= cp <= 0xE007F:
                name = f"TAG '{chr(cp - 0xE0000)}'"
            elif 0x80 <= cp <= 0x9F:
                name = "C1制御文字"
            else:
                name = "不可視文字"
        line_info = ", ".join(str(n) for n in lines[:max_report])
        if len(lines) > max_report:
            line_info += f" 他{len(lines) - max_report}箇所"
        results.append(f"U+{cp:04X} ({name}) - {len(lines)}個 (行: {line_info})")
    return results


def find_migration_targets(src_dir: Path) -> dict[str, list[Path]]:
    """移行対象のファイル/フォルダを検出する。"""
    targets: dict[str, list[Path]] = {
        "settings": [],
        "presets": [],
        "personas": [],
        "logs": [],
    }

    settings_file = src_dir / SETTINGS_FILE_NAME
    if settings_file.is_file():
        targets["settings"].append(settings_file)

    preset_dir = src_dir / PRESET_DIR_NAME
    if preset_dir.is_dir():
        targets["presets"] = sorted(preset_dir.glob("*.json"))

    # BUG-R修正: 隠しファイル・一時ファイルを除外（presets の .json フィルタと一貫性を持たせる）
    personas_dir = src_dir / PERSONAS_DIR_NAME
    if personas_dir.is_dir():
        targets["personas"] = sorted(
            f for f in personas_dir.iterdir()
            if f.is_file() and f.suffix.lower() in _PERSONA_EXTENSIONS
        )

    for log_file in sorted(src_dir.glob(LOG_FILE_PATTERN)):
        targets["logs"].append(log_file)
    legacy_log = src_dir / LEGACY_LOG_FILE
    if legacy_log.is_file():
        targets["logs"].append(legacy_log)

    return targets


def _has_migration_data(d: Path) -> bool:
    """指定ディレクトリに移行可能なデータが存在するか簡易チェックする。

    BUG-W修正: personas/presets はディレクトリ存在だけでなく対象ファイルの有無もチェック。
    BUG-Y修正: 遅延評価にして PermissionError を回避 + 短絡評価を活かす。
    """
    if (d / SETTINGS_FILE_NAME).is_file():
        return True
    preset_dir = d / PRESET_DIR_NAME
    if preset_dir.is_dir() and any(preset_dir.glob("*.json")):
        return True
    personas_dir = d / PERSONAS_DIR_NAME
    try:
        if personas_dir.is_dir() and any(
            f.suffix.lower() in _PERSONA_EXTENSIONS
            for f in personas_dir.iterdir()
            if f.is_file()
        ):
            return True
    except OSError:
        pass
    if any(d.glob(LOG_FILE_PATTERN)):
        return True
    if (d / LEGACY_LOG_FILE).is_file():
        return True
    return False


def validate_source(source_root: Path) -> Path | None:
    """旧バージョンのsrcディレクトリを特定する。

    BUG-X修正: データが見つからない場合は None を返す
    （空の src/ を返してエラーメッセージが混乱する問題を防止）。
    """
    # src/ 内にデータがある構成
    src_dir = source_root / "src"
    # BUG-O修正: src/ が存在しても中にデータがなければルート直下を優先
    if src_dir.is_dir() and _has_migration_data(src_dir):
        return src_dir
    # ルート直下にデータがある場合（src/ なし構成、src/ を直接指定、など）
    if _has_migration_data(source_root):
        return source_root
    return None


def _extract_settings_version(data: dict) -> int | None:
    """設定ファイルから実際の設定バージョンを取得する（ラッパー対応）。

    BUG-H修正: plaintext/DPAPIラッパー形式の場合、外側の version は
    ラッパーのバージョンであり、設定データのバージョンではない。
    内側の data.version を優先して返す。
    """
    fmt = data.get("format")
    if fmt == "plaintext" and isinstance(data.get("data"), dict):
        return data["data"].get("version")
    if fmt == "windows-dpapi":
        # 暗号化されているため内部バージョンは不明 → 外側を信用するしかない
        return data.get("version")
    # ラッパーなし（旧形式: 生JSON）
    return data.get("version")


def _is_dpapi_encrypted(settings_path: Path) -> bool:
    """設定ファイルがDPAPI暗号化されているか判定する。"""
    try:
        data = json.loads(settings_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError, ValueError):
        return False
    return isinstance(data, dict) and data.get("format") == "windows-dpapi"


def check_settings_version(settings_path: Path) -> tuple[bool, str]:
    """設定ファイルのバージョン互換性を確認する。"""
    try:
        data = json.loads(settings_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError, ValueError) as e:
        return False, f"設定ファイルの読み込みに失敗: {e}"

    if not isinstance(data, dict):
        return False, "設定ファイルの形式が不正です"

    version = _extract_settings_version(data)
    if version is None:
        return True, "バージョン情報なし（旧形式、移行可能）"

    # BUG-I修正: 非整数値の場合はエラーとして扱う
    if not isinstance(version, int):
        return False, f"バージョン値が不正です（{version!r}）"

    if version > SETTINGS_FILE_VERSION:
        return False, (
            f"設定ファイルのバージョン({version})が"
            f"このバージョン({SETTINGS_FILE_VERSION})より新しいため移行できません"
        )
    return True, f"バージョン {version}（互換性あり）"


def check_preset_version(preset_path: Path) -> tuple[bool, str]:
    """プリセットファイルのバージョン互換性を確認する。"""
    try:
        data = json.loads(preset_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError, ValueError) as e:
        return False, f"読み込み失敗: {e}"

    if not isinstance(data, dict):
        return False, "形式が不正"

    version = data.get("preset_version")
    if version is None:
        return True, ""  # 旧形式（バージョンキーなし）

    # BUG-I修正: 非整数値の場合はエラーとして扱う
    if not isinstance(version, int):
        return False, f"バージョン値が不正（{version!r}）"

    if version > PRESET_VERSION:
        return False, (
            f"バージョン {version} はこのバージョン({PRESET_VERSION})より新しいため移行不可"
        )
    return True, ""


def _rewrite_paths_in_settings(
    settings_path: Path, old_src: Path, new_src: Path
) -> bool:
    """設定ファイル内の絶対パスを新しいパスに書き換える（BUG1修正）。

    DPAPI暗号化の場合は復号→パス書き換え→再暗号化する。
    復号に失敗した場合は警告を出してスキップする。
    平文JSONの場合はキーを直接書き換える。
    戻り値: 書き換えが発生したら True。
    """
    try:
        raw = settings_path.read_text(encoding="utf-8")
    except (OSError, ValueError):
        return False

    old_prefix = _normalize_path_str(str(old_src))
    new_prefix = _normalize_path_str(str(new_src))

    if old_prefix.lower() == new_prefix.lower():
        return False

    # DPAPI暗号化: 復号→書き換え→再暗号化
    # 平文ラッパー/旧形式: JSONパースして直接書き換え
    try:
        data = json.loads(raw)
    except (json.JSONDecodeError, ValueError):
        return False

    if not isinstance(data, dict):
        return False

    fmt = data.get("format")

    # DPAPI暗号化の場合: 復号→パス書き換え→再暗号化
    if fmt == "windows-dpapi":
        try:
            import importlib
            config_mod = importlib.import_module("config")
            inner = config_mod.decrypt_settings_payload(data)
            changed = _replace_path_values(
                inner, old_prefix, new_prefix, _PATH_KEYS_IN_SETTINGS
            )
            if changed:
                encrypted = config_mod.encrypt_settings_payload(inner)
                content = json.dumps(encrypted, ensure_ascii=False, indent=2) + "\n"
                settings_path.write_text(content, encoding="utf-8")
            return changed
        except Exception as e:
            print(f"  [WARN] DPAPI設定のパス書き換えに失敗しました: {e}")
            print(f"         設定画面でキャラプロンプトのパスを手動で修正してください。")
            return False

    # 平文ラッパー {"format": "plaintext", "version": 1, "data": {...}}
    if fmt == "plaintext" and isinstance(data.get("data"), dict):
        inner = data["data"]
        changed = _replace_path_values(inner, old_prefix, new_prefix, _PATH_KEYS_IN_SETTINGS)
        if changed:
            content = json.dumps(data, ensure_ascii=False, indent=2) + "\n"
            settings_path.write_text(content, encoding="utf-8")
        return changed

    # ラッパーなし（旧形式: 生JSON）
    changed = _replace_path_values(data, old_prefix, new_prefix, _PATH_KEYS_IN_SETTINGS)
    if changed:
        content = json.dumps(data, ensure_ascii=False, indent=2) + "\n"
        settings_path.write_text(content, encoding="utf-8")
    return changed


def _rewrite_paths_in_preset(
    preset_path: Path, old_src: Path, new_src: Path
) -> bool:
    """プリセットJSON内の絶対パスを新しいパスに書き換える。"""
    try:
        data = json.loads(preset_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError, ValueError):
        return False

    if not isinstance(data, dict):
        return False

    old_prefix = _normalize_path_str(str(old_src))
    new_prefix = _normalize_path_str(str(new_src))

    if old_prefix.lower() == new_prefix.lower():
        return False

    changed = False
    # トップレベル
    changed |= _replace_path_values(data, old_prefix, new_prefix, _PATH_KEYS_IN_PRESET)
    # character セクション内
    char = data.get("character")
    if isinstance(char, dict):
        changed |= _replace_path_values(char, old_prefix, new_prefix, _PATH_KEYS_IN_PRESET)

    if changed:
        content = json.dumps(data, ensure_ascii=False, indent=2) + "\n"
        preset_path.write_text(content, encoding="utf-8")
    return changed


def _normalize_path_str(path: str) -> str:
    """パス文字列を正規化する（スラッシュ統一 + 末尾スラッシュ除去）。

    BUG-N修正: ドライブルート（例: "D:/"）の末尾スラッシュを除去して
    prefix + "/" の二重スラッシュ問題を防止する。
    """
    normalized = path.replace("\\", "/")
    # "D:/" → "D:" のように末尾スラッシュを除去（ただしルート "/" は保持）
    if len(normalized) > 1 and normalized.endswith("/"):
        normalized = normalized.rstrip("/")
    return normalized


def _replace_path_values(
    d: dict, old_prefix: str, new_prefix: str, keys: tuple[str, ...]
) -> bool:
    """辞書内の指定キーの値に含まれるパスプレフィクスを置換する。

    BUG6修正: 値のスラッシュ方向を正規化してから比較・置換する。
    書き戻す際もフォワードスラッシュに統一する。
    BUG-M修正: Windows のパスは大文字小文字を区別しないため
    case-insensitive で比較する（書き戻しは new_prefix の casing を使用）。
    """
    changed = False
    old_lower = old_prefix.lower()
    for key in keys:
        val = d.get(key)
        if isinstance(val, str):
            normalized = _normalize_path_str(val)
            norm_lower = normalized.lower()
            # BUG-B修正: 前方一致+パス区切り境界チェック（部分一致の誤置換を防止）
            if norm_lower.startswith(old_lower + "/") or norm_lower == old_lower:
                d[key] = new_prefix + normalized[len(old_prefix):]
                changed = True
    return changed


def _resolve_legacy_log(
    targets: dict[str, list[Path]],
) -> tuple[str | None, Path | None]:
    """レガシーログの移行先を適切に決定する（BUG2修正）。

    avatar_log.jsonl は新版では avatar_log_1.jsonl にリネームされるべき。
    - avatar_log_1.jsonl が移行対象に含まれていない場合 → リネームして移行
    - avatar_log_1.jsonl が既に含まれている場合 → legacy を除外し、
      slot1 コピー後に legacy を追記マージ（BUG-S修正: データ欠損防止）

    BUG-G修正: メッセージを返却し、呼び出し元でヘッダー後に表示する。
    BUG-V修正: legacy をリストから除外して件数表示を正確にする。
    追記元の Path を第2戻り値で返し、コピーフェーズで使用する。
    """
    legacy_paths = [p for p in targets["logs"] if p.name == LEGACY_LOG_FILE]
    if not legacy_paths:
        return None, None

    has_slot1 = any(p.name == SLOT1_LOG_FILE for p in targets["logs"])

    if has_slot1:
        # legacy をリストから除外（件数を正確にする）
        # コピーフェーズで slot1 処理後に legacy を追記マージする
        targets["logs"] = [p for p in targets["logs"] if p.name != LEGACY_LOG_FILE]
        msg = (
            f"  [情報] {LEGACY_LOG_FILE} と {SLOT1_LOG_FILE} の両方が存在します\n"
            f"         {LEGACY_LOG_FILE} の内容は {SLOT1_LOG_FILE} に追記マージします"
        )
        return msg, legacy_paths[0]

    # slot1 が無い → レガシーログを avatar_log_1.jsonl として移行するためマーク
    # （実際のリネームはコピー時に dst_path を変更して対応）
    return None, None


def _get_log_dst_path(log_path: Path, dst_dir: Path) -> Path:
    """ログファイルの移行先パスを返す。レガシーログはslot1にリネーム。"""
    if log_path.name == LEGACY_LOG_FILE:
        return dst_dir / SLOT1_LOG_FILE
    return dst_dir / log_path.name


def _get_log_display_name(log_path: Path) -> str:
    """ログファイルの表示名を返す。"""
    if log_path.name == LEGACY_LOG_FILE:
        return f"{LEGACY_LOG_FILE} → {SLOT1_LOG_FILE}"
    return log_path.name


def copy_file_safe(src: Path, dst: Path, *, overwrite: bool = False) -> bool:
    """ファイルを安全にコピーする。"""
    if dst.exists() and not overwrite:
        return False
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(src), str(dst))
    return True


def copy_venv(src_venv: Path, dst_venv: Path) -> bool:
    """旧バージョンの .venv をコピーする（スピナー付き）。

    コピー失敗時は中途半端なディレクトリを削除して False を返す。
    """
    done = False
    error: Exception | None = None

    def _copy() -> None:
        nonlocal done, error
        try:
            shutil.copytree(str(src_venv), str(dst_venv))
        except Exception as e:
            error = e
        done = True

    t = threading.Thread(target=_copy)
    t.start()

    spinner = ["/", "-", "\\", "|"]
    i = 0
    while not done:
        print(f"  {spinner[i % 4]} .venv をコピー中...", end="\r", flush=True)
        time.sleep(0.2)
        i += 1

    if error is not None:
        print(f"  [NG] .venv のコピーに失敗: {error}       ")
        # 中途半端なコピーを削除
        if dst_venv.exists():
            shutil.rmtree(str(dst_venv), ignore_errors=True)
        return False

    print(f"  [OK] .venv をコピーしました                ")
    return True


def append_log_file(src: Path, dst: Path) -> bool:
    """ログファイルを既存ファイルに追記マージする（.jsonl形式）。

    BUG-A修正: ログは上書きではなく追記して既存データを保全する。
    """
    dst.parent.mkdir(parents=True, exist_ok=True)
    # BUG-U修正: BOM付きUTF-8にも対応（utf-8-sigはBOMがあれば除去、なければそのまま読む）
    src_text = src.read_text(encoding="utf-8-sig")
    if not src_text.strip():
        return False
    # 既存ファイルの末尾が改行でなければ改行を挿入
    if dst.exists() and dst.stat().st_size > 0:
        existing = dst.read_text(encoding="utf-8-sig")
        if existing and not existing.endswith("\n"):
            src_text = "\n" + src_text
    with open(dst, "a", encoding="utf-8") as f:
        f.write(src_text)
    return True


def run_migration_cli() -> int:
    """CLI モードで移行を実行する。"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        return _run_migration_inner(root)
    finally:
        # BUG4修正: 早期リターンでもrootを必ず破棄
        root.destroy()


def _run_migration_inner(root: tk.Tk) -> int:
    """移行処理の本体。root.destroy() は呼び出し元が保証する。"""
    # BUG-F修正: ダイアログが背面に隠れないようフォーカスを強制取得
    root.focus_force()

    source_root_str = filedialog.askdirectory(
        title="旧バージョンのフォルダを選択してください",
        mustexist=True,
    )

    # フォルダ選択ダイアログ後にコンソールウィンドウへフォーカスを戻す
    # batから直接起動するとコンソールが一度もフォアグラウンドにならないため、
    # Altキーイベントで SetForegroundWindow の制限を回避する
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)  # Alt press
            ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)  # Alt release
            ctypes.windll.user32.SetForegroundWindow(hwnd)
    except Exception:
        pass

    if not source_root_str:
        print("\nキャンセルされました。")
        return 1

    source_root = Path(source_root_str)
    src_dir = validate_source(source_root)
    if src_dir is None:
        print(f"\n[ERROR] 指定されたフォルダに移行可能なデータが見つかりません:")
        print(f"  {source_root}")
        print("\nOmokage-Character-Agent のルートフォルダを指定してください。")
        return 1

    dst_dir = Path(__file__).resolve().parent

    # 同じフォルダを指定した場合
    try:
        if src_dir.resolve() == dst_dir.resolve():
            print("\n[ERROR] 移行元と移行先が同じフォルダです。")
            print("旧バージョンのフォルダを指定してください。")
            return 1
    except OSError:
        print("\n[ERROR] フォルダのパス解決に失敗しました。")
        print("ネットワークドライブ等の場合は直接パスを指定してください。")
        return 1

    targets = find_migration_targets(src_dir)

    # BUG2修正: レガシーログの移行先を整理
    # BUG-V修正: legacy_log_to_append は slot1 コピー後に追記マージする元ファイル
    legacy_log_msg, legacy_log_to_append = _resolve_legacy_log(targets)

    # 移行対象の表示
    total_files = sum(len(v) for v in targets.values())
    if total_files == 0:
        print(f"\n移行可能なデータが見つかりませんでした。")
        print(f"  検索先: {src_dir}")
        return 1

    # 移行元のバージョンを読み取り
    old_config = src_dir / "config.py"
    old_version = _read_config_str("APP_VERSION", "不明", old_config)

    print()
    print("=" * 50)
    print(f"  データ移行  v{old_version} → v{APP_VERSION}")
    print("=" * 50)
    print(f"  移行元: {src_dir}")
    print(f"  移行先: {dst_dir}")
    print()

    # 設定ファイル
    dpapi_warned = False
    if targets["settings"]:
        settings_path = targets["settings"][0]
        compatible, msg = check_settings_version(settings_path)
        status = "[OK]" if compatible else "[NG]"
        print(f"  {status} 設定ファイル ({msg})")
        if not compatible:
            print(f"       → スキップします")
            targets["settings"] = []
        elif _is_dpapi_encrypted(settings_path):
            dpapi_warned = True
            print()
            print("  [注意] 設定ファイルはDPAPI暗号化されています。")
            print("  DPAPI暗号化は Windowsユーザーアカウントに紐づいているため、")
            print("  別のPC・別のユーザーへの移行では復号できません。")
            print("  （同じPC・同じユーザーでのフォルダ移動であれば問題ありません）")
            print()
            print("  (y) 設定ファイルを移行する（同一PC向け）")
            print("  (n) 設定ファイルをスキップ（別PC向け: 移行先で再設定）")
            answer = input("\n  設定ファイルを移行しますか？ (y/N): ").strip().lower()
            if answer not in ("y", "yes"):
                print("       → 設定ファイルをスキップします")
                targets["settings"] = []
    else:
        print("  [--] 設定ファイル: なし")

    # プリセット（バージョンチェック付き）
    if targets["presets"]:
        ok_presets: list[Path] = []
        ng_presets: list[tuple[Path, str]] = []
        for p in targets["presets"]:
            compatible, msg = check_preset_version(p)
            if compatible:
                ok_presets.append(p)
            else:
                ng_presets.append((p, msg))
        if ok_presets:
            print(f"  [OK] プリセット: {len(ok_presets)} 件")
            for p in ok_presets:
                print(f"       - {p.name}")
        if ng_presets:
            print(f"  [NG] プリセット: {len(ng_presets)} 件スキップ")
            for p, msg in ng_presets:
                print(f"       - {p.name} ({msg})")
        targets["presets"] = ok_presets
    else:
        print("  [--] プリセット: なし")

    # キャラ設定
    count = len(targets["personas"])
    if count > 0:
        print(f"  [OK] キャラ設定: {count} 件")
        for p in targets["personas"]:
            print(f"       - {p.name}")
    else:
        print("  [--] キャラ設定: なし")

    # ログ
    count = len(targets["logs"])
    if count > 0:
        print(f"  [OK] ログファイル: {count} 件")
        for p in targets["logs"]:
            print(f"       - {_get_log_display_name(p)}")
    else:
        print("  [--] ログファイル: なし")

    # セットアップ環境(.venv)
    # Bug-C修正: src/ を直接選択した場合、親フォルダの .venv も探す
    old_venv = source_root / VENV_DIR_NAME
    if not old_venv.is_dir() and src_dir.parent != source_root:
        old_venv = src_dir.parent / VENV_DIR_NAME
    new_venv = dst_dir.parent / VENV_DIR_NAME
    copy_venv_flag = False
    if old_venv.is_dir():
        if new_venv.is_dir():
            print(f"  [--] セットアップ環境(.venv): 移行先に既に存在（スキップ）")
        else:
            print(f"  [OK] セットアップ環境(.venv): 検出")
            print(f"       コピーすると初回セットアップが不要になります")
            print(f"       ※ 同じPC・同じPythonバージョンの場合のみ有効")
    else:
        print(f"  [--] セットアップ環境(.venv): なし")

    migratable = sum(len(v) for v in targets.values())
    if migratable == 0 and not (old_venv.is_dir() and not new_venv.is_dir()):
        print("\n移行可能なデータがありません。")
        return 1

    # ログは任意
    if targets["logs"]:
        print()
        answer = input("ログファイルも移行しますか？ (y/N): ").strip().lower()
        if answer not in ("y", "yes"):
            targets["logs"] = []
            legacy_log_to_append = None  # Bug1修正: ログ拒否時は legacy も除外
        elif legacy_log_msg:
            # ログ移行が承認された場合のみ legacy マージ情報を表示
            print(legacy_log_msg)

    # .venv コピー確認
    if old_venv.is_dir() and not new_venv.is_dir():
        print()
        answer = input("セットアップ環境(.venv)もコピーしますか？ (Y/n): ").strip().lower()
        if answer not in ("n", "no"):
            copy_venv_flag = True

    # BUG5修正: 競合ファイルごとに個別で上書き判定
    skip_files: set[str] = set()
    merge_log_files: set[str] = set()  # BUG-A修正: 追記マージ対象のログ
    conflicts: list[tuple[str, str]] = []  # (カテゴリ, 表示名)
    log_conflicts: list[tuple[str, str]] = []  # ログ専用の競合リスト

    if targets["settings"]:
        dst_settings = dst_dir / SETTINGS_FILE_NAME
        if dst_settings.exists():
            conflicts.append(("settings", SETTINGS_FILE_NAME))
    for p in targets["presets"]:
        dst_p = dst_dir / PRESET_DIR_NAME / p.name
        if dst_p.exists():
            conflicts.append(("presets", f"{PRESET_DIR_NAME}/{p.name}"))
    for p in targets["personas"]:
        dst_p = dst_dir / PERSONAS_DIR_NAME / p.name
        if dst_p.exists():
            conflicts.append(("personas", f"{PERSONAS_DIR_NAME}/{p.name}"))
    # BUG-A修正: ログの競合は分離して追記オプションを提供
    for p in targets["logs"]:
        dst_p = _get_log_dst_path(p, dst_dir)
        if dst_p.exists():
            log_conflicts.append(("logs", _get_log_display_name(p)))

    if conflicts:
        print()
        print(f"  [注意] 以下のファイルは移行先に既に存在します:")
        if len(conflicts) == 1:
            name = conflicts[0][1]
            print(f"       - {name}")
            answer = input(f"\n上書きしますか？ (y/N): ").strip().lower()
            if answer not in ("y", "yes"):
                skip_files.add(name)
        else:
            for _, name in conflicts:
                print(f"       - {name}")
            print()
            print("  (a) すべて上書き")
            print("  (s) 個別に選択")
            print("  (n) すべてスキップ")
            answer = input("\n選択してください (a/s/N): ").strip().lower()
            if answer == "a":
                pass  # 全部上書き
            elif answer == "s":
                for _, name in conflicts:
                    ans = input(f"  {name} を上書き？ (y/N): ").strip().lower()
                    if ans not in ("y", "yes"):
                        skip_files.add(name)
            else:
                for _, name in conflicts:
                    skip_files.add(name)

    # BUG-A修正: ログ競合は追記/上書き/スキップの3択
    if log_conflicts:
        print()
        print("  [注意] 以下のログファイルは移行先に既に存在します:")
        for _, name in log_conflicts:
            print(f"       - {name}")
        print()
        print("  (m) 追記マージ（既存ログを保持して末尾に追加）")
        print("      ※ 同じ移行を2回実行するとエントリが重複します")
        print("  (o) 上書き（既存ログは消えます）")
        if len(log_conflicts) >= 2:
            print("  (s) 個別に選択")
        print("  (n) スキップ")
        prompt = "(M/o/s/n)" if len(log_conflicts) >= 2 else "(M/o/n)"
        answer = input(f"\n選択してください {prompt}: ").strip().lower()
        if answer == "o":
            pass  # 上書き
        elif answer == "n":
            for _, name in log_conflicts:
                skip_files.add(name)
        elif answer == "s" and len(log_conflicts) >= 2:
            for _, name in log_conflicts:
                print(f"\n  {name}:")
                print("    (m) 追記マージ  (o) 上書き  (n) スキップ")
                ans = input("    選択 (M/o/n): ").strip().lower()
                if ans == "o":
                    pass
                elif ans == "n":
                    skip_files.add(name)
                else:
                    merge_log_files.add(name)
        else:
            # デフォルトは追記マージ（安全側）
            for _, name in log_conflicts:
                merge_log_files.add(name)

    # ── コピー前の健全性チェック ──
    pre_warnings: list[str] = []

    # (1) 移行元: 設定ファイルのJSON整合性
    if targets["settings"] and SETTINGS_FILE_NAME not in skip_files:
        src_settings = targets["settings"][0]
        try:
            data = json.loads(src_settings.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                pre_warnings.append(
                    f"移行元 {SETTINGS_FILE_NAME}: JSON がdict型ではありません"
                )
        except json.JSONDecodeError as e:
            pre_warnings.append(
                f"移行元 {SETTINGS_FILE_NAME}: JSON が壊れています ({e})"
            )
        except OSError as e:
            pre_warnings.append(
                f"移行元 {SETTINGS_FILE_NAME}: 読み込み失敗 ({e})"
            )

    # (2) 移行元: プリセットのJSON整合性
    for p in targets["presets"]:
        display = f"{PRESET_DIR_NAME}/{p.name}"
        if display in skip_files:
            continue
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                pre_warnings.append(f"移行元 {display}: JSON がdict型ではありません")
        except json.JSONDecodeError as e:
            pre_warnings.append(f"移行元 {display}: JSON が壊れています ({e})")
        except OSError as e:
            pre_warnings.append(f"移行元 {display}: 読み込み失敗 ({e})")

    # (3) 移行元: ペルソナの空ファイル検出
    for p in targets["personas"]:
        display = f"{PERSONAS_DIR_NAME}/{p.name}"
        if display in skip_files:
            continue
        try:
            if p.stat().st_size == 0:
                pre_warnings.append(f"移行元 {display}: 空ファイルです")
        except OSError as e:
            pre_warnings.append(f"移行元 {display}: 状態取得失敗 ({e})")

    # (4) 移行元: ログのJSONL整合性（先頭5行）
    for p in targets["logs"]:
        display = _get_log_display_name(p)
        if display in skip_files:
            continue
        try:
            lines = p.read_text(encoding="utf-8").splitlines()
            non_empty = [ln for ln in lines if ln.strip()]
            if non_empty:
                bad = 0
                for ln in non_empty[:5]:
                    try:
                        json.loads(ln)
                    except json.JSONDecodeError:
                        bad += 1
                if bad > 0:
                    pre_warnings.append(
                        f"移行元 {display}: 先頭{min(5, len(non_empty))}行中"
                        f"{bad}行がJSON不正です"
                    )
        except OSError as e:
            pre_warnings.append(f"移行元 {display}: 読み込み失敗 ({e})")

    # (5) 不可視文字スキャン（ASCII Smuggling / プロンプトインジェクション検出）
    _suspicious_warnings: list[str] = []

    # ペルソナファイル（AIに直接読まれるため最もリスクが高い）
    for p in targets["personas"]:
        display = f"{PERSONAS_DIR_NAME}/{p.name}"
        if display in skip_files:
            continue
        try:
            text = p.read_text(encoding="utf-8-sig")
            hits = _scan_suspicious_chars(text)
            if hits:
                _suspicious_warnings.append(f"  {display}:")
                for h in hits:
                    _suspicious_warnings.append(f"    {h}")
        except OSError:
            pass

    # ログファイル
    for p in targets["logs"]:
        display = _get_log_display_name(p)
        if display in skip_files:
            continue
        try:
            text = p.read_text(encoding="utf-8-sig")
            hits = _scan_suspicious_chars(text)
            if hits:
                _suspicious_warnings.append(f"  {display}:")
                for h in hits:
                    _suspicious_warnings.append(f"    {h}")
        except OSError:
            pass

    # プリセットJSON内の文字列値
    for p in targets["presets"]:
        display = f"{PRESET_DIR_NAME}/{p.name}"
        if display in skip_files:
            continue
        try:
            text = p.read_text(encoding="utf-8")
            hits = _scan_suspicious_chars(text)
            if hits:
                _suspicious_warnings.append(f"  {display}:")
                for h in hits:
                    _suspicious_warnings.append(f"    {h}")
        except OSError:
            pass

    # 設定ファイル（平文のみ。DPAPI暗号化は検査不可）
    if targets["settings"] and SETTINGS_FILE_NAME not in skip_files:
        try:
            text = targets["settings"][0].read_text(encoding="utf-8")
            hits = _scan_suspicious_chars(text)
            if hits:
                _suspicious_warnings.append(f"  {SETTINGS_FILE_NAME}:")
                for h in hits:
                    _suspicious_warnings.append(f"    {h}")
        except OSError:
            pass

    if _suspicious_warnings:
        pre_warnings.append(
            "[セキュリティ] 不可視文字が検出されました（プロンプトインジェクションの可能性）:\n"
            + "\n".join(_suspicious_warnings)
        )

    # (6) 移行先: 書き込み権限テスト
    # BUG-AG修正: write成功→unlink失敗でゴミが残る問題を防止
    _test_file = dst_dir / ".migration_write_test"
    try:
        _test_file.write_text("test", encoding="utf-8")
    except OSError:
        pre_warnings.append("移行先フォルダに書き込み権限がありません")
    else:
        try:
            _test_file.unlink()
        except OSError:
            pass  # 書き込みは成功しているので権限警告は不要

    # (7) 移行先: ディスク空き容量チェック
    try:
        total_src_size = 0
        for category in targets.values():
            for p in category:
                try:
                    total_src_size += p.stat().st_size
                except OSError:
                    pass
        # .venv コピー分のサイズも加算
        if copy_venv_flag and old_venv.is_dir():
            for dirpath, _dirnames, filenames in os.walk(old_venv):
                for fn in filenames:
                    try:
                        total_src_size += os.path.getsize(os.path.join(dirpath, fn))
                    except OSError:
                        pass
        if total_src_size > 0:
            disk_usage = shutil.disk_usage(dst_dir)
            if disk_usage.free < total_src_size * 2:
                free_mb = disk_usage.free / (1024 * 1024)
                need_mb = total_src_size / (1024 * 1024)
                pre_warnings.append(
                    f"ディスク空き容量が少ない可能性があります"
                    f"（空き: {free_mb:.1f}MB / 必要: {need_mb:.1f}MB以上推奨）"
                )
    except OSError:
        pass

    if pre_warnings:
        print()
        print(f"  [検証] コピー前チェックで {len(pre_warnings)} 件の警告:")
        for w in pre_warnings:
            print(f"    ! {w}")
        print()
        answer = input("  警告がありますが続行しますか？ (y/N): ").strip().lower()
        if answer not in ("y", "yes"):
            print("\nキャンセルされました。")
            return 1

    # 最終確認
    final_count = sum(len(v) for v in targets.values()) - len(skip_files)
    # legacy 追記分をカウントに含める（slot1 がスキップされない場合）
    if legacy_log_to_append is not None and SLOT1_LOG_FILE not in skip_files:
        final_count += 1
    venv_label = ""
    if copy_venv_flag:
        final_count += 1
        venv_label = "（+ .venv コピー）"
    if final_count <= 0:
        print("\n移行するファイルがありません。")
        return 1
    print()
    answer = input(f"{final_count} 件を移行します{venv_label}。実行しますか？ (y/N): ").strip().lower()
    if answer not in ("y", "yes"):
        print("\nキャンセルされました。")
        return 1

    # 移行実行
    print()
    copied = 0
    errors: list[str] = []
    path_rewrite_count = 0

    # BUG-L修正: resolve() をループ外でキャッシュ
    need_root_rewrite = False
    try:
        need_root_rewrite = src_dir.resolve() != source_root.resolve()
    except OSError:
        pass

    # 設定ファイル
    if targets["settings"] and SETTINGS_FILE_NAME not in skip_files:
        src_path = targets["settings"][0]
        dst_path = dst_dir / SETTINGS_FILE_NAME
        overwrite = dst_path.exists()
        try:
            copy_file_safe(src_path, dst_path, overwrite=overwrite)
            print(f"  [OK] {SETTINGS_FILE_NAME}")
            copied += 1
            # BUG1修正: 設定内の絶対パスを書き換え
            rewritten = _rewrite_paths_in_settings(dst_path, src_dir, dst_dir)
            # BUG-D修正: src_dir と source_root が異なる場合、ルート基準のパスも書き換え
            if need_root_rewrite:
                rewritten |= _rewrite_paths_in_settings(
                    dst_path, source_root, dst_dir.parent
                )
            if rewritten:
                path_rewrite_count += 1
                print(f"       → パスを新しい場所に更新しました")
        except OSError as e:
            errors.append(f"{SETTINGS_FILE_NAME}: {e}")
            print(f"  [NG] {SETTINGS_FILE_NAME}: {e}")

    # プリセット
    for p in targets["presets"]:
        display = f"{PRESET_DIR_NAME}/{p.name}"
        if display in skip_files:
            continue
        dst_path = dst_dir / PRESET_DIR_NAME / p.name
        overwrite = dst_path.exists()
        try:
            copy_file_safe(p, dst_path, overwrite=overwrite)
            print(f"  [OK] {display}")
            copied += 1
            # BUG1修正: プリセット内の絶対パスも書き換え
            rewritten = _rewrite_paths_in_preset(dst_path, src_dir, dst_dir)
            # BUG-D修正: ルート基準のパスも書き換え
            if need_root_rewrite:
                rewritten |= _rewrite_paths_in_preset(
                    dst_path, source_root, dst_dir.parent
                )
            if rewritten:
                path_rewrite_count += 1
                print(f"       → パスを新しい場所に更新しました")
        except OSError as e:
            errors.append(f"{display}: {e}")
            print(f"  [NG] {display}: {e}")

    # キャラ設定
    for p in targets["personas"]:
        display = f"{PERSONAS_DIR_NAME}/{p.name}"
        if display in skip_files:
            continue
        dst_path = dst_dir / PERSONAS_DIR_NAME / p.name
        overwrite = dst_path.exists()
        try:
            copy_file_safe(p, dst_path, overwrite=overwrite)
            print(f"  [OK] {display}")
            copied += 1
        except OSError as e:
            errors.append(f"{display}: {e}")
            print(f"  [NG] {display}: {e}")

    # ログ（BUG-A修正: 追記マージ対応）
    legacy_appended = False
    for p in targets["logs"]:
        display = _get_log_display_name(p)
        if display in skip_files:
            continue
        dst_path = _get_log_dst_path(p, dst_dir)
        try:
            # BUG-Z修正: append_log_file の戻り値で空ファイルを検出
            # BUG-AB修正: 空ファイルスキップ時は copied に加算しない
            if display in merge_log_files:
                if append_log_file(p, dst_path):
                    print(f"  [OK] {display}（追記マージ）")
                    copied += 1
                else:
                    print(f"  [--] {display}（空のためスキップ）")
            else:
                overwrite = dst_path.exists()
                copy_file_safe(p, dst_path, overwrite=overwrite)
                print(f"  [OK] {display}")
                copied += 1
            # BUG-V修正: slot1 コピー/マージ後に legacy の内容を追記
            if p.name == SLOT1_LOG_FILE and legacy_log_to_append is not None:
                try:
                    if append_log_file(legacy_log_to_append, dst_path):
                        print(f"       + {LEGACY_LOG_FILE} の内容を追記しました")
                        legacy_appended = True
                    else:
                        print(f"       - {LEGACY_LOG_FILE} は空のためスキップ")
                        legacy_appended = True  # 空でもスキップ扱い（処理済み）
                except OSError as e:
                    errors.append(f"{LEGACY_LOG_FILE} 追記: {e}")
                    print(f"  [NG] {LEGACY_LOG_FILE} 追記: {e}")
        except OSError as e:
            errors.append(f"{display}: {e}")
            print(f"  [NG] {display}: {e}")

    # BUG-AK修正: slot1 がスキップされた場合、legacy を単独で移行する
    if legacy_log_to_append is not None and not legacy_appended:
        dst_path = dst_dir / SLOT1_LOG_FILE
        display = f"{LEGACY_LOG_FILE} → {SLOT1_LOG_FILE}"
        try:
            if dst_path.exists():
                # 既存ログに追記
                if append_log_file(legacy_log_to_append, dst_path):
                    print(f"  [OK] {display}（追記マージ）")
                    copied += 1
                else:
                    print(f"  [--] {display}（空のためスキップ）")
            else:
                copy_file_safe(legacy_log_to_append, dst_path, overwrite=False)
                print(f"  [OK] {display}")
                copied += 1
        except OSError as e:
            errors.append(f"{LEGACY_LOG_FILE}: {e}")
            print(f"  [NG] {LEGACY_LOG_FILE}: {e}")

    # セットアップ環境(.venv)のコピー
    venv_copied = False
    if copy_venv_flag:
        if copy_venv(old_venv, new_venv):
            copied += 1
            venv_copied = True
        else:
            errors.append(".venv のコピーに失敗")

    # BUG1補足: DPAPI暗号化設定の場合はパス書き換え不可を案内
    # BUG-T修正: dpapi_warned で事前警告済みの場合のみパス案内を表示
    # （事前にスキップされた場合は targets["settings"] が空なのでここに到達しない）
    if dpapi_warned and targets["settings"] and SETTINGS_FILE_NAME not in skip_files:
        old_prefix = _normalize_path_str(str(src_dir))
        new_prefix = _normalize_path_str(str(dst_dir))
        if old_prefix.lower() != new_prefix.lower():
            print()
            print("  [注意] 設定ファイルはDPAPI暗号化されているため、")
            print("  ファイルパスの自動書き換えができませんでした。")
            print("  移行後に「設定画面を開く.bat」で")
            print("  要約プロンプトのパスを確認・修正してください。")

    # ── ヘルパー: パスキー検証 ──
    def _verify_path_keys(
        file_path: Path,
        keys: tuple[str, ...],
        label: str,
        warn_list: list[str],
        *,
        unwrap_plaintext: bool = False,
        sections: tuple[str, ...] = (),
    ) -> None:
        """*file_path* のJSON内の *keys* に該当するパスが実在するか検証する。

        *sections* が指定された場合、トップレベルに加えて各セクション内も検証する。
        """
        try:
            data = json.loads(file_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return  # 読めない場合は別のチェックで報告済み
        if not isinstance(data, dict):
            return
        if unwrap_plaintext and data.get("format") == "plaintext":
            inner = data.get("data")
            if isinstance(inner, dict):
                data = inner
            else:
                return
        # 検証対象: トップレベル + 指定セクション
        targets_to_check: list[tuple[str, dict]] = [(label, data)]
        for sec_name in sections:
            sec = data.get(sec_name)
            if isinstance(sec, dict):
                targets_to_check.append((f"{label}/{sec_name}", sec))
        for target_label, target_dict in targets_to_check:
            for key in keys:
                val = target_dict.get(key)
                if isinstance(val, str) and val.strip():
                    if not Path(val).is_file():
                        warn_list.append(
                            f"{target_label}: {key} のパス '{val}' が存在しません"
                        )

    # ── 移行後の健全性チェック ──
    warnings: list[str] = []

    # 設定ファイル: JSONとして読めるか + バージョン整合
    if targets["settings"] and SETTINGS_FILE_NAME not in skip_files:
        dst_settings = dst_dir / SETTINGS_FILE_NAME
        if dst_settings.is_file():
            try:
                data = json.loads(dst_settings.read_text(encoding="utf-8"))
                if not isinstance(data, dict):
                    warnings.append(f"{SETTINGS_FILE_NAME}: JSONがdict型ではありません")
            except (json.JSONDecodeError, OSError) as e:
                warnings.append(f"{SETTINGS_FILE_NAME}: 移行後のJSON読み込みに失敗 ({e})")
        else:
            warnings.append(f"{SETTINGS_FILE_NAME}: コピー先にファイルが存在しません")

    # プリセット: JSONとして読めるか
    for p in targets["presets"]:
        display = f"{PRESET_DIR_NAME}/{p.name}"
        if display in skip_files:
            continue
        dst_p = dst_dir / PRESET_DIR_NAME / p.name
        if dst_p.is_file():
            try:
                data = json.loads(dst_p.read_text(encoding="utf-8"))
                if not isinstance(data, dict):
                    warnings.append(f"{display}: JSONがdict型ではありません")
            except (json.JSONDecodeError, OSError) as e:
                warnings.append(f"{display}: 移行後のJSON読み込みに失敗 ({e})")
        else:
            warnings.append(f"{display}: コピー先にファイルが存在しません")

    # キャラ設定: ファイルが存在し空でないか
    for p in targets["personas"]:
        display = f"{PERSONAS_DIR_NAME}/{p.name}"
        if display in skip_files:
            continue
        dst_p = dst_dir / PERSONAS_DIR_NAME / p.name
        # BUG-AI修正: stat() の OSError を捕捉
        try:
            if dst_p.is_file():
                if dst_p.stat().st_size == 0:
                    warnings.append(f"{display}: ファイルが空です")
            else:
                warnings.append(f"{display}: コピー先にファイルが存在しません")
        except OSError as e:
            warnings.append(f"{display}: 状態取得失敗 ({e})")

    # ログ: ファイルが存在し空でないか + 各行がJSONとして有効か（先頭5行のみ）
    for p in targets["logs"]:
        display = _get_log_display_name(p)
        if display in skip_files:
            continue
        dst_p = _get_log_dst_path(p, dst_dir)
        # BUG-AJ修正: stat() の OSError を捕捉
        try:
            is_file = dst_p.is_file()
            file_size = dst_p.stat().st_size if is_file else -1
        except OSError as e:
            warnings.append(f"{dst_p.name}: 状態取得失敗 ({e})")
            continue
        if is_file:
            if file_size == 0:
                warnings.append(f"{dst_p.name}: ログファイルが空です")
            else:
                try:
                    lines = dst_p.read_text(encoding="utf-8").splitlines()
                    non_empty = [ln for ln in lines if ln.strip()]
                    bad_lines = 0
                    for ln in non_empty[:5]:
                        try:
                            json.loads(ln)
                        except json.JSONDecodeError:
                            bad_lines += 1
                    if bad_lines > 0:
                        warnings.append(
                            f"{dst_p.name}: 先頭{min(5, len(non_empty))}行中"
                            f"{bad_lines}行がJSON不正です"
                        )
                except OSError as e:
                    warnings.append(f"{dst_p.name}: 読み込み失敗 ({e})")
        # ログは追記マージで新規作成される場合もあるため、存在しない場合はチェックしない

    # パス書き換え結果の検証: 書き換えたパスが実在するか
    # DPAPI暗号化で書き換え不可だった場合も検証対象にする
    if path_rewrite_count > 0 or dpapi_warned:
        # 設定ファイル内のパス
        _verify_path_keys(
            dst_dir / SETTINGS_FILE_NAME, _PATH_KEYS_IN_SETTINGS,
            "設定", warnings, unwrap_plaintext=True,
        )
        # プリセット内のパス（トップレベル + character セクション）
        for p in targets["presets"]:
            display = f"{PRESET_DIR_NAME}/{p.name}"
            if display in skip_files:
                continue
            dst_p = dst_dir / PRESET_DIR_NAME / p.name
            _verify_path_keys(
                dst_p, _PATH_KEYS_IN_PRESET, display, warnings,
                sections=("character",),
            )

    # 結果表示
    print()
    print("=" * 50)
    if errors:
        print(f"  移行完了: {copied} 件成功 / {len(errors)} 件失敗")
        print()
        for err in errors:
            print(f"  [NG] {err}")
    else:
        print(f"  移行完了: {copied} 件すべて成功")
    if path_rewrite_count > 0:
        print(f"  パス更新: {path_rewrite_count} 件")
    if warnings:
        print()
        print(f"  [検証] {len(warnings)} 件の警告:")
        for w in warnings:
            print(f"    ! {w}")
    else:
        print("  [検証] すべてのファイルの健全性を確認しました")
    print("=" * 50)
    if venv_copied:
        print()
        print("  .venv をコピーしました。依存パッケージの差分確認を行います...")
        print("  （バージョンアップ・データ移行.bat が自動で実行します）")

    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(run_migration_cli())
