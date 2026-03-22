"""設定画面ランチャー — 多重起動防止 + クラッシュ時にログを残す"""
import ctypes
import platform
import sys
import traceback
from pathlib import Path

log_path = Path(__file__).parent / "config_error.log"

# ── 多重起動防止（Windows Named Mutex） ──────────────────
# カーネルレベルでアトミックに排他制御するため、
# .bat 連打やPython起動のタイムラグによるレースコンディションが起きない。

_mutex_handle = None
_MUTEX_NAME = "OCA_Config_SingleInstance"
_ERROR_ALREADY_EXISTS = 183


def _acquire_mutex() -> bool:
    """名前付きミューテックスで多重起動を防止する。取得できたら True。"""
    global _mutex_handle
    if platform.system() != "Windows":
        return True
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]
    kernel32.CloseHandle.restype = ctypes.c_bool
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    kernel32.ReleaseMutex.restype = ctypes.c_bool
    kernel32.ReleaseMutex.argtypes = [ctypes.c_void_p]
    _mutex_handle = kernel32.CreateMutexW(None, ctypes.c_bool(True), _MUTEX_NAME)
    if ctypes.get_last_error() == _ERROR_ALREADY_EXISTS:
        kernel32.CloseHandle(_mutex_handle)
        _mutex_handle = None
        return False
    return True


def _release_mutex() -> None:
    global _mutex_handle
    if _mutex_handle is not None:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.ReleaseMutex(_mutex_handle)
        kernel32.CloseHandle(_mutex_handle)
        _mutex_handle = None


if not _acquire_mutex():
    sys.exit(0)

# 前回のエラーログをクリア
try:
    log_path.write_text("", encoding="utf-8")
except OSError:
    pass

try:
    import config
    config.main()
except Exception:
    msg = traceback.format_exc()
    log_path.write_text(msg, encoding="utf-8")
    # pythonw でも見えるようにメッセージボックスで通知
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("設定画面エラー", f"設定画面の起動中にエラーが発生しました。\n\n{msg[:500]}")
        root.destroy()
    except Exception:
        pass
    sys.exit(1)
finally:
    _release_mutex()
