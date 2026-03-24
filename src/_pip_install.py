"""requirements.txt のパッケージを1つずつインストールし、全体の進捗バーを表示する。"""

import pathlib
import subprocess
import sys


def main() -> int:
    req = pathlib.Path("requirements.txt")
    if not req.is_file():
        print("[ERROR] requirements.txt が見つかりません")
        return 1

    pkgs = [
        line.strip()
        for line in req.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.startswith("#")
    ]
    if not pkgs:
        print("[OK] インストールするパッケージはありません")
        return 0

    total = len(pkgs)
    failed: list[str] = []

    for i, pkg in enumerate(pkgs):
        done = i + 1
        pct = 100 * done // total
        filled = 20 * done // total
        bar = "#" * filled + "." * (20 - filled)
        print(f"  [{bar}] {pct:3d}%  ({done}/{total})", end="\r", flush=True)
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", pkg, "-qq"],
            capture_output=True,
        )
        if result.returncode != 0:
            failed.append(pkg)

    bar_done = "#" * 20
    print(f"  [{bar_done}] 100%  ({total}/{total})  ")

    if failed:
        print(f"[WARN] 一部のパッケージのインストールに失敗しました: {', '.join(failed)}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
