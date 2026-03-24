"""requirements.txt のパッケージを1つずつインストールし、全体の進捗バーを表示する。"""

import pathlib
import re
import subprocess
import sys


def _find_needed(pkgs: list[str]) -> list[str]:
    """pip install --dry-run -r requirements.txt を一括実行し、
    インストールが必要なパッケージ名を返す。"""
    req = pathlib.Path("requirements.txt")
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "--dry-run", "-r", str(req)],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        # dry-run 自体が失敗 → 全パッケージをインストール対象とする
        return pkgs

    stdout_lower = result.stdout.lower()
    if "would install" not in stdout_lower:
        return []

    # "Would install numpy-1.26.4 requests-2.32.5" の行からパッケージ名を抽出
    needed_names: set[str] = set()
    for line in result.stdout.splitlines():
        if line.strip().lower().startswith("would install"):
            # "Would install pkg1-1.0 pkg2-2.0" → ["pkg1", "pkg2"]
            for token in line.split()[2:]:
                # "numpy-1.26.4" → "numpy"
                # "nvidia-cuda-runtime-cu12-12.1.105" → "nvidia-cuda-runtime-cu12"
                # 末尾から "-バージョン番号" 部分を除去する
                m = re.match(r"^(.+?)-\d+[\d.]*(?:\.(?:post|dev|a|b|rc)\d*)*$", token)
                if m:
                    needed_names.add(m.group(1).lower())
                else:
                    needed_names.add(token.lower())

    # requirements.txt のパッケージ名と照合
    needed: list[str] = []
    for pkg in pkgs:
        # "requests>=2.28,<3" → "requests"
        name = pkg.split(">=")[0].split("<=")[0].split("==")[0].split("<")[0].split(">")[0].split("!=")[0].split("[")[0].strip().lower()
        if name in needed_names:
            needed.append(pkg)

    return needed if needed else pkgs  # 照合に失敗したら安全側に倒す


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

    # インストールまたは更新が必要なパッケージだけ抽出（一括チェック）
    need_install = _find_needed(pkgs)

    if not need_install:
        print(f"[OK] 全{len(pkgs)}件のパッケージは導入済みです")
        return 0

    already = len(pkgs) - len(need_install)
    # パッケージ名のみ抽出して表示（バージョン指定を除去）
    names = [re.split(r"[>=<!\[]", pkg)[0] for pkg in need_install]
    if already > 0:
        print(f"  {len(need_install)}件のパッケージをインストールします ({already}件は導入済み)")
    else:
        print(f"  {len(need_install)}件のパッケージをインストールします")
    print(f"  対象: {', '.join(names)}")

    total = len(need_install)
    failed: list[str] = []

    for i, pkg in enumerate(need_install):
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", pkg, "-qq"],
            capture_output=True,
        )
        if result.returncode != 0:
            failed.append(pkg)
        # インストール完了後にバーを更新
        done = i + 1
        pct = 100 * done // total
        filled = 20 * done // total
        bar = "#" * filled + "." * (20 - filled)
        print(f"  [{bar}] {pct:3d}%  ({done}/{total})", end="\r", flush=True)

    print()  # 改行

    if failed:
        print(f"[WARN] 一部のパッケージのインストールに失敗しました: {', '.join(failed)}")
        return 1

    print(f"[OK] {len(need_install)}件のパッケージをインストールしました")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
