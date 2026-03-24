"""仮想環境を作成し、スピナーアニメーションで進捗を表示する。"""

import subprocess
import sys
import threading
import time


def main() -> int:
    done = False
    result_code = [0]

    def _create() -> None:
        nonlocal done
        r = subprocess.run([sys.executable, "-m", "venv", ".venv"])
        result_code[0] = r.returncode
        done = True

    t = threading.Thread(target=_create)
    t.start()

    spinner = ["/", "-", "\\", "|"]
    i = 0
    while not done:
        print(f"  {spinner[i % 4]} 仮想環境を作成中...", end="\r", flush=True)
        time.sleep(0.2)
        i += 1

    if result_code[0] == 0:
        print("  [OK] .venv を作成しました        ")
    else:
        print("  [ERROR] .venv の作成に失敗しました")
    return result_code[0]


if __name__ == "__main__":
    raise SystemExit(main())
