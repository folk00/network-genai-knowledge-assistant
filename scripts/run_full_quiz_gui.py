from __future__ import annotations

import runpy
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP_DIR = ROOT / "apps" / "full_quiz_gui"
APP_FILE = APP_DIR / "ans_c01_quiz_gui_v2_counterfix2.py"


def main() -> int:
    if not APP_FILE.exists():
        print(f"Missing app file: {APP_FILE}", file=sys.stderr)
        return 1
    sys.path.insert(0, str(APP_DIR))
    runpy.run_path(str(APP_FILE), run_name="__main__")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

