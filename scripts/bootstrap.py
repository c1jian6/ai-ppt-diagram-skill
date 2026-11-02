#!/usr/bin/env python
"""Check (and optionally install) what the skill needs on this machine.

    python scripts/bootstrap.py --check
    python scripts/bootstrap.py --install     # best effort, idempotent
    python scripts/bootstrap.py --json        # machine-readable report
"""

from __future__ import annotations

import argparse
import json
import platform
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

SYSTEM = platform.system().lower()

DOT_HINTS = {
    "windows": "winget install --id Graphviz.Graphviz -e",
    "darwin": "brew install graphviz",
    "linux": "sudo apt install graphviz   # or: sudo dnf install graphviz",
}


def _pip_has(module: str) -> bool:
    try:
        __import__(module)
        return True
    except Exception:
        return False


def _dot_path() -> str | None:
    try:
        from dot2pptx import gv

        return gv.find_dot()
    except Exception:
        return shutil.which("dot")


def _font() -> tuple[bool, str]:
    try:
        from dot2pptx import fonts

        return fonts.is_cjk_capable(), fonts.describe()
    except Exception as exc:  # pragma: no cover - defensive
        return False, f"unavailable ({exc})"


def collect() -> dict:
    dot = _dot_path()
    font_ok, font_desc = _font()
    return {
        "python": {"ok": sys.version_info >= (3, 9), "value": platform.python_version()},
        "graphviz": {
            "ok": bool(dot),
            "value": dot,
            "optional": False,
            "hint": DOT_HINTS.get(SYSTEM, DOT_HINTS["linux"]),
        },
        "python-pptx": {"ok": _pip_has("pptx"), "optional": False,
                        "hint": f'"{sys.executable}" -m pip install python-pptx'},
        "pillow": {"ok": _pip_has("PIL"), "optional": False,
                   "hint": f'"{sys.executable}" -m pip install pillow'},
        "cjk-font": {"ok": font_ok, "value": font_desc, "optional": True,
                     "hint": "install a CJK font, or set DOT2PPTX_FONT / DOT2PPTX_FONT_FILE"},
    }


def print_report(report: dict) -> None:
    for name, item in report.items():
        mark = "ok  " if item["ok"] else ("warn" if item.get("optional") else "MISS")
        value = item.get("value") or ""
        print(f"[{mark}] {name:<12} {value}")
    missing = [name for name, item in report.items()
               if not item["ok"] and not item.get("optional")]
    if missing:
        print("\nMissing required pieces:")
        for name in missing:
            print(f"  {name}: {report[name]['hint']}")
        print("\nstatus: not ready")
    else:
        print("\nstatus: ready")


def install(report: dict) -> int:
    changed = False
    if not report["python-pptx"]["ok"] or not report["pillow"]["ok"]:
        print("[..] installing python packages")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet",
             "python-pptx>=0.6.21", "Pillow>=9.0"],
            check=False,
        )
        changed = True
    if not report["graphviz"]["ok"]:
        hint = report["graphviz"]["hint"]
        print(f"[..] Graphviz is a system package, run it yourself:\n     {hint}")
    return 0 if changed or report["graphviz"]["ok"] else 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="bootstrap.py", description=__doc__)
    parser.add_argument("--check", action="store_true", help="report only (default)")
    parser.add_argument("--install", action="store_true", help="try to install missing pieces")
    parser.add_argument("--json", action="store_true", help="print the report as JSON")
    args = parser.parse_args(argv)

    report = collect()
    if args.install:
        install(report)
        report = collect()

    if args.json:
        print(json.dumps(report, indent=2, ensure_ascii=False))
    else:
        print_report(report)

    return 0 if all(item["ok"] or item.get("optional") for item in report.values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())
