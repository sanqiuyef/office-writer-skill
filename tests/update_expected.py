#!/usr/bin/env python3
"""Regenerate tests/expected_outputs from the current CLI behavior.

Run from the skill root after an intentional behavior change:
    python -X utf8 tests/update_expected.py
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

from cases import CASES

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
EXPECTED = Path(__file__).resolve().parent / "expected_outputs"


def run_case(case: dict) -> dict:
    import subprocess

    result = subprocess.run(
        [sys.executable, "-X", "utf8", str(SCRIPTS / case["script"]), *case["args"]],
        cwd=SCRIPTS,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=os.environ.copy(),
    )
    if result.returncode != case["exit_code"]:
        raise SystemExit(
            f"{case['name']} exit {result.returncode} != {case['exit_code']}\n"
            f"stdout={result.stdout}\nstderr={result.stderr}"
        )
    payload = json.loads(result.stdout)
    if not payload.get("ok"):
        raise SystemExit(f"{case['name']} returned ok=false: {payload}")
    return payload["data"]


def main() -> None:
    EXPECTED.mkdir(parents=True, exist_ok=True)
    for case in CASES:
        data = run_case(case)
        path = EXPECTED / f"{case['name']}.json"
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"wrote {path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
