"""Lexical trigger baseline must stay aligned with SKILL.md and trigger_cases.json."""

from __future__ import annotations

import json
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CASES = Path(__file__).resolve().parent / "trigger_cases.json"
SCRIPTS = ROOT / "scripts"


def description() -> str:
    text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    front = text.split("---", 2)[1]
    match = re.search(r"^description:\s*(.+)$", front, re.MULTILINE)
    if not match:
        raise AssertionError("SKILL.md description missing")
    return match.group(1)


def trigger_terms() -> list[str]:
    desc = description()
    match = re.search(r"触发：(.+)", desc)
    if not match:
        raise AssertionError("description 缺少 触发： 词表")
    raw = match.group(1).rstrip("。；; ")
    return [term.strip() for term in re.split(r"[、，,]", raw) if term.strip()]


class TriggerBaselineTest(unittest.TestCase):
    def test_lexical_baseline_accuracy(self) -> None:
        data = json.loads(CASES.read_text(encoding="utf-8"))
        cases = data["cases"]
        terms = trigger_terms()
        self.assertGreaterEqual(len(terms), 12)
        misses = []
        for case in cases:
            predicted = any(term in case["prompt"] for term in terms)
            if predicted != case["should_trigger"]:
                misses.append((case["id"], case["prompt"], predicted, case["should_trigger"]))
        self.assertEqual(misses, [], f"触发基线误判: {misses}")

    def test_positive_cases_route_to_existing_scripts(self) -> None:
        cases = json.loads(CASES.read_text(encoding="utf-8"))["cases"]
        positives = [case for case in cases if case["should_trigger"]]
        negatives = [case for case in cases if not case["should_trigger"]]
        self.assertGreaterEqual(len(positives), 12)
        self.assertGreaterEqual(len(negatives), 8)
        for case in positives:
            with self.subTest(case=case["id"]):
                script = case.get("expected_script", "")
                self.assertTrue(script)
                self.assertTrue((SCRIPTS / script).is_file(), script)

    def test_description_budget(self) -> None:
        desc = description()
        self.assertLessEqual(len(desc.encode("utf-8")), 1024)
        self.assertIn("触发：", desc)


if __name__ == "__main__":
    unittest.main()
