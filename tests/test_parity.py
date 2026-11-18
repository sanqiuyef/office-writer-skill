"""Compare the vendored engine with pre-recorded golden JSON.

The goldens under ``tests/expected_outputs/parity_*.json`` were captured from
mcp-office-writer 0.6.0 when this skill was packaged.  Re-run
``python -X utf8 tests/update_expected.py`` after an intentional engine change;
no upstream ``mcp`` package is needed at test time.
"""

from __future__ import annotations

import json
import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
EXPECTED = Path(__file__).resolve().parent / "expected_outputs"


class ParityTest(unittest.TestCase):
    def test_vendored_engine_matches_golden(self) -> None:
        sys.path.insert(0, str(SCRIPTS))
        from office_engine import checker as vendored_checker
        from office_engine import ppt_writer as vendored_ppt
        from office_engine import reader as vendored_reader

        previous_cwd = os.getcwd()
        os.chdir(ROOT)
        try:
            actual = {
                "parity_list_color_schemes": vendored_ppt.list_color_schemes(),
                "parity_read_txt": vendored_reader.read_txt_document("tests/fixtures/sample.txt"),
                "parity_validate_word": vendored_checker.validate_word_document(
                    "tests/fixtures/sample.docx"
                ).to_dict(),
            }
        finally:
            os.chdir(previous_cwd)

        for name, value in actual.items():
            with self.subTest(case=name):
                golden = json.loads((EXPECTED / f"{name}.json").read_text(encoding="utf-8"))
                self.assertEqual(value, golden)

    def test_frontmatter_smoke(self) -> None:
        text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("office-writer", text)


if __name__ == "__main__":
    unittest.main()
