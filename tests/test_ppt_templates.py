"""PPT 模板库契约测试：6 套 16:9 模板的结构、版式顺序与 .potx 可套用性。"""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PPT_DIR = ROOT / "templates" / "ppt"
MANIFEST = ROOT / "assets" / "manifest.json"
SCRIPTS = ROOT / "scripts"

SLUGS = [
    "business-report",
    "consulting",
    "academic-defense",
    "academic-seminar",
    "government-report",
    "government-briefing",
]
STANDARD_LAYOUTS = [
    "Title Slide",
    "Title and Content",
    "Section Header",
    "Two Content",
    "Comparison",
    "Title Only",
    "Blank",
    "Content with Caption",
    "Picture with Caption",
    "Title and Vertical Text",
    "Vertical Title and Text",
]
SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5


class PptTemplateTest(unittest.TestCase):
    def test_all_themes_have_both_forms(self) -> None:
        for slug in SLUGS:
            for suffix in (".pptx", ".potx"):
                path = PPT_DIR / f"{slug}{suffix}"
                with self.subTest(file=path.name):
                    self.assertTrue(path.is_file(), path.name)
                    self.assertTrue(zipfile.is_zipfile(path), path.name)

    def test_demo_decks_are_16x9_with_standard_layouts(self) -> None:
        from pptx import Presentation

        for slug in SLUGS:
            with self.subTest(slug=slug):
                prs = Presentation(str(PPT_DIR / f"{slug}.pptx"))
                self.assertAlmostEqual(prs.slide_width / 914400, SLIDE_WIDTH, places=2)
                self.assertAlmostEqual(prs.slide_height / 914400, SLIDE_HEIGHT, places=2)
                self.assertEqual(len(prs.slides), 8)
                self.assertEqual([layout.name for layout in prs.slide_layouts],
                                 STANDARD_LAYOUTS)

    def test_demo_slides_have_titles(self) -> None:
        from pptx import Presentation

        for slug in SLUGS:
            prs = Presentation(str(PPT_DIR / f"{slug}.pptx"))
            for index, slide in enumerate(prs.slides, start=1):
                with self.subTest(slug=slug, slide=index):
                    title = slide.shapes.title
                    self.assertIsNotNone(title, f"{slug} 第 {index} 页无标题占位符")
                    self.assertTrue(title.text.strip(), f"{slug} 第 {index} 页标题为空")

    def test_potx_is_zero_page_template(self) -> None:
        from pptx import Presentation

        for slug in SLUGS:
            path = PPT_DIR / f"{slug}.potx"
            with self.subTest(slug=slug):
                with zipfile.ZipFile(path) as archive:
                    content_types = archive.read("[Content_Types].xml").decode("utf-8")
                self.assertIn(
                    "presentationml.template.main+xml", content_types, slug
                )
                # python-pptx 不直接认 template.main，借用 skill 的加载器转换
                sys.path.insert(0, str(SCRIPTS))
                from office_engine.ppt_writer import _load_presentation
                prs = _load_presentation(str(path))
                self.assertEqual(len(prs.slides), 0)
                self.assertAlmostEqual(prs.slide_width / 914400, SLIDE_WIDTH, places=2)

    def test_manifest_matches_files(self) -> None:
        data = json.loads(MANIFEST.read_text(encoding="utf-8"))
        entries = {item["slug"]: item for item in data["powerpoint_templates"]}
        self.assertEqual(sorted(entries), sorted(SLUGS))
        for slug, entry in entries.items():
            with self.subTest(slug=slug):
                self.assertTrue((ROOT / entry["demo"]).is_file())
                self.assertTrue((ROOT / entry["layout_template"]).is_file())
                self.assertEqual(entry["slide_size"], "13.333x7.5")
                self.assertEqual(entry["slide_count"], 8)

    def test_create_from_potx(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ppt_template_") as tmp:
            output = Path(tmp) / "deck.pptx"
            params = {"slides": [
                {"layout": "title", "title": "模板套用测试", "subtitle": "副标题"},
                {"layout": "content", "title": "内容页",
                 "body": [{"type": "paragraph", "text": "正文"}]},
            ]}
            result = subprocess.run(
                [sys.executable, "-X", "utf8", str(SCRIPTS / "create_powerpoint.py"),
                 "--output-path", str(output),
                 "--template", str(PPT_DIR / "academic-defense.potx"),
                 "--params", json.dumps(params, ensure_ascii=False)],
                capture_output=True, text=True, encoding="utf-8",
                errors="backslashreplace", check=False,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            payload = json.loads(result.stdout)
            self.assertTrue(payload["ok"], payload)
            self.assertEqual(payload["data"]["slides_count"], 2)

            from pptx import Presentation
            prs = Presentation(str(output))
            self.assertAlmostEqual(prs.slide_width / 914400, SLIDE_WIDTH, places=2)
            self.assertAlmostEqual(prs.slide_height / 914400, SLIDE_HEIGHT, places=2)
            self.assertEqual([slide.slide_layout.name for slide in prs.slides],
                             ["Title Slide", "Title and Content"])


if __name__ == "__main__":
    unittest.main()
