"""Regression tests for the C# backend P2 boundaries.

Covers the gaps that were previously documented as limitations:
Excel date/streaming reads, batch merged-cell safety, PPT chart embedded
workbooks, and Word header/footer styles + section breaks.
"""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
DEFAULT_BINARY = ROOT / "tools/office-native/dist/office-native.exe"


def native_binary() -> Path | None:
    override = os.environ.get("OFFICE_NATIVE_BIN")
    if override:
        candidate = Path(override)
        return candidate if candidate.is_file() else None
    return DEFAULT_BINARY if DEFAULT_BINARY.is_file() else None


HAS_NATIVE = native_binary() is not None
REQUIRES_NATIVE = unittest.skipUnless(
    HAS_NATIVE, "需要 office-native 后端（tools/office-native/dist）"
)


def run_script(script: str, args: list[str]) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env["OFFICE_NATIVE_ENGINE"] = "csharp"
    return subprocess.run(
        [sys.executable, "-X", "utf8", str(SCRIPTS / script), *args],
        cwd=SCRIPTS,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="backslashreplace",
        env=env,
        check=False,
    )


def run_ok(script: str, args: list[str]) -> dict:
    result = run_script(script, args)
    if result.returncode != 0:
        raise AssertionError(
            f"{script} exited {result.returncode}: {result.stdout}\n{result.stderr}"
        )
    payload = json.loads(result.stdout)
    if not payload.get("ok"):
        raise AssertionError(payload)
    return payload["data"]


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class NativeP2Test(unittest.TestCase):
    @REQUIRES_NATIVE
    def test_excel_read_dates_and_streaming_limit(self) -> None:
        from datetime import datetime
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory(prefix="office_native_xlsx_read_") as tmp:
            path = Path(tmp) / "dates.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "数据"
            sheet["A1"] = "日期"
            sheet["B1"] = "值"
            sheet["A2"] = datetime(2026, 1, 2, 3, 4, 5)
            sheet["B2"] = 1.5
            for row in range(3, 2002):
                sheet.cell(row=row, column=1, value="row")
                sheet.cell(row=row, column=2, value=row)
            workbook.save(path)

            data = run_ok("read_excel_workbook.py", [
                "--file-path", str(path),
                "--max-rows-per-sheet", "10",
            ])
            result = data["result"]
            self.assertTrue(result["truncated"])
            self.assertEqual(result["total_rows"], 10)
            self.assertIn("2026-01-02 03:04:05", data["markdown"])
            self.assertEqual(data["markdown"].count("| row |"), 8)

    @REQUIRES_NATIVE
    def test_batch_edit_rejects_merged_non_anchor(self) -> None:
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory(prefix="office_native_merge_") as tmp:
            path = Path(tmp) / "merged.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet["A1"] = "A"
            sheet["B1"] = "B"
            sheet["A2"] = "C"
            sheet["B2"] = "D"
            sheet.merge_cells("A1:B2")
            workbook.save(path)
            before = sha256(path)

            result = run_script("batch_edit_office_document.py", [
                "--file-path", str(path),
                "--params", json.dumps({
                    "operations": [
                        {"op": "set", "path": "/Sheet/B2", "props": {"text": "x"}}
                    ]
                }, ensure_ascii=False),
            ])
            self.assertEqual(result.returncode, 3, result.stdout + result.stderr)
            self.assertEqual(sha256(path), before)

    @REQUIRES_NATIVE
    def test_ppt_chart_embeds_editable_workbook(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_native_ppt_chart_") as tmp:
            path = Path(tmp) / "chart.pptx"
            run_ok("create_powerpoint.py", [
                "--output-path", str(path),
                "--params", json.dumps({
                    "slides": [{
                        "layout": "title_only",
                        "title": "图表",
                        "body": [{
                            "type": "chart",
                            "chart": {
                                "chart_type": "column",
                                "categories": ["Q1", "Q2"],
                                "series": [
                                    {"name": "Sales", "values": [10, 20]},
                                    {"name": "Cost", "values": [5, 8]},
                                ],
                                "x": 1, "y": 1.5, "width": 6, "height": 4,
                            },
                        }],
                    }]
                }, ensure_ascii=False),
            ])
            with zipfile.ZipFile(path) as archive:
                names = archive.namelist()
                chart_names = [n for n in names if n.startswith("ppt/slides/charts/chart") and n.endswith(".xml")]
                self.assertTrue(chart_names)
                chart_xml = archive.read(chart_names[0]).decode("utf-8")
                self.assertIn("externalData", chart_xml)
                self.assertIn("strRef", chart_xml)
                self.assertIn("numRef", chart_xml)
                embedded = [n for n in names if "/embeddings/" in n and n.endswith(".xlsx")]
                self.assertTrue(embedded, names)

    @REQUIRES_NATIVE
    def test_word_header_footer_styles_and_section_break(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_native_word_style_") as tmp:
            path = Path(tmp) / "styles.docx"
            run_ok("create_word_document.py", [
                "--output-path", str(path),
                "--params", json.dumps({
                    "content": {
                        "sections": [
                            {"type": "heading", "text": "第一节", "level": 1},
                            {"type": "paragraph", "text": "正文一"},
                            {"type": "section_break"},
                            {"type": "heading", "text": "第二节", "level": 1},
                            {"type": "paragraph", "text": "正文二"},
                        ]
                    },
                    "header_footer": {
                        "header_text": "页眉",
                        "footer_text": "第{page}页",
                    },
                }, ensure_ascii=False),
            ])
            with zipfile.ZipFile(path) as archive:
                styles = archive.read("word/styles.xml").decode("utf-8")
                self.assertIn('w:styleId="Header"', styles)
                self.assertIn('w:styleId="Footer"', styles)
                headers = [n for n in archive.namelist() if n.startswith("word/header")]
                footers = [n for n in archive.namelist() if n.startswith("word/footer")]
                self.assertTrue(headers)
                self.assertTrue(footers)
                header_xml = archive.read(headers[0]).decode("utf-8")
                footer_xml = archive.read(footers[0]).decode("utf-8")
                self.assertIn('w:val="Header"', header_xml)
                self.assertIn('w:val="Footer"', footer_xml)

                body = archive.read("word/document.xml").decode("utf-8")
                self.assertIn('w:val="nextPage"', body)


    @REQUIRES_NATIVE
    def test_excel_protection_and_sort(self) -> None:
        from openpyxl import load_workbook

        with tempfile.TemporaryDirectory(prefix="office_native_xlsx_prot_") as tmp:
            path = Path(tmp) / "protected.xlsx"
            run_ok("create_excel_workbook.py", [
                "--output-path", str(path),
                "--params", json.dumps({
                    "sheets": [{
                        "name": "数据",
                        "headers": ["名称", "数量"],
                        "data": [["b", 2], ["a", 10], ["c", None], ["d", 5]],
                        "sort": {"keys": [{"column": "数量", "ascending": False}]},
                        "auto_filter": True,
                        "protection": {"password": "secret", "allow_sort": True},
                    }]
                }, ensure_ascii=False),
            ])
            workbook = load_workbook(path)
            sheet = workbook.active
            self.assertEqual(
                [row[0] for row in sheet.iter_rows(min_row=2, max_row=5, values_only=True)],
                ["a", "d", "b", "c"],
            )
            self.assertTrue(sheet.protection.sheet)
            self.assertTrue(sheet.protection.password)
            with zipfile.ZipFile(path) as archive:
                xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("sheetProtection", xml)
                self.assertIn("sortState", xml)

    @REQUIRES_NATIVE
    def test_word_bookmark_columns_border_watermark(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_native_word_caps_") as tmp:
            path = Path(tmp) / "caps.docx"
            run_ok("create_word_document.py", [
                "--output-path", str(path),
                "--params", json.dumps({
                    "content": {"sections": [
                        {"type": "heading", "text": "标题", "level": 1},
                        {"type": "bookmark", "text": "锚点内容",
                         "bookmark": {"name": "mark1"}},
                        {"type": "paragraph", "text": "正文"},
                    ]},
                    "page_setup": {
                        "columns": 2,
                        "column_spacing": 0.3,
                        "page_border": "double",
                        "page_border_color": "C00000",
                        "watermark_text": "内部资料",
                    },
                }, ensure_ascii=False),
            ])
            with zipfile.ZipFile(path) as archive:
                body = archive.read("word/document.xml").decode("utf-8")
                self.assertIn("bookmarkStart", body)
                self.assertIn("mark1", body)
                self.assertIn('w:num="2"', body)
                self.assertIn("pgBorders", body)
                headers = [n for n in archive.namelist() if n.startswith("word/header")]
                self.assertTrue(headers)
                self.assertIn("_x0000_t136", archive.read(headers[0]).decode("utf-8"))

    @REQUIRES_NATIVE
    def test_ppt_sections_show_settings_footer(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_native_ppt_caps_") as tmp:
            path = Path(tmp) / "caps.pptx"
            run_ok("create_powerpoint.py", [
                "--output-path", str(path),
                "--params", json.dumps({
                    "slides": [
                        {"layout": "title", "title": "开场",
                         "footer_text": "内部资料", "show_slide_number": True},
                        {"layout": "content", "title": "一",
                         "body": [{"type": "paragraph", "text": "A"}]},
                        {"layout": "content", "title": "二",
                         "body": [{"type": "paragraph", "text": "B"}]},
                    ],
                    "sections": [
                        {"name": "开场", "slide_start": 1},
                        {"name": "主体", "slide_start": 2},
                    ],
                    "show_settings": {
                        "show_type": "kiosk", "loop": True,
                        "range_start": 1, "range_end": 3,
                    },
                }, ensure_ascii=False),
            ])
            with zipfile.ZipFile(path) as archive:
                presentation = archive.read("ppt/presentation.xml").decode("utf-8")
                props = archive.read("ppt/presProps.xml").decode("utf-8")
                slide1 = archive.read("ppt/slides/slide1.xml").decode("utf-8")
                self.assertIn("sectionLst", presentation)
                self.assertIn("主体", presentation)
                self.assertIn("showPr", props)
                self.assertIn("kiosk", props)
                self.assertIn('loop="1"', props)
                self.assertIn("内部资料", slide1)
                self.assertIn("slidenum", slide1)


if __name__ == "__main__":
    unittest.main()
