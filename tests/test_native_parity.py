"""Parity checks: Python engine vs C# office-native for migrated file-level commands.

These tests are skipped when the published backend is missing. They compare the
document structure produced by both engines, not raw bytes.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
BINARY = ROOT / "tools/office-native/dist/office-native.exe"
HAS_BINARY = BINARY.is_file()


def run_entry(script: str, args: list[str], csharp: bool) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    if csharp:
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


def run_ok(script: str, args: list[str], csharp: bool) -> dict:
    result = run_entry(script, args, csharp)
    if result.returncode != 0:
        raise AssertionError(
            f"{script} (csharp={csharp}) exited {result.returncode}: "
            f"{result.stdout}\n{result.stderr}"
        )
    payload = json.loads(result.stdout)
    if not payload.get("ok"):
        raise AssertionError(payload)
    return payload["data"]


def slide_text_lines(slide) -> set[str]:
    lines: set[str] = set()
    for shape in slide.shapes:
        if shape.has_text_frame:
            lines.update(
                line.strip() for line in shape.text_frame.text.splitlines()
                if line.strip()
            )
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        lines.add(cell.text.strip())
    return lines


def graphical_shapes(slide) -> list[tuple]:
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    return [
        (
            shape.shape_type,
            round((shape.left or 0) / 914400, 2),
            round((shape.top or 0) / 914400, 2),
            round((shape.width or 0) / 914400, 2),
            round((shape.height or 0) / 914400, 2),
        )
        for shape in slide.shapes
        if shape.shape_type != MSO_SHAPE_TYPE.TEXT_BOX
    ]


def json_arg(value: object) -> str:
    return json.dumps(value, ensure_ascii=False)


@unittest.skipUnless(HAS_BINARY, "需要已发布的 office-native 后端")
class NativeParityTest(unittest.TestCase):
    def test_word_create_parity(self) -> None:
        from docx import Document

        with tempfile.TemporaryDirectory(prefix="office_parity_word_") as tmp:
            work = Path(tmp)
            params = {
                "content": {"sections": [
                    {"type": "heading", "text": "标题", "level": 1},
                    {"type": "paragraph", "text": "正文段落"},
                    {"type": "table", "table": {
                        "headers": ["A", "B"], "rows": [["1", "2"]]}},
                    {"type": "bullet_list", "items": ["甲", "乙"]},
                ], "properties": {"title": "文档标题", "author": "tester"}},
                "header_footer": {"footer_text": "第 {page} 页"},
            }
            py_path = work / "py.docx"
            cs_path = work / "cs.docx"
            run_ok("create_word_document.py",
                   ["--output-path", str(py_path), "--params", json_arg(params)], csharp=False)
            run_ok("create_word_document.py",
                   ["--output-path", str(cs_path), "--params", json_arg(params)], csharp=True)

            py_doc = Document(py_path)
            cs_doc = Document(cs_path)
            py_texts = [paragraph.text for paragraph in py_doc.paragraphs]
            cs_texts = [paragraph.text for paragraph in cs_doc.paragraphs]
            self.assertEqual(cs_texts, py_texts)
            self.assertEqual(len(cs_doc.tables), len(py_doc.tables))
            self.assertEqual(cs_doc.core_properties.title, py_doc.core_properties.title)
            self.assertEqual(cs_doc.core_properties.author, py_doc.core_properties.author)

    def test_excel_create_parity(self) -> None:
        from openpyxl import load_workbook

        with tempfile.TemporaryDirectory(prefix="office_parity_excel_") as tmp:
            work = Path(tmp)
            params = {"sheets": [{
                "name": "销售",
                "headers": ["区域", "数量", "金额"],
                "data": [["华北", 120, 3600], ["华东", 90, 2700]],
                "freeze": "A2",
                "header_style": {"bold": True, "bg_color": "DDEBF7"},
                "charts": [{
                    "chart_type": "line", "data_range": "B1:B3",
                    "categories_range": "A2:A3", "anchor": "E2",
                }],
                "conditional_formats": [{
                    "range": "C2:C3", "type": "color_scale",
                    "min_color": "F8696B", "max_color": "63BE7B",
                }],
                "data_validations": [{
                    "range": "B2:B3", "type": "whole",
                    "formula1": "0", "formula2": "1000",
                }],
                "tables": [{"range": "A1:C3", "name": "SalesTable"}],
            }]}
            py_path = work / "py.xlsx"
            cs_path = work / "cs.xlsx"
            run_ok("create_excel_workbook.py",
                   ["--output-path", str(py_path), "--params", json_arg(params)], csharp=False)
            run_ok("create_excel_workbook.py",
                   ["--output-path", str(cs_path), "--params", json_arg(params)], csharp=True)

            py_ws = load_workbook(py_path)["销售"]
            cs_ws = load_workbook(cs_path)["销售"]
            self.assertEqual(
                [[cell.value for cell in row] for row in cs_ws.iter_rows(min_row=1, max_row=3)],
                [[cell.value for cell in row] for row in py_ws.iter_rows(min_row=1, max_row=3)],
            )
            self.assertTrue(cs_ws["A1"].font.bold)
            self.assertEqual(
                cs_ws["A1"].fill.fgColor.rgb[-6:],
                py_ws["A1"].fill.fgColor.rgb[-6:],
            )
            self.assertEqual(len(cs_ws._charts), len(py_ws._charts))
            self.assertEqual(
                len(list(cs_ws.conditional_formatting)),
                len(list(py_ws.conditional_formatting)),
            )
            self.assertEqual(
                len(cs_ws.data_validations.dataValidation),
                len(py_ws.data_validations.dataValidation),
            )
            self.assertEqual(set(cs_ws.tables.keys()), set(py_ws.tables.keys()))

    def test_prepare_create_parity(self) -> None:
        from pptx import Presentation

        with tempfile.TemporaryDirectory(prefix="office_parity_ppt_") as tmp:
            work = Path(tmp)
            params = {"slides": [{
                "layout": "content",
                "title": "标题",
                "body": [
                    {"type": "paragraph", "text": "正文"},
                    {"type": "bullet_list", "items": ["甲", "乙"]},
                    {"type": "table", "table": {
                        "headers": ["A", "B"], "rows": [["1", "2"]],
                        "x": 1, "y": 3, "width": 8, "height": 1.5,
                    }},
                    {"type": "chart", "chart": {
                        "chart_type": "doughnut", "categories": ["x", "y"],
                        "series": [{"name": "s", "values": [3, 7]}],
                        "x": 1, "y": 4.5, "width": 4, "height": 2.5,
                    }},
                ],
            }, {
                "layout": "two_content", "title": "双栏",
                "left": [{"type": "paragraph", "text": "左"}],
                "right": [{"type": "paragraph", "text": "右"}],
            }]}
            py_path = work / "py.pptx"
            cs_path = work / "cs.pptx"
            run_ok("create_powerpoint.py",
                   ["--output-path", str(py_path), "--params", json_arg(params)], csharp=False)
            run_ok("create_powerpoint.py",
                   ["--output-path", str(cs_path), "--params", json_arg(params)], csharp=True)

            py_prs = Presentation(py_path)
            cs_prs = Presentation(cs_path)
            self.assertEqual(len(cs_prs.slides), len(py_prs.slides))
            for py_slide, cs_slide in zip(py_prs.slides, cs_prs.slides):
                # 列表承载方式不同：C# 并入正文占位符，Python 用独立文本框。
                # 比较文本行集合与图形元素（表格/图表/占位符）几何。
                self.assertEqual(slide_text_lines(cs_slide), slide_text_lines(py_slide))
                self.assertEqual(graphical_shapes(cs_slide), graphical_shapes(py_slide))

    def test_validate_and_txt_parity(self) -> None:
        fixtures = [
            ROOT / "tests/fixtures/validate_clean.docx",
            ROOT / "tests/fixtures/validate_issues.docx",
            ROOT / "tests/fixtures/validate_toc_stale.docx",
        ]
        for fixture in fixtures:
            with self.subTest(fixture=fixture.name):
                py = run_ok("validate_word_document.py",
                            ["--file-path", str(fixture)], csharp=False)
                cs = run_ok("validate_word_document.py",
                            ["--file-path", str(fixture)], csharp=True)
                self.assertEqual(
                    [(check["id"], check["severity"]) for check in cs["checks"]],
                    [(check["id"], check["severity"]) for check in py["checks"]],
                )
                self.assertEqual(cs["errors"], py["errors"])
                self.assertEqual(cs["warnings"], py["warnings"])

        with tempfile.TemporaryDirectory(prefix="office_parity_txt_") as tmp:
            path = Path(tmp) / "a.txt"
            path.write_text("第一行\n第二行\n", encoding="utf-8")
            py = run_ok("read_txt_document.py",
                        ["--file-path", str(path), "--max-chars", "100"], csharp=False)
            cs = run_ok("read_txt_document.py",
                        ["--file-path", str(path), "--max-chars", "100"], csharp=True)
            self.assertEqual(cs["markdown"], py["markdown"])


if __name__ == "__main__":
    unittest.main()
