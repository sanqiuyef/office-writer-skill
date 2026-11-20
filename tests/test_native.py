"""Integration tests for the optional C# office-native backend.

These tests are skipped when the backend binary or the required Office
application is unavailable, so the Python-only test suite stays runnable.
"""

from __future__ import annotations

import csv
import io
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


def progid_exists(prog_id: str) -> bool:
    if os.name != "nt":
        return False
    import winreg

    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog_id))
        return True
    except OSError:
        return False


def process_running(name: str) -> bool:
    if os.name != "nt":
        return False
    proc = subprocess.run(
        ["tasklist", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        encoding="mbcs",
        errors="replace",
        check=False,
    )
    for row in csv.reader(io.StringIO(proc.stdout)):
        if row and row[0].strip().lower() == f"{name.lower()}.exe":
            return True
    return False


HAS_NATIVE = native_binary() is not None
REQUIRES_NATIVE = unittest.skipUnless(HAS_NATIVE, "需要 office-native 后端（tools/office-native/dist）")
REQUIRES_PPT = unittest.skipUnless(
    HAS_NATIVE and progid_exists("PowerPoint.Application") and not process_running("POWERPNT"),
    "需要 PowerPoint 且当前没有运行中的 POWERPNT",
)
REQUIRES_EXCEL = unittest.skipUnless(
    HAS_NATIVE and progid_exists("Excel.Application") and not process_running("EXCEL"),
    "需要 Excel 且当前没有运行中的 EXCEL",
)


def run_script(script: str, args: list[str]) -> subprocess.CompletedProcess:
    env = os.environ.copy()
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


def json_arg(value: object) -> str:
    return json.dumps(value, ensure_ascii=False)


class NativeBackendTest(unittest.TestCase):
    @REQUIRES_NATIVE
    def test_word_revisions_footnotes_content_controls(self) -> None:
        from docx import Document

        with tempfile.TemporaryDirectory(prefix="office_native_word_") as tmp:
            path = Path(tmp) / "doc.docx"
            document = Document()
            document.add_paragraph("正文第一段")
            document.add_paragraph("第二段")
            document.save(path)

            data = run_ok("word_manage_revisions.py", [
                "--file-path", str(path), "--action", "track-insert",
                "--params", json_arg({"text": "新增修订", "paragraph_index": 1}),
            ])
            self.assertEqual(data["action"], "track-insert")
            listed = run_ok("word_manage_revisions.py", [
                "--file-path", str(path), "--action", "list",
            ])
            self.assertEqual(listed["insertions"], 1)
            accepted = run_ok("word_manage_revisions.py", [
                "--file-path", str(path), "--action", "accept-all",
            ])
            self.assertEqual(accepted["insertions_resolved"], 1)

            footnote = run_ok("word_manage_footnotes.py", [
                "--file-path", str(path), "--action", "add",
                "--params", json_arg({"text": "脚注内容", "anchor_text": "第二段"}),
            ])
            self.assertEqual(footnote["footnote_id"], 1)

            run_ok("word_manage_content_controls.py", [
                "--file-path", str(path), "--action", "add-text",
                "--params", json_arg({"tag": "name", "text": "张三", "paragraph_index": 1}),
            ])
            controls = run_ok("word_manage_content_controls.py", [
                "--file-path", str(path), "--action", "list",
            ])
            self.assertEqual(controls["count"], 1)
            self.assertEqual(controls["controls"][0]["text"], "张三")

            with zipfile.ZipFile(path) as archive:
                names = archive.namelist()
                xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("word/footnotes.xml", names)
            self.assertIn("<w:sdt>", xml)
            self.assertIn("新增修订", xml)

    @REQUIRES_PPT
    def test_ppt_effects_and_smartart(self) -> None:
        from pptx import Presentation

        with tempfile.TemporaryDirectory(prefix="office_native_ppt_") as tmp:
            path = Path(tmp) / "deck.pptx"
            presentation = Presentation()
            slide = presentation.slides.add_slide(presentation.slide_layouts[1])
            slide.shapes.title.text = "测试"
            presentation.save(path)

            transition = run_ok("ppt_manage_effects.py", [
                "--file-path", str(path), "--action", "transition-set",
                "--params", json_arg({"slide": 1, "effect": "fade", "duration": 1.0}),
            ])
            self.assertEqual(transition["entry_effect"], 1793)
            listed = run_ok("ppt_manage_effects.py", [
                "--file-path", str(path), "--action", "transition-list",
                "--params", json_arg({"slide": 1}),
            ])
            self.assertEqual(listed["entry_effect"], 1793)

            comment = run_ok("ppt_manage_effects.py", [
                "--file-path", str(path), "--action", "comments-add",
                "--params", json_arg({"slide": 1, "text": "请修改标题"}),
            ])
            self.assertEqual(comment["text"], "请修改标题")

            layouts = run_ok("ppt_manage_smartart.py", [
                "--file-path", str(path), "--action", "layouts",
            ])
            self.assertGreater(layouts["count"], 0)
            added = run_ok("ppt_manage_smartart.py", [
                "--file-path", str(path), "--action", "add",
                "--params", json_arg({
                    "slide": 1,
                    "layout": layouts["layouts"][0]["name"],
                    "texts": ["步骤一", "步骤二"],
                }),
            ])
            self.assertEqual(added["nodes_written"], 2)

    @REQUIRES_EXCEL
    def test_excel_pivots_and_pdf(self) -> None:
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory(prefix="office_native_excel_") as tmp:
            path = Path(tmp) / "book.xlsx"
            pdf = Path(tmp) / "out.pdf"
            workbook = Workbook()
            data_sheet = workbook.active
            data_sheet.title = "数据"
            data_sheet.append(["区域", "月份", "金额"])
            for row in (
                ["华北", "1月", 120], ["华北", "2月", 130],
                ["华东", "1月", 90], ["华东", "2月", 95],
            ):
                data_sheet.append(row)
            workbook.create_sheet("透视")
            workbook.save(path)

            created = run_ok("excel_manage_pivots.py", [
                "--file-path", str(path), "--action", "create",
                "--params", json_arg({
                    "source_sheet": "数据", "source_range": "A1:C5",
                    "target_sheet": "透视", "target_cell": "A3",
                    "name": "PivotTest", "row_field": "区域",
                    "column_field": "月份", "data_field": "金额",
                    "function": "sum",
                }),
            ])
            self.assertEqual(created["name"], "PivotTest")
            listed = run_ok("excel_manage_pivots.py", [
                "--file-path", str(path), "--action", "list",
            ])
            self.assertEqual(listed["count"], 1)

            exported = run_ok("excel_print.py", [
                "--file-path", str(path), "--action", "pdf",
                "--output", str(pdf), "--params", json_arg({
                    "overwrite": True, "confirm_overwrite": True,
                }),
            ])
            self.assertEqual(exported["action"], "pdf")
            self.assertTrue(pdf.is_file())
            self.assertGreater(pdf.stat().st_size, 0)

            with zipfile.ZipFile(path) as archive:
                self.assertTrue(
                    any(name.startswith("xl/pivotTables/") for name in archive.namelist())
                )


if __name__ == "__main__":
    unittest.main()
