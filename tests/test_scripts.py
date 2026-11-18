#!/usr/bin/env python3
"""Deterministic CLI contract tests for office-writer.

Run from the skill root:
    python -X utf8 -m unittest discover -s tests -v
"""

from __future__ import annotations

import csv
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from cases import CASES

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
EXPECTED = Path(__file__).resolve().parent / "expected_outputs"
FIXTURES = Path(__file__).resolve().parent / "fixtures"

FAKE_DSOFFICE_SOURCE = '''\
import json
import os
import sys
from pathlib import Path


def main():
    args = sys.argv[1:]
    output_dir = None
    for index, arg in enumerate(args):
        if arg == "--output-dir" and index + 1 < len(args):
            output_dir = Path(args[index + 1])
    missing = [item for item in os.environ.get("FAKE_MISSING_FONTS", "").split(",") if item]
    images = []
    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        page = output_dir / "page-0001.png"
        page.write_bytes(b"\\x89PNG\\r\\n\\x1a\\n")
        images.append({"index": 1, "width": 1, "height": 1, "path": str(page)})
    payload = {
        "backend": "fake",
        "pageCount": int(os.environ.get("FAKE_PAGE_COUNT", "2")),
        "images": images,
        "missingFonts": missing,
    }
    print(json.dumps(payload))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
'''


def write_fake_dsoffice(work: Path) -> Path:
    """Create a fake dsoffice launcher usable through MCP_OFFICE_DSOFFICE."""
    helper = work / "fake_dsoffice.py"
    helper.write_text(FAKE_DSOFFICE_SOURCE, encoding="utf-8")
    if os.name == "nt":
        launcher = work / "dsoffice.cmd"
        launcher.write_text(
            f'@echo off\r\n"{sys.executable}" "{helper}" %*\r\n',
            encoding="utf-8",
        )
    else:
        launcher = work / "dsoffice"
        launcher.write_text(
            f'#!/bin/sh\nexec "{sys.executable}" "{helper}" "$@"\n',
            encoding="utf-8",
        )
        launcher.chmod(0o755)
    return launcher


def find_tesseract() -> str | None:
    candidates = [
        os.environ.get("OFFICE_TESSERACT_CMD"),
        os.environ.get("TESSERACT_CMD"),
        shutil.which("tesseract"),
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(candidate)
    return None


OFFICE_PROCESS_NAMES = {"winword.exe", "excel.exe", "powerpnt.exe"}


def running_office_pids() -> dict[str, set[int]]:
    """Snapshot Word/Excel/PowerPoint PIDs (Windows only)."""
    snapshot = {name: set() for name in OFFICE_PROCESS_NAMES}
    if os.name != "nt":
        return snapshot
    proc = subprocess.run(
        ["tasklist", "/FO", "CSV", "/NH"],
        capture_output=True,
        check=False,
    )
    text = proc.stdout.decode("mbcs", errors="replace")
    for row in csv.reader(io.StringIO(text)):
        if len(row) < 2:
            continue
        name = row[0].strip().lower()
        if name not in snapshot:
            continue
        try:
            snapshot[name].add(int(row[1].strip()))
        except ValueError:
            continue
    return snapshot


def wait_for_new_office_processes(before: dict[str, set[int]], timeout: float = 20.0) -> set[tuple[str, int]]:
    """Wait for processes started during a test to exit; return leftovers."""
    deadline = time.time() + timeout
    while True:
        after = running_office_pids()
        leftovers = {
            (name, pid)
            for name, pids in after.items()
            for pid in pids - before.get(name, set())
        }
        if not leftovers or time.time() >= deadline:
            return leftovers
        time.sleep(0.5)


def com_is_available() -> bool:
    if os.name != "nt":
        return False
    sys.path.insert(0, str(SCRIPTS))
    from office_engine.office_com import com_is_available as available

    return available()


REQUIRES_COM = unittest.skipUnless(
    com_is_available(), "需要 Windows + pywin32 + Microsoft Office/WPS"
)

ENTRY_SCRIPTS = [
    "create_word_document.py",
    "edit_word_document.py",
    "create_word_from_markdown.py",
    "create_excel_workbook.py",
    "edit_excel_workbook.py",
    "create_powerpoint.py",
    "edit_powerpoint.py",
    "create_pdf_document.py",
    "convert_to_pdf.py",
    "render_document_pages.py",
    "recalculate_excel_formulas.py",
    "get_open_office_documents.py",
    "edit_open_word_document.py",
    "edit_open_excel_workbook.py",
    "edit_open_powerpoint.py",
    "ppt_list_color_schemes.py",
    "ppt_apply_color_scheme.py",
    "ppt_optimize_slide_layout.py",
    "ppt_clean_unused_placeholders.py",
    "ppt_analyze_template.py",
    "ppt_create_framework_diagram.py",
    "ppt_add_flowchart_slide.py",
    "read_word_document.py",
    "read_excel_workbook.py",
    "read_powerpoint.py",
    "read_pdf_document.py",
    "read_pdf_structured.py",
    "read_txt_document.py",
    "inspect_office_document.py",
    "validate_word_document.py",
    "merge_office_template.py",
    "batch_edit_office_document.py",
]


def run_script(script: str, args: list[str], env: dict | None = None) -> subprocess.CompletedProcess:
    merged = os.environ.copy()
    if env:
        merged.update(env)
    return subprocess.run(
        [sys.executable, "-X", "utf8", str(SCRIPTS / script), *args],
        cwd=SCRIPTS,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=merged,
    )


def json_arg(value: object) -> str:
    return json.dumps(value, ensure_ascii=False)


class ScriptContractTest(unittest.TestCase):
    def test_fixed_cases(self) -> None:
        for case in CASES:
            with self.subTest(case=case["name"]):
                result = run_script(case["script"], case["args"])
                self.assertEqual(result.returncode, case["exit_code"], result.stderr)
                payload = json.loads(result.stdout)
                self.assertTrue(payload["ok"], payload)
                self.assertEqual(payload["tool"], case["script"][:-3])
                expected = json.loads((EXPECTED / f"{case['name']}.json").read_text(encoding="utf-8"))
                self.assertEqual(payload["data"], expected)

    def test_usage_error_exit_2(self) -> None:
        for script, args in (
            ("read_word_document.py", []),
            ("read_word_document.py", ["--params", "{bad json"]),
            ("convert_to_pdf.py", ["--input-path", "../tests/fixtures/sample.docx", "--if-exists", "bad"]),
            ("ppt_apply_color_scheme.py", ["--file-path", "../tests/fixtures/sample.pptx", "--scheme-name", "nope"]),
            (
                "create_word_document.py",
                ["--output-path", "unused.docx", "--params", json_arg({"content": {"sections": "bad"}})],
            ),
        ):
            with self.subTest(script=script, args=args):
                result = run_script(script, args)
                self.assertEqual(result.returncode, 2, result.stderr)
                payload = json.loads(result.stdout)
                self.assertFalse(payload["ok"])
                self.assertEqual(payload["tool"], script[:-3])
                self.assertTrue(payload["error"])

    def test_runtime_error_exit_3(self) -> None:
        for script, args in (
            ("read_word_document.py", ["--file-path", "../tests/fixtures/missing.docx"]),
            ("convert_to_pdf.py", ["--input-path", "../tests/fixtures/missing.docx"]),
            ("ppt_apply_color_scheme.py", ["--file-path", "../tests/fixtures/missing.pptx"]),
        ):
            with self.subTest(script=script, args=args):
                result = run_script(script, args)
                self.assertEqual(result.returncode, 3, result.stderr)
                payload = json.loads(result.stdout)
                self.assertFalse(payload["ok"])
                self.assertEqual(payload["tool"], script[:-3])
                self.assertTrue(payload["error"])

    def test_help_every_entrypoint(self) -> None:
        on_disk = sorted(path.name for path in SCRIPTS.glob("*.py") if not path.name.startswith("_"))
        self.assertEqual(on_disk, sorted(ENTRY_SCRIPTS))
        for script in ENTRY_SCRIPTS:
            with self.subTest(script=script):
                result = run_script(script, ["--help"])
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertIn("usage:", result.stdout.lower())

    def test_stdout_is_single_json_object(self) -> None:
        for script, args in (
            ("ppt_list_color_schemes.py", []),
            ("read_word_document.py", ["--file-path", "../tests/fixtures/missing.docx"]),
        ):
            with self.subTest(script=script):
                result = run_script(script, args)
                text = result.stdout.strip()
                self.assertTrue(text.startswith("{") and text.endswith("}"), text)
                json.loads(text)

    def test_params_file(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_params_") as tmp:
            params_path = Path(tmp) / "params.json"
            params_path.write_text(
                json_arg({"file_path": str(FIXTURES / "sample.txt")}),
                encoding="utf-8",
            )
            result = run_script("read_txt_document.py", ["--params-file", str(params_path)])
            self.assertEqual(result.returncode, 0, result.stderr)
            payload = json.loads(result.stdout)
            self.assertIn("示例文本", payload["data"]["markdown"])

    def test_params_override_defaults(self) -> None:
        result = run_script(
            "read_txt_document.py",
            ["--params", json_arg({"file_path": str(FIXTURES / "sample.txt"), "max_chars": 10})],
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(json.loads(result.stdout)["data"]["result"]["truncated"])

        with tempfile.TemporaryDirectory(prefix="office_writer_params_") as tmp:
            pptx_path = Path(tmp) / "deck.pptx"
            shutil.copy2(FIXTURES / "sample.pptx", pptx_path)
            result = run_script(
                "ppt_apply_color_scheme.py",
                ["--params", json_arg({"file_path": str(pptx_path), "scheme_name": "modern-dark"})],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(json.loads(result.stdout)["data"]["scheme"], "modern-dark")

    def test_create_read_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_roundtrip_") as tmp:
            work = Path(tmp)

            docx_path = work / "report.docx"
            word_params = {
                "content": {
                    "sections": [
                        {"type": "heading", "text": "测试标题", "level": 1},
                        {"type": "paragraph", "text": "测试正文内容。"},
                        {
                            "type": "table",
                            "table": {"headers": ["指标", "值"], "rows": [["交付率", "97%"]]},
                        },
                    ]
                }
            }
            result = run_script(
                "create_word_document.py",
                ["--output-path", str(docx_path), "--params", json_arg(word_params)],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            created = json.loads(result.stdout)["data"]
            self.assertTrue(created["success"])
            self.assertGreater(created["file_size"], 0)
            self.assertTrue(docx_path.is_file())

            read_result = run_script("read_word_document.py", ["--file-path", str(docx_path)])
            self.assertEqual(read_result.returncode, 0, read_result.stderr)
            self.assertIn("测试正文内容", json.loads(read_result.stdout)["data"]["markdown"])

            validate_result = run_script("validate_word_document.py", ["--file-path", str(docx_path)])
            self.assertEqual(validate_result.returncode, 0, validate_result.stderr)
            self.assertEqual(json.loads(validate_result.stdout)["data"]["errors"], 0)

            inspect_result = run_script("inspect_office_document.py", ["--file-path", str(docx_path)])
            self.assertEqual(inspect_result.returncode, 0, inspect_result.stderr)
            self.assertGreater(json.loads(inspect_result.stdout)["data"]["node_count"], 0)

            xlsx_path = work / "table.xlsx"
            excel_params = {
                "sheets": [{
                    "name": "数据",
                    "headers": ["区域", "数量"],
                    "data": [["华北", 120], ["华东", 180]],
                    "freeze": "A2",
                }]
            }
            result = run_script(
                "create_excel_workbook.py",
                ["--output-path", str(xlsx_path), "--params", json_arg(excel_params)],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            read_result = run_script("read_excel_workbook.py", ["--file-path", str(xlsx_path)])
            self.assertEqual(read_result.returncode, 0, read_result.stderr)
            self.assertIn("华北", json.loads(read_result.stdout)["data"]["markdown"])

            pptx_path = work / "deck.pptx"
            ppt_params = {
                "slides": [
                    {"layout": "title", "title": "测试演示", "subtitle": "副标题"},
                    {"layout": "content", "title": "要点", "body": [
                        {"type": "bullet_list", "items": ["第一点", "第二点"]}
                    ]},
                ]
            }
            result = run_script(
                "create_powerpoint.py",
                ["--output-path", str(pptx_path), "--params", json_arg(ppt_params)],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            read_result = run_script("read_powerpoint.py", ["--file-path", str(pptx_path)])
            self.assertEqual(read_result.returncode, 0, read_result.stderr)
            self.assertIn("测试演示", json.loads(read_result.stdout)["data"]["markdown"])

            pdf_path = work / "document.pdf"
            pdf_params = {
                "content": {
                    "sections": [
                        {"type": "heading", "text": "PDF 标题", "level": 1},
                        {"type": "paragraph", "text": "PDF 正文内容。"},
                    ]
                }
            }
            result = run_script(
                "create_pdf_document.py",
                ["--output-path", str(pdf_path), "--params", json_arg(pdf_params)],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            read_result = run_script("read_pdf_document.py", ["--file-path", str(pdf_path)])
            self.assertEqual(read_result.returncode, 0, read_result.stderr)
            self.assertIn("PDF 正文内容", json.loads(read_result.stdout)["data"]["markdown"])

    def test_markdown_edit_merge_batch_and_ppt_effects(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_ops_") as tmp:
            work = Path(tmp)

            markdown = work / "source.md"
            markdown.write_text("# 标题\n\n正文段落。\n", encoding="utf-8")
            docx_path = work / "from_markdown.docx"
            result = run_script(
                "create_word_from_markdown.py",
                ["--source-path", str(markdown), "--output-path", str(docx_path)],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            read_result = run_script("read_word_document.py", ["--file-path", str(docx_path)])
            self.assertIn("正文段落", json.loads(read_result.stdout)["data"]["markdown"])

            copy_path = work / "edit.docx"
            shutil.copy2(FIXTURES / "sample.docx", copy_path)
            edit_params = {"append_sections": [{"type": "paragraph", "text": "追加段落。"}]}
            result = run_script(
                "edit_word_document.py",
                ["--file-path", str(copy_path), "--params", json_arg(edit_params)],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(json.loads(result.stdout)["data"]["operation_count"], 1)
            read_result = run_script("read_word_document.py", ["--file-path", str(copy_path)])
            self.assertIn("追加段落", json.loads(read_result.stdout)["data"]["markdown"])

            merged_path = work / "merged.docx"
            merge_params = {
                "data": {
                    "company": "示例公司",
                    "footer": "内部资料",
                    "title": "项目总结",
                    "summary": "本季度完成核心模块。",
                    "body": "详细内容。",
                }
            }
            result = run_script(
                "merge_office_template.py",
                [
                    "--template-path", str(ROOT / "templates" / "word_template.docx"),
                    "--output-path", str(merged_path),
                    "--params", json_arg(merge_params),
                ],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            merged = json.loads(result.stdout)["data"]
            self.assertGreaterEqual(merged["replaced_count"], 5)
            self.assertEqual(merged["unresolved_placeholders"], [])

            batch_input = work / "batch.docx"
            batch_output = work / "batch_out.docx"
            shutil.copy2(FIXTURES / "sample.docx", batch_input)
            batch_params = {
                "operations": [
                    {"op": "set", "path": "/body/p[2]", "props": {"text": "更新后的段落。"}}
                ]
            }
            result = run_script(
                "batch_edit_office_document.py",
                [
                    "--file-path", str(batch_input),
                    "--output-path", str(batch_output),
                    "--params", json_arg(batch_params),
                ],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue(json.loads(result.stdout)["data"]["saved"])

            ppt_copy = work / "effects.pptx"
            shutil.copy2(FIXTURES / "sample.pptx", ppt_copy)
            for script, extra in (
                ("ppt_apply_color_scheme.py", []),
                ("ppt_optimize_slide_layout.py", []),
                ("ppt_clean_unused_placeholders.py", []),
            ):
                with self.subTest(script=script):
                    result = run_script(script, ["--file-path", str(ppt_copy), *extra])
                    self.assertEqual(result.returncode, 0, result.stderr)

            framework_path = work / "framework.pptx"
            framework_params = {
                "diagram": {
                    "title": "流程",
                    "stages": [
                        {"label": "输入", "nodes": [{"id": "a", "text": "数据"}]},
                        {"label": "输出", "nodes": [{"id": "b", "text": "结果"}]},
                    ],
                    "edges": [{"from_id": "a", "to_id": "b"}],
                }
            }
            result = run_script(
                "ppt_create_framework_diagram.py",
                ["--output-path", str(framework_path), "--params", json_arg(framework_params)],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(json.loads(result.stdout)["data"]["slides_count"], 1)


    def test_render_page_count_uses_actual_artifacts(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_render_") as tmp:
            work = Path(tmp)
            source = work / "sample.pdf"
            shutil.copy2(FIXTURES / "sample.pdf", source)
            launcher = write_fake_dsoffice(work)
            output_dir = work / "pages"
            result = run_script(
                "render_document_pages.py",
                ["--input-path", str(source), "--output-dir", str(output_dir)],
                env={"MCP_OFFICE_DSOFFICE": str(launcher), "FAKE_PAGE_COUNT": "2"},
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            pngs = sorted(output_dir.glob("*.png"))
            self.assertEqual(data["page_count"], 1)
            self.assertEqual(data["page_count_reported"], 2)
            self.assertEqual(len(data["images"]), 1)
            self.assertEqual(data["page_count"], len(pngs))
            self.assertIn("按实际产物修正", data["message"])

    def test_convert_font_substitutions(self) -> None:
        sys.path.insert(0, str(SCRIPTS))
        from office_engine.office_convert import plan_font_substitutions

        self.assertEqual(
            plan_font_substitutions(["Courier"], {"Consolas", "Cascadia Mono"}),
            [{"missing": "Courier", "substitute": "Consolas"}],
        )
        self.assertEqual(
            plan_font_substitutions(["Courier"], {"Cascadia Mono"}),
            [{"missing": "Courier", "substitute": "Cascadia Mono"}],
        )
        self.assertEqual(plan_font_substitutions(["Courier"], set()), [])
        self.assertEqual(plan_font_substitutions(["Arial"], {"Consolas"}), [])

        with tempfile.TemporaryDirectory(prefix="office_writer_fonts_") as tmp:
            work = Path(tmp)
            source = work / "font.docx"
            shutil.copy2(FIXTURES / "sample.docx", source)
            launcher = write_fake_dsoffice(work)
            env = {
                "MCP_OFFICE_DSOFFICE": str(launcher),
                "FAKE_MISSING_FONTS": "Courier,NotAFont",
            }
            result = run_script(
                "convert_to_pdf.py",
                ["--input-path", str(source), "--output-path", str(work / "font.pdf")],
                env=env,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            self.assertEqual(data["missing_fonts"], ["Courier", "NotAFont"])
            self.assertIn("缺失字体", data["message"])
            self.assertIn("font_substitutions", data)
            for item in data["font_substitutions"]:
                self.assertEqual(item["missing"], "Courier")
                self.assertIn(item["substitute"], ("Consolas", "Cascadia Mono", "DejaVu Sans Mono"))

            result = run_script(
                "convert_to_pdf.py",
                ["--input-path", str(source), "--output-path", str(work / "font2.pdf")],
                env={"MCP_OFFICE_DSOFFICE": str(launcher), "FAKE_MISSING_FONTS": "NotAFont"},
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            self.assertEqual(data["font_substitutions"], [])
            self.assertIn("缺失字体（已按字形度量替换）: NotAFont", data["message"])

    def test_batch_edit_in_place_backup(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_backup_") as tmp:
            work = Path(tmp)
            target = work / "inplace.docx"
            shutil.copy2(FIXTURES / "sample.docx", target)
            original = target.read_bytes()
            operations = {
                "operations": [
                    {"op": "set", "path": "/body/p[2]", "props": {"text": "第一次更新。"}}
                ]
            }
            result = run_script(
                "batch_edit_office_document.py",
                ["--file-path", str(target), "--params", json_arg(operations)],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            self.assertTrue(data["saved"])
            backup = Path(data["backup_path"])
            self.assertTrue(backup.is_file())
            self.assertRegex(backup.name, r"^inplace\.docx\.bak-\d{8}-\d{6}(?:-\d+)?$")
            self.assertEqual(backup.read_bytes(), original)
            read_result = run_script("read_word_document.py", ["--file-path", str(target)])
            self.assertEqual(read_result.returncode, 0, read_result.stderr)
            self.assertIn("第一次更新", json.loads(read_result.stdout)["data"]["markdown"])

            result = run_script(
                "batch_edit_office_document.py",
                [
                    "--file-path", str(target),
                    "--no-backup",
                    "--params", json_arg({
                        "operations": [
                            {"op": "set", "path": "/body/p[2]", "props": {"text": "第二次更新。"}}
                        ]
                    }),
                ],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            self.assertNotIn("backup_path", data)
            self.assertEqual(len(list(work.glob("inplace.docx.bak-*"))), 1)

            output = work / "out.docx"
            result = run_script(
                "batch_edit_office_document.py",
                [
                    "--file-path", str(target),
                    "--output-path", str(output),
                    "--params", json_arg(operations),
                ],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertNotIn("backup_path", json.loads(result.stdout)["data"])
            self.assertTrue(output.is_file())
            self.assertEqual(len(list(work.glob("inplace.docx.bak-*"))), 1)

    def test_ppt_layout_audit_and_fix_spacing(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_layout_") as tmp:
            work = Path(tmp)
            deck = work / "layout.pptx"
            create_params = {
                "slides": [{
                    "layout": "blank",
                    "body": [
                        {"type": "shape", "shape": {
                            "shape_type": "rect", "text": "A",
                            "x": 1, "y": 1, "width": 4, "height": 2,
                            "bg_color": "2563EB",
                        }},
                        {"type": "shape", "shape": {
                            "shape_type": "rect", "text": "B",
                            "x": 3, "y": 2, "width": 4, "height": 2,
                            "bg_color": "EF4444",
                        }},
                        {"type": "shape", "shape": {
                            "shape_type": "rect", "text": "C",
                            "x": 9, "y": 6, "width": 3, "height": 2,
                            "bg_color": "22C55E",
                        }},
                    ],
                }]
            }
            result = run_script(
                "create_powerpoint.py",
                ["--output-path", str(deck), "--params", json_arg(create_params)],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            original = deck.read_bytes()

            result = run_script("ppt_optimize_slide_layout.py", ["--file-path", str(deck)])
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            types = {warning["type"] for warning in data["warnings"]}
            self.assertIn("out_of_bounds", types)
            self.assertIn("overlap", types)
            self.assertGreaterEqual(data["warning_count"], 2)
            self.assertEqual(deck.read_bytes(), original)

            result = run_script(
                "ppt_optimize_slide_layout.py",
                ["--file-path", str(deck), "--fix-spacing"],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            self.assertEqual(data["slides_optimized"], 1)
            self.assertNotEqual(deck.read_bytes(), original)

    def test_missing_dsoffice_install_hint(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_dsoffice_") as tmp:
            work = Path(tmp)
            source = work / "sample.docx"
            shutil.copy2(FIXTURES / "sample.docx", source)
            output = work / "out.pdf"
            output.write_bytes(b"existing-output")
            result = run_script(
                "convert_to_pdf.py",
                ["--input-path", str(source), "--output-path", str(output)],
                env={
                    "MCP_OFFICE_DSOFFICE": "",
                    "MODULE_OFFICE_DSOFFICE": "",
                    "PATH": "",
                    "OFFICE_CONVERT_BACKEND": "dsoffice",
                },
            )
            self.assertEqual(result.returncode, 3, result.stderr)
            payload = json.loads(result.stdout)
            self.assertFalse(payload["ok"])
            self.assertIn("npm i -g dsoffice", payload["error"])
            self.assertIn("Node.js >= 22.19", payload["error"])
            self.assertIn("Get-Command dsoffice", payload["error"])
            self.assertEqual(output.read_bytes(), b"existing-output")

    def test_validate_deterministic_checks(self) -> None:
        clean = run_script(
            "validate_word_document.py",
            ["--file-path", "../tests/fixtures/validate_clean.docx"],
        )
        self.assertEqual(clean.returncode, 0, clean.stderr)
        clean_data = json.loads(clean.stdout)["data"]
        self.assertEqual(clean_data["errors"], 0)
        self.assertEqual(clean_data["warnings"], 0)
        clean_checks = {check["id"]: check for check in clean_data["checks"]}
        for check_id in (
            "font_installed", "heading_numbering", "page_number_field", "toc_field",
            "table_columns", "empty_paragraph_ratio", "mixed_fonts",
        ):
            with self.subTest(check=check_id):
                self.assertTrue(clean_checks[check_id]["passed"], clean_checks[check_id])

        issues = run_script(
            "validate_word_document.py",
            ["--file-path", "../tests/fixtures/validate_issues.docx"],
        )
        self.assertEqual(issues.returncode, 0, issues.stderr)
        issue_data = json.loads(issues.stdout)["data"]
        self.assertEqual(issue_data["errors"], 0)
        self.assertGreaterEqual(issue_data["warnings"], 8)
        checks = {check["id"]: check for check in issue_data["checks"]}
        for check_id in (
            "heading_levels", "font_installed", "heading_numbering",
            "page_number_field", "toc_field", "table_columns",
            "empty_paragraph_ratio", "mixed_fonts",
        ):
            with self.subTest(check=check_id):
                self.assertFalse(checks[check_id]["passed"], checks[check_id])
        self.assertIn("Courier", checks["font_installed"]["message"])
        self.assertIn("不连续", checks["heading_numbering"]["message"])
        self.assertIn("PAGE", checks["page_number_field"]["message"])
        self.assertIn("TOC", checks["toc_field"]["message"])
        self.assertIn("首行", checks["table_columns"]["details"])
        self.assertIn("空段落", checks["empty_paragraph_ratio"]["details"])
        self.assertIn("宋体", checks["mixed_fonts"]["details"])
        self.assertIn("黑体", checks["mixed_fonts"]["details"])

        stale = run_script(
            "validate_word_document.py",
            ["--file-path", "../tests/fixtures/validate_toc_stale.docx"],
        )
        self.assertEqual(stale.returncode, 0, stale.stderr)
        stale_checks = {
            check["id"]: check for check in json.loads(stale.stdout)["data"]["checks"]
        }
        self.assertFalse(stale_checks["toc_field"]["passed"])
        self.assertIn("未标记为需更新", stale_checks["toc_field"]["message"])
        self.assertTrue(stale_checks["page_number_field"]["passed"])

    def test_generated_word_fields_are_word_compatible(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_word_fields_") as tmp:
            work = Path(tmp)
            docx_path = work / "fields.docx"
            params = {
                "scene": "report",
                "content": {
                    "sections": [
                        {"type": "toc", "text": "目录", "toc_levels": "1-3"},
                        {"type": "heading", "text": "报告标题", "level": 1},
                        {"type": "paragraph", "text": "正文段落。"},
                    ]
                },
            }
            result = run_script(
                "create_word_document.py",
                ["--output-path", str(docx_path), "--params", json_arg(params)],
            )
            self.assertEqual(result.returncode, 0, result.stderr)

            word_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            with zipfile.ZipFile(docx_path) as archive:
                document_root = ET.fromstring(archive.read("word/document.xml"))
                footer_root = ET.fromstring(archive.read("word/footer1.xml"))

            field_count = 0
            for root in (document_root, footer_root):
                for parent in root.iter():
                    for child in parent:
                        if child.tag in {word_ns + "fldChar", word_ns + "instrText"}:
                            field_count += 1
                            self.assertEqual(parent.tag, word_ns + "r")
            self.assertGreaterEqual(field_count, 4)

            for paragraph_properties in document_root.iter(word_ns + "pPr"):
                children = list(paragraph_properties)
                tags = [child.tag for child in children]
                if word_ns + "pBdr" not in tags:
                    continue
                border_index = tags.index(word_ns + "pBdr")
                for successor in ("spacing", "ind", "jc"):
                    if word_ns + successor in tags:
                        self.assertLess(border_index, tags.index(word_ns + successor))

            validation = run_script(
                "validate_word_document.py", ["--file-path", str(docx_path)]
            )
            self.assertEqual(validation.returncode, 0, validation.stderr)
            checks = {
                check["id"]: check
                for check in json.loads(validation.stdout)["data"]["checks"]
            }
            self.assertTrue(checks["page_number_field"]["passed"], checks["page_number_field"])
            self.assertTrue(checks["toc_field"]["passed"], checks["toc_field"])

    def test_com_backend_candidates(self) -> None:
        sys.path.insert(0, str(SCRIPTS))
        from office_engine.office_convert import _com_candidates

        self.assertEqual(
            _com_candidates("word"),
            [
                ("office", "Word.Application"),
                ("wps", "KWps.Application"),
                ("wps-fallback", "Word.Application"),
            ],
        )
        self.assertEqual(
            _com_candidates("excel"),
            [
                ("office", "Excel.Application"),
                ("wps", "KET.Application"),
                ("wps-fallback", "Excel.Application"),
            ],
        )
        self.assertEqual(
            _com_candidates("ppt"),
            [
                ("office", "PowerPoint.Application"),
                ("wps", "KWPP.Application"),
                ("wps-fallback", "PowerPoint.Application"),
            ],
        )

    @REQUIRES_COM
    def test_convert_with_com_real(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_com_convert_") as tmp:
            work = Path(tmp)
            before = running_office_pids()
            for suffix in ("docx", "xlsx", "pptx"):
                with self.subTest(format=suffix):
                    output = work / f"com_{suffix}.pdf"
                    result = run_script(
                        "convert_to_pdf.py",
                        [
                            "--input-path", str(FIXTURES / f"sample.{suffix}"),
                            "--output-path", str(output),
                            "--if-exists", "overwrite",
                        ],
                        env={"OFFICE_CONVERT_BACKEND": "com"},
                    )
                    self.assertEqual(result.returncode, 0, result.stderr)
                    data = json.loads(result.stdout)["data"]
                    self.assertEqual(data["backend"], "com")
                    self.assertEqual(data["backend_kind"], "com")
                    self.assertIn(data["com_engine"], ("office", "wps"))
                    self.assertTrue(data["com_prog_id"])
                    self.assertTrue(output.is_file(), output)
                    self.assertEqual(output.read_bytes()[:4], b"%PDF")
            self.assertEqual(wait_for_new_office_processes(before), set())

    @REQUIRES_COM
    def test_auto_fallback_when_dsoffice_missing(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_com_fallback_") as tmp:
            work = Path(tmp)
            missing = work / "missing-dsoffice.exe"
            env = {
                "MCP_OFFICE_DSOFFICE": "",
                "MODULE_OFFICE_DSOFFICE": str(missing),
                "OFFICE_CONVERT_BACKEND": "auto",
            }
            before = running_office_pids()

            pdf_path = work / "fallback.pdf"
            result = run_script(
                "convert_to_pdf.py",
                ["--input-path", str(FIXTURES / "sample.docx"), "--output-path", str(pdf_path)],
                env=env,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            self.assertEqual(data["backend"], "com")
            self.assertTrue(data["fallback_from"])
            self.assertEqual(data["fallback_from"][0]["backend"], "dsoffice")
            self.assertIn("MODULE_OFFICE_DSOFFICE", data["fallback_from"][0]["error"])
            self.assertEqual(pdf_path.read_bytes()[:4], b"%PDF")

            pages_dir = work / "pages"
            result = run_script(
                "render_document_pages.py",
                [
                    "--input-path", str(FIXTURES / "sample.docx"),
                    "--output-dir", str(pages_dir),
                    "--pages", "1",
                    "--dpi", "96",
                ],
                env=env,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            pngs = sorted(pages_dir.glob("*.png"))
            self.assertEqual(data["backend"], "com+pymupdf")
            self.assertEqual(data["page_count"], len(pngs))
            self.assertEqual(data["page_count"], 1)
            self.assertEqual(len(data["images"]), 1)
            self.assertEqual(Path(data["images"][0]["path"]).parent, pages_dir)
            self.assertFalse(list(pages_dir.glob("*.pdf")))
            self.assertTrue(data.get("fallback_from"))
            self.assertGreater(data["images"][0]["width"], 0)
            self.assertGreater(data["images"][0]["height"], 0)
            self.assertEqual(wait_for_new_office_processes(before), set())

    @REQUIRES_COM
    def test_recalculate_with_com_real(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_com_recalc_") as tmp:
            work = Path(tmp)
            source = work / "formulas.xlsx"
            result = run_script(
                "create_excel_workbook.py",
                [
                    "--output-path", str(source),
                    "--params", json_arg({
                        "sheets": [{
                            "name": "数据",
                            "headers": ["A", "B", "合计"],
                            "data": [[1, 2, "=A2+B2"], [3, 4, "=A3+B3"]],
                            "allow_formulas": True,
                        }]
                    }),
                ],
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            before = running_office_pids()
            output = work / "formulas_recalc.xlsx"
            result = run_script(
                "recalculate_excel_formulas.py",
                ["--input-path", str(source), "--output-path", str(output)],
                env={"OFFICE_CONVERT_BACKEND": "com"},
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            self.assertEqual(data["backend"], "com")
            self.assertIn(data["calculation"], ("CalculateFullRebuild", "Calculate"))
            self.assertTrue(output.is_file())

            from openpyxl import load_workbook

            workbook = load_workbook(output, data_only=True)
            worksheet = workbook["数据"]
            self.assertEqual(worksheet["C2"].value, 3)
            self.assertEqual(worksheet["C3"].value, 7)
            self.assertEqual(wait_for_new_office_processes(before), set())

    def test_no_backend_error_lists_three_options(self) -> None:
        with tempfile.TemporaryDirectory(prefix="office_writer_no_backend_") as tmp:
            work = Path(tmp)
            for script, args in (
                (
                    "convert_to_pdf.py",
                    ["--input-path", str(FIXTURES / "sample.docx"),
                     "--output-path", str(work / "none.pdf")],
                ),
                (
                    "render_document_pages.py",
                    ["--input-path", str(FIXTURES / "sample.docx"),
                     "--output-dir", str(work / "pages")],
                ),
                (
                    "recalculate_excel_formulas.py",
                    ["--input-path", str(FIXTURES / "sample.xlsx"),
                     "--output-path", str(work / "none.xlsx")],
                ),
            ):
                with self.subTest(script=script):
                    result = run_script(
                        script, args, env={"OFFICE_CONVERT_BACKEND": "none"}
                    )
                    self.assertEqual(result.returncode, 3, result.stderr)
                    error = json.loads(result.stdout)["error"]
                    for marker in (
                        "npm i -g dsoffice", "Microsoft Office", "WPS Office", "pywin32",
                    ):
                        self.assertIn(marker, error)

            result = run_script(
                "convert_to_pdf.py",
                ["--input-path", str(FIXTURES / "sample.docx"),
                 "--output-path", str(work / "bogus.pdf")],
                env={"OFFICE_CONVERT_BACKEND": "bogus"},
            )
            self.assertEqual(result.returncode, 3, result.stderr)
            self.assertIn("auto/dsoffice/com/none", json.loads(result.stdout)["error"])

    def test_ppt_reflow_geometries(self) -> None:
        def snapshot(path: Path) -> dict:
            from pptx import Presentation

            presentation = Presentation(str(path))
            values = {}
            for slide_index, slide in enumerate(presentation.slides, start=1):
                for shape in slide.shapes:
                    fill = None
                    try:
                        fill = str(shape.fill.fore_color.rgb)
                    except Exception:
                        pass
                    runs = []
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                color = None
                                try:
                                    color = str(run.font.color.rgb)
                                except Exception:
                                    pass
                                runs.append((
                                    run.text, str(run.font.size), run.font.bold,
                                    run.font.name, color,
                                ))
                    text = shape.text if shape.has_text_frame else ""
                    values[(slide_index, shape.shape_id)] = (text, fill, tuple(runs))
            return values

        cases = {
            "out_of_bounds": [
                {"shape_type": "rect", "text": "OOB", "x": 9, "y": 6,
                 "width": 3, "height": 2, "bg_color": "2563EB"},
            ],
            "overlap": [
                {"shape_type": "rect", "text": "A", "x": 1, "y": 1,
                 "width": 4, "height": 2, "bg_color": "2563EB"},
                {"shape_type": "rect", "text": "B", "x": 3, "y": 2,
                 "width": 4, "height": 2, "bg_color": "EF4444"},
            ],
            "mixed": [
                {"shape_type": "rect", "text": "A", "x": 1, "y": 1,
                 "width": 4, "height": 2, "bg_color": "2563EB"},
                {"shape_type": "rect", "text": "B", "x": 3, "y": 2,
                 "width": 4, "height": 2, "bg_color": "EF4444"},
                {"shape_type": "rect", "text": "C", "x": 9, "y": 6,
                 "width": 3, "height": 2, "bg_color": "22C55E"},
            ],
            "oversize": [
                {"shape_type": "rect", "text": "BIG", "x": 0, "y": 0,
                 "width": 12, "height": 9, "bg_color": "F59E0B"},
            ],
        }
        with tempfile.TemporaryDirectory(prefix="office_writer_reflow_") as tmp:
            work = Path(tmp)
            for name, shapes in cases.items():
                with self.subTest(case=name):
                    deck = work / f"{name}.pptx"
                    params = {"slides": [{
                        "layout": "blank",
                        "body": [{"type": "shape", "shape": shape} for shape in shapes],
                    }]}
                    result = run_script(
                        "create_powerpoint.py",
                        ["--output-path", str(deck), "--params", json_arg(params)],
                    )
                    self.assertEqual(result.returncode, 0, result.stderr)
                    original_bytes = deck.read_bytes()
                    original_snapshot = snapshot(deck)

                    audit = run_script(
                        "ppt_optimize_slide_layout.py", ["--file-path", str(deck)]
                    )
                    self.assertEqual(audit.returncode, 0, audit.stderr)
                    self.assertGreater(
                        json.loads(audit.stdout)["data"]["warning_count"], 0
                    )
                    self.assertEqual(deck.read_bytes(), original_bytes)

                    result = run_script(
                        "ppt_optimize_slide_layout.py",
                        ["--file-path", str(deck), "--reflow"],
                    )
                    self.assertEqual(result.returncode, 0, result.stderr)
                    data = json.loads(result.stdout)["data"]
                    self.assertGreaterEqual(data["reflow_count"], 1)
                    self.assertEqual(data["remaining_warning_count"], 0)
                    self.assertTrue(data["reflow"])
                    for item in data["reflow"]:
                        self.assertIn(
                            item["action"], ("move", "resize", "move_resize")
                        )
                        self.assertIn("shape_id", item)
                        self.assertIn("from", item)
                        self.assertIn("to", item)
                    self.assertNotEqual(deck.read_bytes(), original_bytes)
                    self.assertEqual(snapshot(deck), original_snapshot)

                    audit_after = run_script(
                        "ppt_optimize_slide_layout.py", ["--file-path", str(deck)]
                    )
                    self.assertEqual(audit_after.returncode, 0, audit_after.stderr)
                    self.assertEqual(
                        json.loads(audit_after.stdout)["data"]["warning_count"], 0
                    )
                    if name == "oversize":
                        actions = {item["action"] for item in data["reflow"]}
                        self.assertTrue({"resize", "move_resize"} & actions)

    def test_scanned_pdf_ocr(self) -> None:
        tesseract = find_tesseract()
        if not tesseract:
            self.skipTest("未找到 Tesseract，跳过扫描件 OCR 实测")
        from PIL import Image, ImageDraw, ImageFont

        with tempfile.TemporaryDirectory(prefix="office_writer_ocr_") as tmp:
            work = Path(tmp)
            image = Image.new("RGB", (1600, 600), "white")
            draw = ImageDraw.Draw(image)
            try:
                font = ImageFont.truetype("arial.ttf", 96)
            except OSError:
                font = ImageFont.load_default()
            draw.text((80, 220), "SCANNED INVOICE 2026", fill="black", font=font)
            pdf_path = work / "scan.pdf"
            image.save(pdf_path, "PDF", resolution=150.0)

            result = run_script(
                "read_pdf_document.py",
                ["--file-path", str(pdf_path)],
                env={"OFFICE_TESSERACT_CMD": tesseract, "MCP_OFFICE_OCR_LANG": "eng"},
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads(result.stdout)["data"]
            self.assertIn("OCR", data["result"]["message"])
            text = (data["markdown"] + data["result"].get("content", "")).upper()
            self.assertTrue(
                any(token in text for token in ("SCANNED", "INVOICE", "2026")),
                text[:500],
            )


if __name__ == "__main__":
    unittest.main()
