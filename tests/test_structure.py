"""Structural checks for the office-writer skill package."""

from __future__ import annotations

import ast
import json
import re
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
ENGINE = SCRIPTS / "office_engine"

ENTRY_SCRIPTS = {
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
    "replace_open_document_content.py",
    "word_manage_comments.py",
    "word_set_properties.py",
    "excel_add_chart.py",
    "ppt_set_hyperlinks.py",
    "word_manage_revisions.py",
    "word_manage_footnotes.py",
    "word_manage_content_controls.py",
    "ppt_manage_effects.py",
    "ppt_manage_smartart.py",
    "ppt_manage_shapes.py",
    "excel_manage_pivots.py",
    "excel_print.py",
}

REQUIRED_TOP_LEVEL = {
    "SKILL.md",
    "README.md",
    "LICENSE",
    "requirements.txt",
    "lifecycle.yaml",
    ".gitignore",
    ".env.example",
    "agents/openai.yaml",
    "assets/README.md",
    "config/defaults.json",
    "references/wps_com.md",
    "references/fonts_page_numbers.md",
    "references/tables_slides.md",
    "references/render_qa.md",
    "references/cli_args.md",
    "references/conversion.md",
    "scripts/_common.py",
    "scripts/office_engine/__init__.py",
    "templates/word_template.docx",
    "templates/excel_template.xlsx",
    "templates/powerpoint_template.pptx",
    "tests/cases.py",
    "tests/test_scripts.py",
    "tests/test_structure.py",
    "tests/test_triggers.py",
    "tests/test_parity.py",
    "tests/trigger_cases.json",
    "tests/README.md",
    "tests/expected_outputs/parity_list_color_schemes.json",
    "tests/expected_outputs/parity_read_txt.json",
    "tests/expected_outputs/parity_validate_word.json",
}


class StructureTest(unittest.TestCase):
    def test_frontmatter_and_length(self) -> None:
        text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        self.assertTrue(text.startswith("---\n"))
        parts = text.split("---", 2)
        front = parts[1].strip()
        keys = [line.split(":", 1)[0].strip() for line in front.splitlines() if line.strip()]
        self.assertEqual(keys, ["name", "description"])
        self.assertIn("name: office-writer", front)
        match = re.search(r"^description:\s*(.+)$", front, re.MULTILINE)
        self.assertIsNotNone(match)
        description = match.group(1)
        self.assertLessEqual(len(description.encode("utf-8")), 1024)
        self.assertIn("触发：", description)
        self.assertLessEqual(len(text.splitlines()), 300)
        self.assertLessEqual(len(parts[2].splitlines()), 300)

    def test_required_sections(self) -> None:
        text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        for heading in (
            "## 触发条件",
            "## 何时不要用",
            "## 能力→脚本索引",
            "## 常用示例",
            "## 任务→脚本组合",
            "## 输出与退出码",
            "## 铁律",
            "## references 索引",
        ):
            with self.subTest(heading=heading):
                self.assertIn(heading, text)

    def test_referenced_paths_exist(self) -> None:
        text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        referenced = {
            item.split()[0]
            for item in re.findall(
                r"`((?:scripts|references|assets|templates|config|tests)/[^`]+)`", text
            )
            if "<" not in item and "*" not in item and "{{" not in item
        }
        missing = [path for path in sorted(referenced) if not (ROOT / path).exists()]
        self.assertEqual(missing, [], f"SKILL.md 引用了不存在的路径: {missing}")

    def test_required_files(self) -> None:
        for path in sorted(REQUIRED_TOP_LEVEL):
            with self.subTest(path=path):
                self.assertTrue((ROOT / path).is_file(), path)
        for directory in ("agents", "assets", "config", "references", "scripts", "templates", "tests"):
            with self.subTest(directory=directory):
                self.assertTrue((ROOT / directory).is_dir(), directory)

    def test_entry_scripts(self) -> None:
        on_disk = {path.name for path in SCRIPTS.glob("*.py") if not path.name.startswith("_")}
        self.assertEqual(on_disk, ENTRY_SCRIPTS)
        for name in sorted(ENTRY_SCRIPTS):
            with self.subTest(script=name):
                tree = ast.parse((SCRIPTS / name).read_text(encoding="utf-8"))
                functions = {
                    node.name for node in tree.body if isinstance(node, ast.FunctionDef)
                }
                self.assertEqual(functions, {"build_parser", "run", "main"})

    def test_no_mcp_runtime_dependency(self) -> None:
        pattern = re.compile(r"(?m)^\s*(?:import\s+mcp\w*|from\s+mcp\w*\s+import)")
        paths = list(SCRIPTS.rglob("*.py")) + list((ROOT / "tests").rglob("*.py"))
        for path in paths:
            body = path.read_text(encoding="utf-8")
            with self.subTest(path=str(path.relative_to(ROOT))):
                self.assertIsNone(
                    pattern.search(body),
                    path.name,
                )
                self.assertNotIn("mcp" + "__", body, path.name)

    def test_vendored_headers(self) -> None:
        expected = {
            "models", "word_writer", "excel_writer", "ppt_writer", "reader",
            "markdown_writer", "pdf_writer", "office_convert", "document_ops",
            "template_merge", "guards", "utils", "checker", "scenes", "covers",
            "office_com",
        }
        # open_replace 是本地模块（借鉴 unpack→edit→zip 思路，活动文档走 COM），
        # 不是 vendored 副本，因此不要求 mcp_office_writer 文件头。
        local = {"open_replace"}
        modules = {path.stem for path in ENGINE.glob("*.py") if path.name != "__init__.py"}
        self.assertEqual(modules, expected | local)
        for name in sorted(expected):
            with self.subTest(module=name):
                lines = (ENGINE / f"{name}.py").read_text(encoding="utf-8").splitlines()[:3]
                joined = "\n".join(lines)
                self.assertIn(f"Vendored from mcp_office_writer/{name}.py", joined)
                self.assertIn("0.6.0, 2026-11", joined)
                self.assertIn("Modifications:", joined)
        self.assertEqual((ENGINE / "__init__.py").read_text(encoding="utf-8"), "")

    def test_no_hardcoded_absolute_paths(self) -> None:
        pattern = re.compile(r"(?<![A-Za-z0-9_])[A-Za-z]:[\\/]")
        for path in SCRIPTS.rglob("*.py"):
            with self.subTest(path=str(path.relative_to(ROOT))):
                self.assertIsNone(pattern.search(path.read_text(encoding="utf-8")), path.name)

    def test_references_size_and_headers(self) -> None:
        files = sorted((ROOT / "references").glob("*.md"))
        self.assertGreaterEqual(len(files), 5)
        for path in files:
            with self.subTest(file=path.name):
                self.assertLess(path.stat().st_size, 15 * 1024, path.name)
                head = "\n".join(path.read_text(encoding="utf-8").splitlines()[:8])
                self.assertIn("来源", head)
                self.assertIn("适用边界", head)

    def test_config_has_no_secrets(self) -> None:
        data = json.loads((ROOT / "config" / "defaults.json").read_text(encoding="utf-8"))

        def walk(value):
            if isinstance(value, dict):
                for key, child in value.items():
                    lowered = str(key).lower()
                    self.assertNotIn("api_key", lowered)
                    self.assertNotIn("secret", lowered)
                    self.assertNotIn("token", lowered)
                    walk(child)
            elif isinstance(value, list):
                for child in value:
                    walk(child)
            elif isinstance(value, str):
                self.assertFalse(value.startswith("sk-"), value)

        walk(data)

    def test_requirements_without_mcp(self) -> None:
        body = (ROOT / "requirements.txt").read_text(encoding="utf-8")
        for line in body.splitlines():
            stripped = line.strip()
            if stripped and not stripped.startswith("#"):
                self.assertFalse(
                    re.match(r"mcp(?:\s|[<>=;]|$)", stripped, re.IGNORECASE),
                    stripped,
                )
        for package in ("python-docx", "openpyxl", "python-pptx", "pydantic", "pymupdf", "fpdf2", "pywin32", "pytesseract", "Pillow", "pypdf", "xlsxwriter"):
            self.assertIn(package, body)
        self.assertIn("AGPL", body)
        self.assertIn("LGPL-3.0", body)

    def test_templates_are_valid_zip_packages(self) -> None:
        for name in ("word_template.docx", "excel_template.xlsx", "powerpoint_template.pptx"):
            with self.subTest(template=name):
                self.assertTrue(zipfile.is_zipfile(ROOT / "templates" / name), name)

    def test_license_notices(self) -> None:
        body = (ROOT / "LICENSE").read_text(encoding="utf-8")
        self.assertIn("MIT License", body)
        notices = (ROOT / "THIRD_PARTY_NOTICES.md").read_text(encoding="utf-8")
        for marker in ("Third-party notices", "PyMuPDF", "AGPL-3.0", "fpdf2", "LGPL-3.0"):
            self.assertIn(marker, notices)

    def test_readme_length(self) -> None:
        lines = (ROOT / "README.md").read_text(encoding="utf-8").splitlines()
        self.assertLessEqual(len(lines), 100)


if __name__ == "__main__":
    unittest.main()
