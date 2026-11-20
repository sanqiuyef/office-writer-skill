"""Shared bridge to the optional C# office-native backend."""

from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path

import _common as common

ENGINE_ENV = "OFFICE_NATIVE_ENGINE"
COMMANDS = {
    "create_word_document": "create-word-document",
    "edit_word_document": "edit-word-document",
    "create_word_from_markdown": "create-word-from-markdown",
    "read_word_document": "read-word-document",
    "validate_word_document": "validate-word-document",
    "edit_open_word_document": "edit-open-word-document",
    "word_manage_comments": "word-manage-comments",
    "word_set_properties": "word-set-properties",
    "word_manage_revisions": "word-manage-revisions",
    "word_manage_footnotes": "word-manage-footnotes",
    "word_manage_content_controls": "word-manage-content-controls",
    "create_excel_workbook": "create-excel-workbook",
    "edit_excel_workbook": "edit-excel-workbook",
    "read_excel_workbook": "read-excel-workbook",
    "recalculate_excel_formulas": "recalculate-excel-formulas",
    "edit_open_excel_workbook": "edit-open-excel-workbook",
    "excel_add_chart": "excel-add-chart",
    "excel_manage_pivots": "excel-manage-pivots",
    "excel_print": "excel-print",
    "create_powerpoint": "create-powerpoint",
    "edit_powerpoint": "edit-powerpoint",
    "read_powerpoint": "read-powerpoint",
    "ppt_list_color_schemes": "ppt-list-color-schemes",
    "ppt_apply_color_scheme": "ppt-apply-color-scheme",
    "ppt_optimize_slide_layout": "ppt-optimize-slide-layout",
    "ppt_clean_unused_placeholders": "ppt-clean-unused-placeholders",
    "ppt_analyze_template": "ppt-analyze-template",
    "ppt_create_framework_diagram": "ppt-create-framework-diagram",
    "ppt_add_flowchart_slide": "ppt-add-flowchart-slide",
    "edit_open_powerpoint": "edit-open-powerpoint",
    "ppt_set_hyperlinks": "ppt-set-hyperlinks",
    "ppt_manage_effects": "ppt-manage-effects",
    "ppt_manage_smartart": "ppt-manage-smartart",
    "create_pdf_document": "create-pdf-document",
    "read_pdf_document": "read-pdf-document",
    "read_pdf_structured": "read-pdf-structured",
    "merge_office_template": "merge-office-template",
    "batch_edit_office_document": "batch-edit-office-document",
    "convert_to_pdf": "convert-to-pdf",
    "render_document_pages": "render-document-pages",
    "inspect_office_document": "inspect-office-document",
    "get_open_office_documents": "get-open-office-documents",
    "replace_open_document_content": "replace-open-document-content",
    "read_txt_document": "read-txt-document",
}

SKILL_DIR = Path(__file__).resolve().parent.parent
BUILD_HINT = (
    "office-native 后端未构建。请在 skill 根目录运行："
    "pwsh -File tools/office-native/build.ps1；"
    "或设置 OFFICE_NATIVE_BIN 指向已编译的 office-native.exe"
)


def _candidates():
    env = os.environ.get("OFFICE_NATIVE_BIN")
    if env:
        yield Path(env)
    yield SKILL_DIR / "tools/office-native/dist/office-native.exe"
    yield SKILL_DIR / "tools/office-native/bin/Release/net8.0/office-native.exe"
    yield SKILL_DIR / "tools/office-native/bin/Debug/net8.0/office-native.exe"


IMPLEMENTED = set(COMMANDS)


def use_csharp(tool: str | None = None) -> bool:
    """默认使用 C# 后端；OFFICE_NATIVE_ENGINE=python 强制回旧 Python 引擎。"""
    value = os.environ.get(ENGINE_ENV, "").strip().lower()
    if value in {"python", "legacy", "py"}:
        return False
    if value in {"", "csharp", "native", "1", "true"}:
        return tool is None or tool in IMPLEMENTED
    return False


def run_entry(tool: str, params: dict) -> tuple[dict, int]:
    command = COMMANDS.get(tool)
    if command is None:
        raise common.ToolRuntimeError(f"没有对应的 C# 命令: {tool}")
    try:
        timeout = 60 if tool == "get_open_office_documents" else None
        return run_native(command, params, timeout=timeout)
    except common.ToolRuntimeError as exc:
        # 活动文档枚举在 Python 端是 60s 超时返回空列表；保持该降级契约。
        if tool == "get_open_office_documents" and "超时" in str(exc):
            return {"word": [], "excel": [], "powerpoint": []}, 0
        raise


def run_entry_from_args(tool: str, args, params: dict) -> tuple[dict, int]:
    """Merge explicit argparse flags over --params, then dispatch to C#."""
    merged = dict(params)
    for key, value in vars(args).items():
        if key in {"params", "params_file"} or value is None:
            continue
        # store_true 的默认 False 不能覆盖 --params 里显式传的 True。
        if value is False and merged.get(key):
            continue
        merged[key] = value
    mode = merged.get("if_exists")
    if mode == "overwrite":
        merged["overwrite"] = True
    elif mode == "rename":
        merged["rename"] = True
        output = merged.get("output_path")
        if output:
            path = Path(output)
            if path.exists() and (path.is_dir() or path.stat().st_size > 0):
                counter = 1
                while True:
                    candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
                    if not candidate.exists():
                        break
                    counter += 1
                merged["output_path"] = str(candidate)
    return run_entry(tool, merged)


def native_binary() -> Path:
    for candidate in _candidates():
        if candidate.is_file():
            return candidate
    raise common.ToolRuntimeError(BUILD_HINT)


def run_native(command: str, params: dict, timeout: int | None = None) -> tuple[dict, int]:
    """Run one office-native command and map its JSON/exit code to the CLI contract."""
    binary = native_binary()
    if timeout is None:
        timeout = int(common.config_get("timeouts.native_operation_seconds", 300))
    try:
        proc = subprocess.run(
            # stdin 传参避免 Windows 32K 命令行长度上限
            [str(binary), command, "--params-file", "-"],
            input=json.dumps(params, ensure_ascii=False),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="backslashreplace",
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise common.ToolRuntimeError(f"office-native 执行超时（>{timeout}s）") from exc
    except OSError as exc:
        raise common.ToolRuntimeError(f"office-native 无法启动: {exc}") from exc

    stdout = (proc.stdout or "").strip()
    try:
        payload = json.loads(stdout) if stdout else None
    except ValueError as exc:
        detail = stdout[:300] or (proc.stderr or "").strip()[:300]
        raise common.ToolRuntimeError(
            f"office-native 输出无法解析（exit={proc.returncode}）: {detail}"
        ) from exc
    if not isinstance(payload, dict):
        raise common.ToolRuntimeError(
            f"office-native 输出不是 JSON 对象（exit={proc.returncode}）"
        )
    if proc.returncode == 2:
        raise common.UsageError(str(payload.get("error") or "office-native 参数错误"))
    if proc.returncode == 1 and payload.get("ok"):
        data = payload.get("data")
        if isinstance(data, dict):
            return data, 1
    if proc.returncode != 0 or not payload.get("ok"):
        raise common.ToolRuntimeError(str(payload.get("error") or "office-native 执行失败"))
    data = payload.get("data")
    if not isinstance(data, dict):
        raise common.ToolRuntimeError("office-native 返回的 data 不是 JSON 对象")
    return data, 0
