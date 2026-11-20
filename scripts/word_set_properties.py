#!/usr/bin/env python3
"""读取或写入 Word 文档核心属性。"""

from __future__ import annotations

import argparse

from docx import Document

import _common as common
import _native as native
from office_engine import utils
from office_engine.models import WordProperties
from office_engine.word_writer import _apply_core_properties

TOOL = "word_set_properties"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "读取或写入 Word 核心属性（标题/主题/作者/关键词/创建时间等）；不支持自定义属性。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--list", dest="list", action="store_true", default=None)
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    want_list = common.bool_arg(args, params, "list", False)
    raw_properties = params.get("properties")
    if isinstance(raw_properties, dict):
        unsupported = sorted(set(raw_properties) - set(WordProperties.model_fields))
        if unsupported:
            raise common.UsageError(
                "不支持的自定义属性字段：" + "、".join(unsupported)
                + "（仅支持核心属性）"
            )
    properties = common.model_arg(params, "properties", WordProperties, required=False)

    try:
        utils.validate_input_file(file_path, (".docx",))
    except ValueError as exc:
        raise common.ToolRuntimeError(str(exc)) from exc
    try:
        doc = Document(file_path)
    except Exception as exc:
        raise common.ToolRuntimeError(f"无法打开 Word 文档: {file_path}（{exc}）") from exc

    def _current_properties() -> dict:
        """读取当前核心属性，时间统一为 ISO 8601。"""
        cp = doc.core_properties
        return {
            "title": cp.title,
            "subject": cp.subject,
            "author": cp.author,
            "keywords": cp.keywords,
            "comments": cp.comments,
            "category": cp.category,
            "last_modified_by": cp.last_modified_by,
            "created": cp.created.isoformat() if cp.created else None,
            "modified": cp.modified.isoformat() if cp.modified else None,
        }

    updates = properties.model_dump(exclude_none=True) if properties else {}
    if want_list or not updates:
        return {
            "success": True,
            "file_path": file_path,
            "properties": _current_properties(),
            "backup_path": None,
        }, 0

    try:
        _apply_core_properties(doc, properties)
    except ValueError as exc:
        raise common.UsageError(str(exc)) from exc
    backup_path = utils.atomic_save(doc.save, file_path)
    return {
        "success": True,
        "file_path": file_path,
        "properties": _current_properties(),
        "updated_fields": sorted(updates),
        "backup_path": backup_path,
    }, 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
