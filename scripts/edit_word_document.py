#!/usr/bin/env python3
"""向已有 Word 文档追加内容块。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.models import FontConfig, SectionItem
from office_engine.word_writer import edit_word_document

TOOL = "edit_word_document"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "向已有 Word 文档末尾追加内容块（优先尝试活动 COM 文档）。")
    parser.add_argument("--file-path", dest="file_path")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    append_sections = common.model_list_arg(params, "append_sections", SectionItem)
    default_font = common.model_arg(params, "default_font", FontConfig, required=False)
    result = edit_word_document(file_path, append_sections, default_font)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
