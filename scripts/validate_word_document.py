#!/usr/bin/env python3
"""验证 Word 文档质量；报告有 errors 时退出码为 1。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.checker import validate_word_document

TOOL = "validate_word_document"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "验证 Word 文档：空白页、行距、表格、标题层级/编号、字体安装、PAGE/TOC 域、空段落与中英文字体混用。")
    parser.add_argument("--file-path", dest="file_path")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    report = validate_word_document(file_path=file_path).to_dict()
    return report, 1 if report.get("errors", 0) > 0 else 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
