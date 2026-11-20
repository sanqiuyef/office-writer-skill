#!/usr/bin/env python3
"""读取 Excel 工作簿并返回 Markdown 表格与结构化统计。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.reader import read_excel_workbook

TOOL = "read_excel_workbook"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "读取 Excel (.xlsx/.xlsm) 工作簿；输出 markdown 与上游结构化 result。")
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--max-rows-per-sheet", dest="max_rows_per_sheet")
    parser.add_argument("--max-chars", dest="max_chars")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    max_rows_per_sheet = common.int_arg(args, params, "max_rows_per_sheet", 1000)
    max_chars = common.int_arg(args, params, "max_chars", 50000)
    if max_rows_per_sheet < 1:
        raise common.UsageError("max_rows_per_sheet 必须大于 0")
    if max_chars < 1:
        raise common.UsageError("max_chars 必须大于 0")
    result = read_excel_workbook(
        file_path=file_path,
        max_rows_per_sheet=max_rows_per_sheet,
        max_chars=max_chars,
    )
    return common.read_payload(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
