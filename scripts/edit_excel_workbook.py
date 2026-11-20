#!/usr/bin/env python3
"""向已有 Excel 工作簿追加工作表。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.excel_writer import edit_excel_workbook
from office_engine.models import SheetDef

TOOL = "edit_excel_workbook"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "向已有 Excel 工作簿追加工作表（优先尝试活动 COM 工作簿）。")
    parser.add_argument("--file-path", dest="file_path")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    append_sheets = common.model_list_arg(params, "append_sheets", SheetDef)
    result = edit_excel_workbook(file_path, append_sheets)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
