#!/usr/bin/env python3
"""创建 Excel 工作簿。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.excel_writer import create_excel_workbook
from office_engine.models import SheetDef

TOOL = "create_excel_workbook"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "创建 Excel；支持多工作表、格式、合并、冻结与筛选。")
    parser.add_argument("--output-path", dest="output_path")
    parser.add_argument("--template")
    common.add_overwrite_flags(parser)
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    output_path = common.arg(args, params, "output_path", required=True)
    sheets = common.model_list_arg(params, "sheets", SheetDef)
    template = common.arg(args, params, "template")
    if_exists = common.output_policy(args, params, output_path)
    result = create_excel_workbook(output_path, sheets, template, if_exists)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
