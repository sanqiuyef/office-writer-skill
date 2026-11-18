#!/usr/bin/env python3
"""重算 Excel 公式并写回缓存值。"""

from __future__ import annotations

import argparse

import _common as common
from office_engine.office_convert import recalculate_excel_formulas

TOOL = "recalculate_excel_formulas"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "用 LibreOffice 引擎重算 Excel 公式并把缓存值写回。")
    parser.add_argument("--input-path", dest="input_path")
    parser.add_argument("--output-path", dest="output_path")
    parser.add_argument("--if-exists", dest="if_exists")
    return parser


def run(args):
    params = common.parse_params(args)
    input_path = common.arg(args, params, "input_path", required=True)
    output_path = common.arg(args, params, "output_path")
    if_exists = common.choice(
        common.arg(args, params, "if_exists", "overwrite"),
        {"overwrite", "error", "rename"},
        "if_exists",
    )
    timeout_result = {
        "success": False,
        "file_path": output_path or input_path,
        "message": f"公式重算超时（>{common.DSOFFICE_OPERATION_TIMEOUT}s），首次运行可能正在下载引擎或启动 Excel/WPS，可重试",
    }
    result = common.with_timeout(
        recalculate_excel_formulas,
        common.DSOFFICE_OPERATION_TIMEOUT,
        timeout_result,
        input_path,
        output_path,
        if_exists,
    )
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
