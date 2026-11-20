#!/usr/bin/env python3
"""用 dsoffice/LibreOffice 将 Office 文档转换为 PDF。"""

from __future__ import annotations

import argparse
from pathlib import Path

import _common as common
import _native as native
from office_engine.office_convert import convert_to_pdf

TOOL = "convert_to_pdf"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "将 Office 文档高保真转换为 PDF（内置 dsoffice/LibreOffice 引擎）。")
    parser.add_argument("--input-path", dest="input_path")
    parser.add_argument("--output-path", dest="output_path")
    common.add_overwrite_flags(parser)
    parser.add_argument("--sheet")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    input_path = common.arg(args, params, "input_path", required=True)
    output_path = common.arg(args, params, "output_path")
    requested = output_path or str(Path(input_path).with_suffix(".pdf"))
    if_exists = common.output_policy(args, params, requested)
    sheet = common.arg(args, params, "sheet")
    timeout_result = {
        "success": False,
        "file_path": output_path or input_path,
        "message": (
            f"PDF 转换超时（>{common.DSOFFICE_OPERATION_TIMEOUT}s），"
            "首次运行可能正在下载 dsoffice 引擎或启动 Office/WPS，可重试"
        ),
    }
    result = common.with_timeout(
        convert_to_pdf,
        common.DSOFFICE_OPERATION_TIMEOUT,
        timeout_result,
        input_path,
        output_path,
        if_exists,
        sheet,
    )
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
