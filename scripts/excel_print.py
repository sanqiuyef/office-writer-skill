#!/usr/bin/env python3
"""Excel 导出 PDF / 打印（office-native COM 后端）。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native

TOOL = "excel_print"
ACTIONS = {"pdf", "print"}


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "Excel 导出 PDF 或打印：pdf/print；print 需要显式 confirm_print=true。"
        "需要 Windows + Excel 与 C# office-native 后端。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--action", help="pdf/print")
    parser.add_argument("--sheet", help="只导出/打印指定工作表")
    parser.add_argument("--output", help="PDF 输出路径（action=pdf）")
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    action = common.choice(common.arg(args, params, "action", "pdf"), ACTIONS, "action")
    params.update({"file_path": file_path, "action": action})
    for name in ("sheet", "output"):
        value = common.arg(args, params, name)
        if value is not None:
            params[name] = value
    return native.run_native("excel-print", params)


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
