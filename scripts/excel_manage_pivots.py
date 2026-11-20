#!/usr/bin/env python3
"""Excel 数据透视表：创建/列出/刷新（office-native COM 后端）。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native

TOOL = "excel_manage_pivots"
ACTIONS = {"create", "list", "refresh"}


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "Excel 数据透视表：create/list/refresh；需要 Windows + Excel 与 C# office-native 后端。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--action", help="create/list/refresh")
    parser.add_argument("--sheet", help="目标/查询工作表")
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    action = common.choice(common.arg(args, params, "action", "list"), ACTIONS, "action")
    params.update({"file_path": file_path, "action": action})
    sheet = common.arg(args, params, "sheet")
    if sheet is not None:
        params["sheet"] = sheet
    return native.run_native("excel-pivots", params)


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
