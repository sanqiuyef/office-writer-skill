#!/usr/bin/env python3
"""列出 Windows COM 中已打开的 Word/WPS、Excel 与 PowerPoint 文档。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.office_com import get_open_office_documents

TOOL = "get_open_office_documents"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "列出活动 Word/WPS、Excel、PowerPoint 文档；60 秒超时返回空列表。")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    result = common.with_timeout(
        get_open_office_documents,
        common.COM_OPERATION_TIMEOUT,
        {"word": [], "excel": [], "powerpoint": []},
    )
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
