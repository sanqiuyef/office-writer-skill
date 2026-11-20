#!/usr/bin/env python3
"""向已有 PPT 追加幻灯片。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.models import SlideDef
from office_engine.ppt_writer import edit_powerpoint

TOOL = "edit_powerpoint"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "向已有 PPT 追加幻灯片（优先尝试活动 COM 演示文稿）。")
    parser.add_argument("--file-path", dest="file_path")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    append_slides = common.model_list_arg(params, "append_slides", SlideDef)
    result = edit_powerpoint(file_path, append_slides)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
