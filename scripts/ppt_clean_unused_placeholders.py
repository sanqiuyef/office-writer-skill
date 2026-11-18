#!/usr/bin/env python3
"""清理已有 PPT 中没有文字的空白占位符。"""

from __future__ import annotations

import argparse

import _common as common
from office_engine.ppt_writer import clean_unused_placeholders

TOOL = "ppt_clean_unused_placeholders"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "清理已有 PPT 中没有文字的空白占位符。")
    parser.add_argument("--file-path", dest="file_path")
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    result = clean_unused_placeholders(file_path=file_path)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
