#!/usr/bin/env python3
"""读取 PowerPoint 演示文稿并返回 Markdown 文本与结构化统计。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.reader import read_powerpoint

TOOL = "read_powerpoint"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "读取 PowerPoint (.pptx)；输出 markdown 与上游结构化 result。")
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--max-chars", dest="max_chars")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    max_chars = common.int_arg(args, params, "max_chars", 50000)
    if max_chars < 1:
        raise common.UsageError("max_chars 必须大于 0")
    result = read_powerpoint(file_path=file_path, max_chars=max_chars)
    return common.read_payload(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
