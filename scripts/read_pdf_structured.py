#!/usr/bin/env python3
"""读取 PDF 的结构化内容（元数据、书签、链接、逐页文本与文本块）。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.reader import read_pdf_structured

TOOL = "read_pdf_structured"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "读取 PDF 元数据、书签、链接与逐页结构化文本。")
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--max-chars", dest="max_chars")
    parser.add_argument("--include-blocks", dest="include_blocks", action="store_true", default=argparse.SUPPRESS)
    parser.add_argument("--extract-images", dest="extract_images", action="store_true", default=argparse.SUPPRESS)
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    max_chars = common.int_arg(args, params, "max_chars", 50000)
    if max_chars < 1:
        raise common.UsageError("max_chars 必须大于 0")
    include_blocks = common.bool_arg(args, params, "include_blocks", False)
    extract_images = common.bool_arg(args, params, "extract_images", False)
    result = read_pdf_structured(
        file_path=file_path,
        max_chars=max_chars,
        include_blocks=include_blocks,
        extract_images=extract_images,
    )
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
