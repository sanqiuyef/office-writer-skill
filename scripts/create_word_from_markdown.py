#!/usr/bin/env python3
"""把磁盘上的 Markdown 长文转换为 Word。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.markdown_writer import create_word_from_markdown
from office_engine.models import CreateResult

TOOL = "create_word_from_markdown"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "把磁盘上的 Markdown 长文转换为 Word（适合报告/制度/论文）。")
    parser.add_argument("--source-path", dest="source_path")
    parser.add_argument("--output-path", dest="output_path")
    parser.add_argument("--template")
    parser.add_argument("--scene")
    common.add_overwrite_flags(parser)
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    source_path = common.arg(args, params, "source_path", required=True)
    output_path = common.arg(args, params, "output_path", required=True)
    template = common.arg(args, params, "template")
    scene = common.arg(args, params, "scene", "report")
    if_exists = common.output_policy(args, params, output_path)
    timeout_result = CreateResult(
        success=False,
        file_path=output_path,
        message=(
            f"Markdown 转 Word 超时（>{common.MARKDOWN_CONVERT_TIMEOUT}s），"
            "请检查源文件是否过大"
        ),
    )
    result = common.with_timeout(
        create_word_from_markdown,
        common.MARKDOWN_CONVERT_TIMEOUT,
        timeout_result,
        source_path,
        output_path,
        template,
        scene,
        if_exists,
    )
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
