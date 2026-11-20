#!/usr/bin/env python3
"""创建 Word 文档。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.models import CoverConfig, FontConfig, HeaderFooter, PageSetup, WordContent
from office_engine.word_writer import create_word_document

TOOL = "create_word_document"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "创建 Word；支持模板、场景排版、页眉页脚、封面与多种内容块。")
    parser.add_argument("--output-path", dest="output_path")
    parser.add_argument("--template")
    parser.add_argument("--scene")
    common.add_overwrite_flags(parser)
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    output_path = common.arg(args, params, "output_path", required=True)
    content = common.model_arg(params, "content", WordContent)
    template = common.arg(args, params, "template")
    page_setup = common.model_arg(params, "page_setup", PageSetup, required=False)
    header_footer = common.model_arg(params, "header_footer", HeaderFooter, required=False)
    default_font = common.model_arg(params, "default_font", FontConfig, required=False)
    scene = common.arg(args, params, "scene")
    cover = common.model_arg(params, "cover_config", CoverConfig, required=False)
    cover_config = cover.model_dump(exclude_none=True) if cover else None
    if_exists = common.output_policy(args, params, output_path)
    result = create_word_document(
        output_path,
        content,
        template,
        page_setup,
        header_footer,
        default_font,
        scene,
        cover_config,
        if_exists,
    )
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
