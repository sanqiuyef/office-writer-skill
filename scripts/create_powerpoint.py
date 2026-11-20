#!/usr/bin/env python3
"""创建 PowerPoint 演示文稿。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.models import SlideDef
from office_engine.ppt_writer import create_powerpoint

TOOL = "create_powerpoint"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "创建 PowerPoint；支持布局、段落、表格、图片、形状、图表与备注。")
    parser.add_argument("--output-path", dest="output_path")
    parser.add_argument("--template")
    common.add_overwrite_flags(parser)
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    output_path = common.arg(args, params, "output_path", required=True)
    if params.get("sections") or params.get("show_settings") or any(
            isinstance(item, dict) and (item.get("footer_text") or item.get("show_slide_number"))
            for item in (params.get("slides") or [])):
        raise common.UsageError(
            "sections/show_settings/footer_text/show_slide_number 仅 C# 后端支持；"
            "请去掉 OFFICE_NATIVE_ENGINE=python"
        )
    slides = common.model_list_arg(params, "slides", SlideDef)
    template = common.arg(args, params, "template")
    if_exists = common.output_policy(args, params, output_path)
    result = create_powerpoint(output_path, slides, template, if_exists)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
