#!/usr/bin/env python3
"""分析已有 PPT/PPTX 模板的尺寸、布局、字体与颜色。"""

from __future__ import annotations

import argparse

import _common as common
from office_engine.ppt_writer import analyze_powerpoint_template

TOOL = "ppt_analyze_template"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "分析 PPT/PPTX 模板的尺寸、布局、字体与颜色。")
    parser.add_argument("--file-path", dest="file_path")
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    result = analyze_powerpoint_template(file_path=file_path)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
