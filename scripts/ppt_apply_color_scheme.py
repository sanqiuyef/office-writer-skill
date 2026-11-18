#!/usr/bin/env python3
"""对已有 PPT 应用指定配色方案。"""

from __future__ import annotations

import argparse

import _common as common
from office_engine.ppt_writer import apply_color_to_slides, list_color_schemes

TOOL = "ppt_apply_color_scheme"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "对已有 PPT 应用指定配色方案（标题/正文/背景/形状/表格）。")
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--scheme-name", dest="scheme_name")
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    scheme_name = str(common.arg(args, params, "scheme_name", "professional-blue"))
    valid_schemes = {item["name"] for item in list_color_schemes()}
    if scheme_name not in valid_schemes:
        raise common.UsageError(
            f"未知配色方案: {scheme_name}；可用: {', '.join(sorted(valid_schemes))}"
        )
    result = apply_color_to_slides(file_path=file_path, scheme_name=scheme_name)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
