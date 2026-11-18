#!/usr/bin/env python3
"""优化已有 PPT 的内容元素布局。"""

from __future__ import annotations

import argparse

import _common as common
from office_engine.ppt_writer import audit_slide_layout, optimize_slide_layout, reflow_slide_layout

TOOL = "ppt_optimize_slide_layout"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "检查 PPT 形状越界/重叠并返回 warnings；--reflow 自动重排并写回，--fix-spacing 按原逻辑优化纵向间距。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument(
        "--fix-spacing",
        dest="fix_spacing",
        action="store_true",
        help="写回文件并均匀分布带文字形状的垂直间距（原优化行为）",
    )
    parser.add_argument(
        "--reflow",
        dest="reflow",
        action="store_true",
        help="自动重排越界/重叠形状并写回（保留 0.1in 边距、0.15in 间距；只改几何）",
    )
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    if getattr(args, "reflow", False):
        result = reflow_slide_layout(file_path=file_path)
    elif getattr(args, "fix_spacing", False):
        result = optimize_slide_layout(file_path=file_path)
    else:
        result = audit_slide_layout(file_path=file_path)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
