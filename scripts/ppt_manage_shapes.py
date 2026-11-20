#!/usr/bin/env python3
"""PPT 形状绘制：add/list（默认 Python；活动文档 COM 优先，文件级 python-pptx 回退）。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.models import SlideShapeDef
from office_engine.ppt_writer import add_shape_to_presentation, list_slide_shapes

TOOL = "ppt_manage_shapes"
ACTIONS = {"add", "list"}


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "PPT 形状绘制：add/list；文件已打开时走活动文档 COM，未打开时文件级写入并自动备份。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--action", help="add/list")
    parser.add_argument("--slide", type=int, help="幻灯片序号（1 起）")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    action = common.choice(common.arg(args, params, "action", "add"), ACTIONS, "action")
    slide = common.arg(args, params, "slide", 1)
    if not isinstance(slide, int) or isinstance(slide, bool):
        raise common.UsageError("slide 必须是整数（1 起）")
    try:
        if action == "add":
            raw = params.get("shape")
            if not isinstance(raw, dict):
                raise common.UsageError(
                    "shape 必须是 JSON 对象，至少包含 width 与 height（英寸）"
                )
            shape = SlideShapeDef.model_validate(raw)
            result = add_shape_to_presentation(file_path, slide, shape)
        else:
            result = list_slide_shapes(file_path, slide)
    except ValueError as exc:
        raise common.UsageError(str(exc)) from exc
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
