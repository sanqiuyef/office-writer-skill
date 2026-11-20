#!/usr/bin/env python3
"""PPT SmartArt：列出布局/插入（office-native COM 后端）。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native

TOOL = "ppt_manage_smartart"
ACTIONS = {"layouts", "add"}


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "PPT SmartArt：layouts/add；需要 Windows + PowerPoint 与 C# office-native 后端。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--action", help="layouts/add")
    parser.add_argument("--slide", type=int, help="幻灯片序号（1 起）")
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    action = common.choice(common.arg(args, params, "action", "layouts"), ACTIONS, "action")
    params.update({"file_path": file_path, "action": action})
    slide = common.arg(args, params, "slide")
    if slide is not None:
        params["slide"] = slide
    return native.run_native("ppt-smartart", params)


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
