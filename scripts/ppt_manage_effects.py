#!/usr/bin/env python3
"""PPT 动画/切换/批注（office-native COM 后端）。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native

TOOL = "ppt_manage_effects"
ACTIONS = {
    "animations-add", "animations-list", "animations-clear",
    "transition-set", "transition-clear", "transition-list",
    "comments-add", "comments-list", "comments-delete",
}


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "PPT 动画/切换/批注：animations-*、transition-*、comments-*；"
        "需要 Windows + PowerPoint 与 C# office-native 后端。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--action")
    parser.add_argument("--slide", type=int, help="幻灯片序号（1 起）")
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    action = common.choice(
        common.arg(args, params, "action", "animations-list"), ACTIONS, "action"
    )
    params.update({"file_path": file_path, "action": action})
    slide = common.arg(args, params, "slide")
    if slide is not None:
        params["slide"] = slide
    return native.run_native("ppt-effects", params)


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
