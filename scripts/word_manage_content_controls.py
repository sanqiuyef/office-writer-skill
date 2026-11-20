#!/usr/bin/env python3
"""Word 内容控件：文本/日期/下拉/复选框的增删改查（office-native 后端）。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native

TOOL = "word_manage_content_controls"
ACTIONS = {
    "add-text", "add-date", "add-dropdown", "add-checkbox",
    "list", "set", "remove",
}


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "Word 内容控件：add-text/add-date/add-dropdown/add-checkbox/list/set/remove；"
        "需要 C# office-native 后端（tools/office-native）。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--action")
    parser.add_argument("--tag", help="内容控件 tag")
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    action = common.choice(common.arg(args, params, "action", "list"), ACTIONS, "action")
    params.update({"file_path": file_path, "action": action})
    tag = common.arg(args, params, "tag")
    if tag is not None:
        params["tag"] = tag
    return native.run_native("word-content-controls", params)


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
