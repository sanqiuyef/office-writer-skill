#!/usr/bin/env python3
"""Word 脚注：新增/列出（office-native 后端）。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native

TOOL = "word_manage_footnotes"
ACTIONS = {"add", "list"}


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "Word 脚注：add/list；需要 C# office-native 后端（tools/office-native）。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--action", help="add/list")
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    action = common.choice(common.arg(args, params, "action", "list"), ACTIONS, "action")
    params.update({"file_path": file_path, "action": action})
    return native.run_native("word-footnotes", params)


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
