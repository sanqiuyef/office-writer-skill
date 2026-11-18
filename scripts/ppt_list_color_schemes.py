#!/usr/bin/env python3
"""列出所有可用的 PPT 预设配色方案。"""

from __future__ import annotations

import argparse

import _common as common
from office_engine.ppt_writer import list_color_schemes

TOOL = "ppt_list_color_schemes"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "列出 12 套 PPT 预设配色方案（名称、说明与预览色）。")
    return parser


def run(args):
    return {"schemes": list_color_schemes()}, 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
