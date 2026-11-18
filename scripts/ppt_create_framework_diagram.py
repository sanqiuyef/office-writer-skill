#!/usr/bin/env python3
"""生成 PPT 框架图（阶段标签 + 节点 + 箭头）。"""

from __future__ import annotations

import argparse

import _common as common
from office_engine.models import FrameworkDiagram
from office_engine.ppt_writer import create_framework_diagram

TOOL = "ppt_create_framework_diagram"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "生成 PPT 框架图（阶段标签 + 节点 + 箭头，支持 12 套配色）。")
    parser.add_argument("--output-path", dest="output_path")
    parser.add_argument("--if-exists", dest="if_exists")
    return parser


def run(args):
    params = common.parse_params(args)
    output_path = common.arg(args, params, "output_path", required=True)
    diagram = common.model_arg(params, "diagram", FrameworkDiagram)
    if_exists = common.choice(
        common.arg(args, params, "if_exists", "overwrite"),
        {"overwrite", "error", "rename"},
        "if_exists",
    )
    result = create_framework_diagram(output_path, diagram, if_exists)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
