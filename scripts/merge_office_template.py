#!/usr/bin/env python3
"""用 JSON 数据填充 Office 模板中的 {{key}} 占位符。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.template_merge import merge_office_template

TOOL = "merge_office_template"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "用 JSON 数据填充 DOCX/XLSX/PPTX 模板中的 {{key}} 占位符。")
    parser.add_argument("--template-path", dest="template_path")
    parser.add_argument("--output-path", dest="output_path")
    common.add_overwrite_flags(parser)
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    template_path = common.arg(args, params, "template_path", required=True)
    output_path = common.arg(args, params, "output_path", required=True)
    if_exists = common.output_policy(args, params, output_path)
    data = params.get("data")
    if not isinstance(data, dict):
        raise common.UsageError("缺少参数 data（通过 --params 或 --params-file 传入 JSON 对象）")
    result = merge_office_template(template_path, output_path, data, if_exists)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
