#!/usr/bin/env python3
"""读取 Office 文档节点树并返回稳定路径。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.document_ops import inspect_office_document

TOOL = "inspect_office_document"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "读取 DOCX/XLSX/PPTX 的节点树与稳定路径（供 batch_edit 使用）。")
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--max-nodes", dest="max_nodes")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    max_nodes = common.int_arg(args, params, "max_nodes", 1000)
    if not 1 <= max_nodes <= 10000:
        raise common.UsageError("max_nodes 必须在 1 到 10000 之间")
    result = inspect_office_document(file_path=file_path, max_nodes=max_nodes)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
