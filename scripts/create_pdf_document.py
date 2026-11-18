#!/usr/bin/env python3
"""创建 PDF 文档。"""

from __future__ import annotations

import argparse

import _common as common
from office_engine.models import PdfContent, PdfFontConfig, PdfPageSetup
from office_engine.pdf_writer import create_pdf_document

TOOL = "create_pdf_document"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "创建 PDF；支持标题、段落、表格、图片、列表、分页与中文字体。")
    parser.add_argument("--output-path", dest="output_path")
    parser.add_argument("--if-exists", dest="if_exists")
    return parser


def run(args):
    params = common.parse_params(args)
    output_path = common.arg(args, params, "output_path", required=True)
    content = common.model_arg(params, "content", PdfContent)
    page_setup = common.model_arg(params, "page_setup", PdfPageSetup, required=False)
    font_config = common.model_arg(params, "font_config", PdfFontConfig, required=False)
    if_exists = common.choice(
        common.arg(args, params, "if_exists", "overwrite"),
        {"overwrite", "error", "rename"},
        "if_exists",
    )
    result = create_pdf_document(output_path, content, page_setup, font_config, if_exists)
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
