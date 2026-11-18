#!/usr/bin/env python3
"""将 Office/PDF 文档逐页渲染为 PNG。"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import _common as common
from office_engine.office_convert import render_document_pages

TOOL = "render_document_pages"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "将 Office/PDF 逐页渲染为 PNG（支持页码、工作表区域与 DPI）。")
    parser.add_argument("--input-path", dest="input_path")
    parser.add_argument("--output-dir", dest="output_dir")
    parser.add_argument("--pages")
    parser.add_argument("--dpi")
    parser.add_argument("--sheet")
    parser.add_argument("--cell-range", dest="cell_range")
    parser.add_argument("--if-exists", dest="if_exists")
    return parser


def run(args):
    params = common.parse_params(args)
    input_path = common.arg(args, params, "input_path", required=True)
    output_dir = common.arg(args, params, "output_dir")
    pages = str(common.arg(args, params, "pages", "all"))
    dpi = common.int_arg(args, params, "dpi", 144)
    sheet = common.arg(args, params, "sheet")
    cell_range = common.arg(args, params, "cell_range")
    if_exists = common.choice(
        common.arg(args, params, "if_exists", "overwrite"),
        {"overwrite", "error"},
        "if_exists",
    )
    if pages.lower() != "all" and not re.fullmatch(r"\d+(?:\s*,\s*\d+)*", pages):
        raise common.UsageError("pages 必须是 'all' 或 '1,3,5' 样式")
    if cell_range and not sheet:
        raise common.UsageError("cell_range 需要与 sheet 同时指定")
    if input_path.lower().endswith((".xlsx", ".xls")) and not (sheet and cell_range):
        raise common.UsageError("Excel 渲染需要同时指定 --sheet 与 --cell-range（如 --sheet 数据 --cell-range A1:D20）")
    timeout_result = {
        "success": False,
        "output_dir": output_dir or "",
        "message": (
            f"页面渲染超时（>{common.DSOFFICE_OPERATION_TIMEOUT}s），"
            "首次运行可能正在下载引擎或启动 Office/WPS，可重试"
        ),
    }
    result = common.with_timeout(
        render_document_pages,
        common.DSOFFICE_OPERATION_TIMEOUT,
        timeout_result,
        input_path,
        output_dir,
        pages,
        dpi,
        sheet,
        cell_range,
        if_exists,
    )
    common.ensure_success(result)
    data = common.result_data(result)
    # dsoffice 的 pageCount 是上游元数据（例如 Excel 按工作表计），可能与实际 PNG 数不一致。
    # 以输出目录中的真实 PNG 产物为准修正 page_count，并保留上游原值便于排查。
    reported_count = data.get("page_count")
    actual_count = None
    candidate_dir = data.get("output_dir") or output_dir
    if candidate_dir:
        output_path = Path(candidate_dir)
        if output_path.is_dir():
            actual_count = sum(
                1 for item in output_path.iterdir()
                if item.is_file() and item.suffix.lower() == ".png"
            )
    if actual_count is None:
        actual_count = len(data.get("images") or [])
    data["page_count"] = actual_count
    if isinstance(reported_count, int) and reported_count != actual_count:
        data["page_count_reported"] = reported_count
        data["message"] = (
            f"渲染完成，共 {actual_count} 页"
            f"（上游元数据报告 {reported_count} 页，已按实际产物修正）"
        )
    else:
        data["message"] = f"渲染完成，共 {actual_count} 页"
    return data, 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
