#!/usr/bin/env python3
"""向已有 Excel 工作簿的指定工作表插入图表。"""

from __future__ import annotations

import argparse

from openpyxl import load_workbook

import _common as common
import _native as native
from office_engine.excel_writer import _close_workbook, add_chart_to_sheet
from office_engine.models import ExcelChartDef
from office_engine.utils import atomic_save, validate_input_file

TOOL = "excel_add_chart"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "向已有 Excel (.xlsx/.xlsm) 插入图表；支持 column/bar/line/pie/doughnut/area/scatter/radar。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--sheet", help="目标工作表名；省略时使用活动工作表")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    chart_def = common.model_arg(params, "chart", ExcelChartDef)
    sheet_name = common.arg(args, params, "sheet")

    try:
        validate_input_file(file_path, (".xlsx", ".xlsm"))
    except ValueError as exc:
        raise common.ToolRuntimeError(str(exc)) from exc

    try:
        wb = load_workbook(
            file_path,
            keep_vba=file_path.lower().endswith(".xlsm"),
            rich_text=True,
        )
    except Exception as exc:
        raise common.ToolRuntimeError(f"无法打开 Excel 文件 {file_path}: {exc}") from exc
    try:
        if sheet_name is None:
            ws = wb.active
            if ws is None:
                raise common.UsageError("工作簿没有可用的工作表")
        else:
            if sheet_name not in wb.sheetnames:
                raise common.UsageError(f"工作表不存在: {sheet_name}")
            ws = wb[sheet_name]
        try:
            chart_info = add_chart_to_sheet(ws, chart_def, wb)
        except ValueError as exc:
            raise common.UsageError(str(exc)) from exc
        sheet_title = ws.title
        backup_path = atomic_save(wb.save, file_path)
    finally:
        _close_workbook(wb)

    data = {
        "success": True,
        "file_path": file_path,
        "sheet": sheet_title,
        "chart_type": chart_info["chart_type"],
        "anchor": chart_info["anchor"],
        "backup_path": backup_path,
    }
    if chart_info.get("warnings"):
        data["warnings"] = chart_info["warnings"]
    return data, 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
