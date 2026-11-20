# Vendored from mcp_office_writer/excel_writer.py (mcp-office-writer 0.6.0, 2026-11).
# Upstream path: mcp_office_writer/excel_writer.py (repository-relative).
# Modifications: added italic/vertical alignment/borders, column-format width and font fields, row heights,
# sheet visibility, charts, conditional formats, data validations, tables, defined names, hyperlinks,
# comments, images and page setup; shared helpers are reused by excel_add_chart.
# Also added sheet protection and multi-key row sorting (with autoFilter sortState) for SheetDef.

"""
mcp-office-writer - Excel (.xlsx) 写入/编辑逻辑
"""
from __future__ import annotations

import logging
import math
import re
from functools import cmp_to_key
from pathlib import Path
from typing import TYPE_CHECKING

from openpyxl import Workbook, load_workbook
from openpyxl.chart import (
    AreaChart,
    BarChart,
    DoughnutChart,
    LineChart,
    PieChart,
    RadarChart,
    ScatterChart,
)
from openpyxl.chart.reference import Reference
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
    range_to_tuple,
)
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.filters import AutoFilter, SortCondition, SortState
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableStyleInfo

if TYPE_CHECKING:
    from .models import (
        CellStyle,
        ConditionalFormatDef,
        CreateResult,
        DataValidationDef,
        DefinedNameDef,
        EditResult,
        ExcelChartDef,
        ExcelCommentDef,
        ExcelHyperlinkDef,
        ExcelImageDef,
        ExcelPageSetupDef,
        ExcelTableDef,
        SheetDef,
        SheetProtectionDef,
    )

logger = logging.getLogger(__name__)

# Excel 内置图例位置（openpyxl 使用 b/tr/l/r/t 短码）
_LEGEND_POSITIONS = {"right": "r", "bottom": "b", "top": "t", "left": "l"}
# openpyxl 没有 PAPERSIZE_B5 常量，OOXML 内置 B5 的编号是 13
_PAPER_SIZE_FALLBACK = {"B5": 13}
_CHART_CLASSES = {
    "line": LineChart,
    "pie": PieChart,
    "doughnut": DoughnutChart,
    "area": AreaChart,
    "scatter": ScatterChart,
    "radar": RadarChart,
}


def _close_workbook(workbook) -> None:
    """Release both the workbook and the optional macro archive it owns."""
    try:
        workbook.close()
    finally:
        vba_archive = getattr(workbook, "vba_archive", None)
        if vba_archive is not None:
            try:
                vba_archive.close()
            except (OSError, ValueError):
                # An archive may already have been closed by openpyxl internals.
                pass


def _build_font(style: CellStyle | None, default_size: int = 11) -> Font:
    if style is None:
        return Font(size=default_size)
    return Font(
        name=style.font_name,
        size=style.font_size or default_size,
        bold=style.bold,
        italic=style.italic,
        color=style.font_color,
    )


def _build_fill(style: CellStyle | None) -> PatternFill | None:
    if style and style.bg_color:
        return PatternFill(start_color=style.bg_color, end_color=style.bg_color, fill_type="solid")
    return None


def _build_alignment(style: CellStyle | None) -> Alignment | None:
    if style is None:
        return None
    h = style.alignment
    if h or style.vertical_alignment or style.wrap_text:
        return Alignment(horizontal=h, vertical=style.vertical_alignment, wrap_text=style.wrap_text)
    return None


def _build_border(style: CellStyle | None) -> Border | None:
    if style is None or style.border is None:
        return None
    spec = style.border
    sides = [Side(style=spec.style, color=spec.color) for _ in range(4)]
    return Border(left=sides[0], top=sides[1], right=sides[2], bottom=sides[3])


def _apply_cell_style(cell, style: CellStyle | None, default_size: int = 11):
    """对单元格应用样式"""
    if style is None:
        return
    if style.font_name or style.font_size or style.bold or style.italic or style.font_color:
        cell.font = _build_font(style, default_size)
    if style.bg_color:
        cell.fill = _build_fill(style)
    if style.alignment or style.vertical_alignment or style.wrap_text:
        cell.alignment = _build_alignment(style)
    border = _build_border(style)
    if border is not None:
        cell.border = border


def _get_column_style(sheet_def: SheetDef, col_idx: int) -> CellStyle:
    """获取某一列的列格式（如果有）"""
    from .models import CellStyle

    if sheet_def.column_formats and col_idx < len(sheet_def.column_formats):
        cf = sheet_def.column_formats[col_idx]
        cs = CellStyle(
            font_name=cf.font_name,
            font_size=cf.font_size,
            font_color=cf.font_color,
            bg_color=cf.bg_color,
            bold=cf.bold,
            italic=cf.italic,
            wrap_text=cf.wrap_text,
            number_format=cf.number_format,
            alignment=cf.alignment,
            vertical_alignment=cf.vertical_alignment,
        )
        return cs
    return CellStyle()


def _range_target(default_ws, workbook, range_text: str):
    """把 'A1:C10'/'Sheet1!A1:C10' 解析成 (工作表, 边界)。"""
    text = range_text.strip()
    if not text:
        raise ValueError("图表数据区域不能为空")
    if "!" in text:
        sheet_name, boundaries = range_to_tuple(text)
        try:
            target_ws = workbook[sheet_name]
        except KeyError as exc:
            raise ValueError(f"图表引用的工作表不存在: {sheet_name}") from exc
    else:
        target_ws = default_ws
        boundaries = range_boundaries(text)
    if any(value is None for value in boundaries):
        raise ValueError(
            f"范围必须包含具体行列（如 A1:C10），不支持整列/整行: {range_text!r}"
        )
    return target_ws, boundaries


def _validate_range_text(value: str, label: str) -> None:
    """拒绝整列/整行范围，避免 openpyxl 在 None 边界上抛裸 TypeError。"""
    try:
        boundaries = range_boundaries(value)
    except ValueError as exc:
        raise ValueError(f"无效的{label}范围: {value!r}") from exc
    if any(item is None for item in boundaries):
        raise ValueError(f"{label}范围必须包含具体行列（如 A1:C10），不支持整列/整行: {value!r}")


def _range_reference(default_ws, workbook, range_text: str) -> Reference:
    """把 'A1:C10'/'Sheet1!A1:C10' 解析成绑定到工作表的 Reference。"""
    target_ws, boundaries = _range_target(default_ws, workbook, range_text)
    min_col, min_row, max_col, max_row = boundaries
    return Reference(
        target_ws,
        min_col=min_col,
        min_row=min_row,
        max_col=max_col,
        max_row=max_row,
    )


def _add_scatter_series(chart, chart_def: ExcelChartDef, default_ws, workbook) -> list[str]:
    """散点图需要 X/Y 数值引用；categories_range 存在时 data_range 的每一列都是一个 Y 系列。"""
    from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
    from openpyxl.chart.series import Series, SeriesLabel

    data_ws, (dmin_col, dmin_row, dmax_col, dmax_row) = _range_target(
        default_ws, workbook, chart_def.data_range
    )
    if chart_def.categories_range:
        x_ws, (xmin_col, xmin_row, xmax_col, xmax_row) = _range_target(
            data_ws, workbook, chart_def.categories_range
        )
        y_cols = list(range(dmin_col, dmax_col + 1))
        y_min_row = dmin_row + 1 if chart_def.titles_from_data else dmin_row
    else:
        if dmax_col - dmin_col != 1:
            raise ValueError(
                "scatter 的 data_range 需要恰好两列（X,Y），或显式提供 categories_range"
            )
        x_ws = data_ws
        xmin_col = xmax_col = dmin_col
        xmin_row = dmin_row + 1 if chart_def.titles_from_data else dmin_row
        xmax_row = dmax_row
        y_cols = [dmax_col]
        y_min_row = xmin_row

    x_ref = Reference(
        x_ws, min_col=xmin_col, min_row=xmin_row, max_col=xmax_col, max_row=xmax_row
    )
    for index, y_col in enumerate(y_cols):
        y_ref = Reference(
            data_ws, min_col=y_col, min_row=y_min_row, max_col=y_col, max_row=dmax_row
        )
        title = None
        if chart_def.titles_from_data:
            title_ref = Reference(
                data_ws, min_col=y_col, min_row=dmin_row, max_col=y_col, max_row=dmin_row
            )
            title = SeriesLabel(strRef=StrRef(f=str(title_ref)))
        chart.series.append(
            Series(
                idx=index,
                order=index,
                xVal=AxDataSource(numRef=NumRef(f=str(x_ref))),
                yVal=NumDataSource(numRef=NumRef(f=str(y_ref))),
                tx=title,
            )
        )
    return []


def _new_chart(chart_type: str):
    if chart_type == "column":
        return BarChart(barDir="col")
    if chart_type == "bar":
        return BarChart(barDir="bar")
    chart_cls = _CHART_CLASSES.get(chart_type)
    if chart_cls is None:
        raise ValueError(f"不支持的图表类型: {chart_type}")
    return chart_cls()


def add_chart_to_sheet(ws, chart_def: ExcelChartDef, workbook) -> dict:
    """向工作表插入图表，返回实际写入的类型、锚点与被忽略的选项。"""
    chart = _new_chart(chart_def.chart_type)
    if chart_def.title:
        chart.title = chart_def.title
    warnings: list[str] = []
    if chart_def.chart_type == "scatter":
        warnings.extend(_add_scatter_series(chart, chart_def, ws, workbook))
    else:
        data_ref = _range_reference(ws, workbook, chart_def.data_range)
        chart.add_data(data_ref, titles_from_data=chart_def.titles_from_data)
        if chart_def.categories_range:
            categories_ref = _range_reference(
                data_ref.worksheet, workbook, chart_def.categories_range
            )
            chart.set_categories(categories_ref)

    if chart_def.legend == "none":
        chart.legend = None
    else:
        chart.legend.position = _LEGEND_POSITIONS[chart_def.legend]

    if chart_def.x_axis_title:
        if hasattr(chart, "x_axis"):
            chart.x_axis.title = chart_def.x_axis_title
        else:
            warnings.append(f"{chart_def.chart_type} 图表没有 x 轴，已忽略 x_axis_title")
    if chart_def.y_axis_title:
        if hasattr(chart, "y_axis"):
            chart.y_axis.title = chart_def.y_axis_title
        else:
            warnings.append(f"{chart_def.chart_type} 图表没有 y 轴，已忽略 y_axis_title")

    if chart_def.style is not None:
        chart.style = chart_def.style
    chart.width = chart_def.width
    chart.height = chart_def.height

    try:
        ws[chart_def.anchor]
    except (KeyError, TypeError, ValueError) as exc:
        raise ValueError(f"无效的图表锚点单元格: {chart_def.anchor}") from exc
    ws.add_chart(chart, chart_def.anchor)
    return {
        "chart_type": chart_def.chart_type,
        "anchor": chart_def.anchor,
        "warnings": warnings,
    }


def add_conditional_format_to_sheet(ws, cf_def: ConditionalFormatDef) -> None:
    """把单条条件格式规则追加到指定区域。"""
    _validate_range_text(cf_def.range, "条件格式")
    if cf_def.type in ("cell_is", "formula") and not cf_def.formula:
        raise ValueError(f"{cf_def.type} 条件格式必须提供 formula")
    fill = None
    font = None
    if cf_def.fill_color:
        fill = PatternFill(
            start_color=cf_def.fill_color,
            end_color=cf_def.fill_color,
            fill_type="solid",
        )
    if cf_def.font_color:
        font = Font(color=cf_def.font_color)

    if cf_def.type == "cell_is":
        if not cf_def.operator:
            raise ValueError("cell_is 条件格式必须提供 operator")
        rule = CellIsRule(
            operator=cf_def.operator,
            formula=list(cf_def.formula or []),
            fill=fill,
            font=font,
        )
    elif cf_def.type == "formula":
        rule = FormulaRule(formula=list(cf_def.formula or []), fill=fill, font=font)
    elif cf_def.type == "color_scale":
        colors = {
            "start_type": "min",
            "start_color": cf_def.min_color or "F8696B",
            "end_type": "max",
            "end_color": cf_def.max_color or "63BE7B",
        }
        if cf_def.mid_color:
            colors.update(
                mid_type="percentile",
                mid_value=50,
                mid_color=cf_def.mid_color,
            )
        rule = ColorScaleRule(**colors)
    elif cf_def.type == "data_bar":
        rule = DataBarRule(
            start_type="min",
            end_type="max",
            color=cf_def.bar_color or "638EC6",
            showValue=True,
        )
    elif cf_def.type == "icon_set":
        rule = IconSetRule(
            cf_def.icon_style or "3TrafficLights1",
            "percent",
            [0, 33, 67],
        )
    else:
        raise ValueError(f"不支持的条件格式类型: {cf_def.type}")
    ws.conditional_formatting.add(cf_def.range, rule)


def add_data_validation_to_sheet(ws, dv_def: DataValidationDef) -> None:
    """把单条数据验证规则追加到指定区域。"""
    from openpyxl.worksheet.datavalidation import DataValidation

    _validate_range_text(dv_def.range, "数据验证")

    validation = DataValidation(
        type=dv_def.type,
        formula1=dv_def.formula1,
        formula2=dv_def.formula2,
        allow_blank=dv_def.allow_blank,
        showDropDown=not dv_def.show_dropdown,
        showErrorMessage=bool(dv_def.error or dv_def.error_title),
        showInputMessage=bool(dv_def.prompt or dv_def.prompt_title),
        promptTitle=dv_def.prompt_title,
        prompt=dv_def.prompt,
        errorTitle=dv_def.error_title,
        error=dv_def.error,
    )
    ws.add_data_validation(validation)
    validation.add(dv_def.range)


def add_table_to_sheet(ws, table_def: ExcelTableDef) -> str:
    """添加超级表；show_totals_row 时自动把 ref 向下扩一行容纳总计行。"""
    _validate_range_text(table_def.range, "表格")
    min_col, min_row, max_col, max_row = range_boundaries(table_def.range)
    if max_row <= min_row:
        raise ValueError("表格区域必须包含表头行和至少一行数据")
    data_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    ref = data_ref
    table = Table(displayName=table_def.name, ref=ref)
    if table_def.show_totals_row:
        # OOXML 要求 totalsRowCount 与 ref 一致：总计行必须包含在表格范围内。
        ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row + 1}"
        table.ref = ref
        table.totalsRowCount = 1
        table.totalsRowShown = True
        # 自动筛选只覆盖表头与数据行，不把总计行计入。
        table.autoFilter = AutoFilter(ref=data_ref)
    table.tableStyleInfo = TableStyleInfo(
        name=table_def.style,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=table_def.show_row_stripes,
        showColumnStripes=False,
    )
    ws.add_table(table)
    return ref


def add_defined_name(workbook, defined_name: DefinedNameDef, ws=None) -> None:
    """添加工作簿级或工作表级定义名称；同名（忽略大小写）时明确报错，不静默覆盖。"""
    attr_text = defined_name.refers_to.strip()
    if attr_text.startswith("="):
        attr_text = attr_text[1:].strip()
    if not attr_text:
        raise ValueError(f"定义名称 {defined_name.name} 的 refers_to 不能为空")

    if defined_name.local:
        if ws is None:
            raise ValueError(f"工作表级定义名称 {defined_name.name} 必须指定所在工作表")
        target = ws.defined_names
        scope = f"工作表 {ws.title!r}"
    else:
        target = workbook.defined_names
        scope = "工作簿"
    existing = {str(name).casefold() for name in target}
    if defined_name.name.casefold() in existing:
        raise ValueError(f"{scope}中已存在定义名称: {defined_name.name!r}")
    target[defined_name.name] = DefinedName(name=defined_name.name, attr_text=attr_text)


def add_hyperlink_to_sheet(ws, link_def: ExcelHyperlinkDef) -> None:
    """给单元格写入外部超链接；只改字体颜色/下划线，保留已有粗体、填充与边框。"""
    from copy import copy

    cell = ws[link_def.cell]
    cell.hyperlink = link_def.url
    if link_def.tooltip:
        cell.hyperlink.tooltip = link_def.tooltip
    # cell.style = "Hyperlink" 会整体替换命名样式，丢掉表头粗体/底色/边框。
    font = copy(cell.font)
    font.color = "0563C1"
    font.underline = "single"
    cell.font = font


def add_comment_to_sheet(ws, comment_def: ExcelCommentDef) -> None:
    """给单元格写入批注。"""
    ws[comment_def.cell].comment = Comment(comment_def.text, comment_def.author)


def add_image_to_sheet(ws, image_def: ExcelImageDef) -> None:
    """按锚点插入图片；模型宽高为 cm，按 96 DPI 换算为像素，只给一维时保持比例。"""
    path = Path(image_def.path)
    if not path.is_file():
        raise ValueError(f"图片文件不存在: {image_def.path}")
    image = XLImage(str(path))
    original_width = float(image.width or 0)
    original_height = float(image.height or 0)
    if image_def.width is not None and image_def.height is not None:
        image.width = round(image_def.width / 2.54 * 96)
        image.height = round(image_def.height / 2.54 * 96)
    elif image_def.width is not None:
        image.width = round(image_def.width / 2.54 * 96)
        if original_width and original_height:
            image.height = round(image.width * original_height / original_width)
    elif image_def.height is not None:
        image.height = round(image_def.height / 2.54 * 96)
        if original_width and original_height:
            image.width = round(image.height * original_width / original_height)
    ws.add_image(image, image_def.anchor)


def apply_page_setup(ws, page_setup: ExcelPageSetupDef) -> None:
    """写入打印方向、纸张、缩放、打印区域/标题、页边距与居中选项。"""
    if page_setup.orientation:
        ws.page_setup.orientation = page_setup.orientation
    if page_setup.paper_size:
        ws.page_setup.paperSize = _paper_size_value(ws, page_setup.paper_size)
    if page_setup.fit_to_width is not None:
        ws.page_setup.fitToWidth = page_setup.fit_to_width
    if page_setup.fit_to_height is not None:
        ws.page_setup.fitToHeight = page_setup.fit_to_height
    if page_setup.fit_to_width is not None or page_setup.fit_to_height is not None:
        properties = ws.sheet_properties.pageSetUpPr or PageSetupProperties()
        properties.fitToPage = True
        ws.sheet_properties.pageSetUpPr = properties
    if page_setup.print_area:
        ws.print_area = page_setup.print_area
    if page_setup.print_title_rows:
        ws.print_title_rows = page_setup.print_title_rows
    if page_setup.print_title_cols:
        ws.print_title_cols = page_setup.print_title_cols
    for attr, value in (
        ("top", page_setup.margin_top),
        ("bottom", page_setup.margin_bottom),
        ("left", page_setup.margin_left),
        ("right", page_setup.margin_right),
    ):
        if value is not None:
            setattr(ws.page_margins, attr, value)
    if page_setup.center_horizontally is not None:
        ws.print_options.horizontalCentered = page_setup.center_horizontally
    if page_setup.center_vertically is not None:
        ws.print_options.verticalCentered = page_setup.center_vertically


def _paper_size_value(ws, paper_size: str) -> int:
    if paper_size in _PAPER_SIZE_FALLBACK:
        return _PAPER_SIZE_FALLBACK[paper_size]
    constant = f"PAPERSIZE_{paper_size.upper()}"
    return getattr(ws, constant)


# ── 工作表保护 ───────────────────────────────────────────────────

# openpyxl 的 SheetProtection 布尔属性 True 表示“禁止”，与模型的 allow_* 语义相反。
_PROTECTION_ALLOW_FLAGS = (
    ("formatCells", "allow_format_cells"),
    ("formatColumns", "allow_format_columns"),
    ("formatRows", "allow_format_rows"),
    ("insertRows", "allow_insert_rows"),
    ("insertColumns", "allow_insert_columns"),
    ("deleteRows", "allow_delete_rows"),
    ("deleteColumns", "allow_delete_columns"),
    ("sort", "allow_sort"),
    ("autoFilter", "allow_auto_filter"),
    ("selectLockedCells", "allow_select_locked_cells"),
    ("selectUnlockedCells", "allow_select_unlocked_cells"),
)


def apply_sheet_protection(ws, protection: SheetProtectionDef) -> None:
    """启用工作表保护：密码交给 openpyxl 转 legacy hash，allow_* 取反写入禁止位。"""
    guard = ws.protection
    if protection.password:
        # openpyxl 的 password setter 会顺带启用保护，并且存的是 hash 而不是明文。
        guard.password = protection.password
    guard.sheet = True
    for attribute, field in _PROTECTION_ALLOW_FLAGS:
        setattr(guard, attribute, not getattr(protection, field))


# ── 按行排序 ─────────────────────────────────────────────────────

_COLUMN_LETTER_RE = re.compile(r"^[A-Za-z]{1,3}$")
_COLUMN_NUMBER_RE = re.compile(r"^\d+$")
_MAX_COLUMNS = 16384
# 排序值类别：数字 < 字符串 < 空（None 永远排最后）
_RANK_NUMBER = 0
_RANK_STRING = 1
_RANK_EMPTY = 2


class _SheetSortPlan:
    """排序计划：重排后的数据行 + 列/方向 + sortState 区域。"""

    __slots__ = ("rows", "columns", "descending", "ref")

    def __init__(self, rows, columns, descending, ref):
        self.rows = rows
        self.columns = columns
        self.descending = descending
        self.ref = ref


def _sort_value(value):
    """把单元格值归一为 (类别, 比较键)；类别只用于区分数字/字符串/空。"""
    if isinstance(value, bool):
        # bool 是 int 子类：按 1/0 当数字处理，与 C# 后端保持一致。
        return (_RANK_NUMBER, int(value))
    if isinstance(value, int):
        return (_RANK_NUMBER, value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return (_RANK_EMPTY, None)
        return (_RANK_NUMBER, value)
    if isinstance(value, str):
        if not value:
            # Excel 把空字符串视同空白单元格。
            return (_RANK_EMPTY, None)
        # 先不区分大小写，再按原值兜底，保证同字母不同大小写有稳定次序。
        return (_RANK_STRING, (value.lower(), value))
    return (_RANK_EMPTY, None)


def _compare_sort_values(left, right) -> int:
    """比较同类别（都是数字或都是字符串）的两个归一值，返回 -1/0/1。"""
    left_rank, left_value = left
    right_rank, right_value = right
    if left_rank != right_rank:
        # 仅在升降序一致的类别比较里用到；空值由调用方单独处理。
        return -1 if left_rank < right_rank else 1
    if left_rank == _RANK_EMPTY:
        return 0
    if left_rank == _RANK_NUMBER:
        if left_value == right_value:
            return 0
        return -1 if left_value < right_value else 1
    if left_value[0] != right_value[0]:
        return -1 if left_value[0] < right_value[0] else 1
    if left_value[1] == right_value[1]:
        return 0
    return -1 if left_value[1] < right_value[1] else 1


def _sort_cell_value(row, column: int):
    """取数据行中第 column 列（1 基）的值；短行按空值处理。"""
    index = column - 1
    return row[index] if 0 <= index < len(row) else None


def _resolve_sort_column(
    sheet_def: SheetDef, key_column: str, columns_written: int, has_header: bool
) -> int:
    """把排序键解析成 1 基列号：有表头按表头名，无表头按列字母/列号。"""
    text = key_column.strip()
    if has_header:
        headers = sheet_def.headers or []
        if not headers:
            raise ValueError(
                "排序配置 has_header=true，但工作表未提供 headers，无法按表头名定位排序列"
            )
        if text in headers:
            return headers.index(text) + 1
        folded = text.casefold()
        for index, header in enumerate(headers, 1):
            if str(header).casefold() == folded:
                return index
        raise ValueError(
            f"排序键列 {key_column!r} 不在表头中（可用表头: "
            f"{', '.join(str(header) for header in headers)}）"
        )

    if _COLUMN_LETTER_RE.match(text):
        column = column_index_from_string(text.upper())
    elif _COLUMN_NUMBER_RE.match(text):
        column = int(text)
    else:
        raise ValueError(f"排序列必须是列字母（如 'B'）或 1 基数列号: {key_column!r}")
    if column < 1 or column > _MAX_COLUMNS:
        raise ValueError(f"排序列号超出工作表列范围 1-{_MAX_COLUMNS}: {key_column!r}")
    if column > columns_written:
        raise ValueError(
            f"排序列 {text!r} 超出数据范围（本次仅写入 {columns_written} 列）"
        )
    return column


def _written_columns(sheet_def: SheetDef) -> int:
    counts = [len(row) for row in sheet_def.data]
    if sheet_def.headers:
        counts.append(len(sheet_def.headers))
    return max(counts, default=0)


def _sort_range_bounds(sheet_def: SheetDef):
    """解析 sort.range（允许带 Sheet 前缀，仅取边界）；未指定时返回 None。"""
    if sheet_def.sort is None or not sheet_def.sort.range:
        return None
    text = sheet_def.sort.range.strip()
    try:
        boundaries = range_to_tuple(text)[1] if "!" in text else range_boundaries(text)
    except ValueError as exc:
        raise ValueError(f"无效的排序区域: {sheet_def.sort.range!r}") from exc
    if any(item is None for item in boundaries):
        raise ValueError(
            f"排序区域必须包含具体行列（如 A1:C10），不支持整列/整行: {sheet_def.sort.range!r}"
        )
    return boundaries


def prepare_sheet_sort(sheet_def: SheetDef) -> _SheetSortPlan | None:
    """按 SheetDef.sort 稳定多键重排本次写入的数据行；未配置排序时返回 None。"""
    if sheet_def.sort is None:
        return None
    if sheet_def.merged_cells:
        raise ValueError(
            "工作表同时配置了 merged_cells 与 sort：排序会移动合并区域并破坏合并，"
            "请去掉 merged_cells 或改为写入已排好序的数据"
        )

    sort = sheet_def.sort
    if sort.has_header and not sheet_def.headers:
        raise ValueError(
            "排序配置 has_header=true，但工作表未提供 headers，无法按表头名定位排序列"
        )
    columns_written = _written_columns(sheet_def)
    if columns_written == 0:
        raise ValueError("排序配置了 sort，但工作表没有任何表头或数据列")
    columns = [
        _resolve_sort_column(sheet_def, key.column, columns_written, sort.has_header)
        for key in sort.keys
    ]
    descending = [not key.ascending for key in sort.keys]

    bounds = _sort_range_bounds(sheet_def)
    if bounds is not None:
        min_col, _, max_col, _ = bounds
        for key, column in zip(sort.keys, columns):
            if not min_col <= column <= max_col:
                raise ValueError(
                    f"排序键 {key.column!r}（第 {column} 列）不在排序区域 "
                    f"{sort.range!r} 的列范围内"
                )

    rows = list(sheet_def.data)
    # 表头行永远不参与排序，避免表头被数据顶走。
    data_start_row = 2 if sheet_def.headers else 1
    if bounds is None:
        selected = list(range(len(rows)))
    else:
        min_row, max_row = bounds[1], bounds[3]
        selected = [
            index for index in range(len(rows))
            if min_row <= data_start_row + index <= max_row
        ]
    if not selected:
        if bounds is None:
            raise ValueError("排序配置了 sort，但没有可排序的数据行")
        raise ValueError(
            f"排序区域 {sort.range!r} 未覆盖任何数据行（数据从第 {data_start_row} 行开始）"
        )

    keyed = [
        ([_sort_value(_sort_cell_value(rows[index], column)) for column in columns], index)
        for index in selected
    ]

    def compare(left, right) -> int:
        for position, (left_value, right_value) in enumerate(zip(left[0], right[0])):
            left_blank = left_value[0] == _RANK_EMPTY
            right_blank = right_value[0] == _RANK_EMPTY
            # Excel 升序/降序都保持空白单元格在末尾，所以空白判定不受方向影响。
            if left_blank or right_blank:
                if left_blank and right_blank:
                    continue
                return 1 if left_blank else -1
            result = _compare_sort_values(left_value, right_value)
            if result:
                return -result if descending[position] else result
        return 0

    keyed.sort(key=cmp_to_key(compare))
    ordered = [rows[index] for _, index in keyed]
    for position, index in enumerate(selected):
        rows[index] = ordered[position]

    sort_ref = None
    if bounds is not None:
        min_col, min_row, max_col, max_row = bounds
        sort_ref = (
            f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        )
    return _SheetSortPlan(rows, columns, descending, sort_ref)


def _sort_state_ref(plan: _SheetSortPlan, sheet_def: SheetDef, filter_ref: str) -> str:
    """sortState 的区域：优先 sort.range，否则与筛选区域一致（含表头行）。"""
    if plan.ref:
        return plan.ref
    return filter_ref


def attach_sort_state(ws, sheet_def: SheetDef, plan: _SheetSortPlan) -> None:
    """在自动筛选上追加 sortState/sortCondition；没有 autoFilter 时不额外造元素。"""
    if not ws.auto_filter.ref:
        return
    ref = _sort_state_ref(plan, sheet_def, ws.auto_filter.ref)
    _, min_row, _, max_row = range_boundaries(ref)
    # 与 Excel/openpyxl 一致：sortState 区域含表头，sortCondition 只覆盖数据行。
    data_start_row = 2 if sheet_def.headers else 1
    if min_row < data_start_row <= max_row:
        min_row = data_start_row
    conditions = []
    for column, is_descending in zip(plan.columns, plan.descending):
        letter = get_column_letter(column)
        conditions.append(
            SortCondition(
                ref=f"{letter}{min_row}:{letter}{max_row}",
                descending=True if is_descending else None,
            )
        )
    ws.auto_filter.sortState = SortState(ref=ref, sortCondition=conditions)


def _write_sheet(ws, sheet_def: SheetDef, workbook=None):
    """写入一个工作表"""

    # 排序在写行之前完成：计划失败时整张表都不写入，避免留下半成品。
    sort_plan = prepare_sheet_sort(sheet_def)
    data_rows = sort_plan.rows if sort_plan is not None else sheet_def.data

    # 写入表头
    if sheet_def.headers:
        for ci, h in enumerate(sheet_def.headers, 1):
            cell = ws.cell(row=1, column=ci, value=h)
            _apply_cell_style(cell, sheet_def.header_style)
            # 如果没有特别指定表头样式，默认加粗
            if sheet_def.header_style is None:
                cell.font = Font(bold=True)

    start_row = 2 if sheet_def.headers else 1

    # 写入数据
    for ri, row_data in enumerate(data_rows, start_row):
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci)
            # 公式识别：字符串以 = 开头
            if sheet_def.allow_formulas and isinstance(val, str) and val.startswith("="):
                cell.value = val  # openpyxl 自动识别为公式
            elif isinstance(val, str) and val.startswith("="):
                cell.value = "'" + val
            else:
                cell.value = val

            # 应用列格式
            col_style = _get_column_style(sheet_def, ci - 1)
            if col_style:
                _apply_cell_style(cell, col_style)
                if col_style.number_format:
                    cell.number_format = col_style.number_format

    # 列宽（ColumnFormat.width 覆盖 column_widths）
    if sheet_def.column_widths:
        for ci, w in enumerate(sheet_def.column_widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
    if sheet_def.column_formats:
        for ci, cf in enumerate(sheet_def.column_formats, 1):
            if cf.width is not None:
                ws.column_dimensions[get_column_letter(ci)].width = cf.width

    # 行高
    if sheet_def.row_heights:
        for row, height in sheet_def.row_heights.items():
            ws.row_dimensions[int(row)].height = float(height)

    # 合并单元格
    if sheet_def.merged_cells:
        for mc in sheet_def.merged_cells:
            ws.merge_cells(mc.range)

    # 冻结窗格
    if sheet_def.freeze:
        ws.freeze_panes = sheet_def.freeze

    # 自动筛选（排序只写进 autoFilter 的 sortState，不额外造 autoFilter）
    if sheet_def.auto_filter and sheet_def.headers and data_rows:
        last_col = len(sheet_def.headers)
        last_row = len(data_rows) + (1 if sheet_def.headers else 0)
        ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{last_row}"
        if sort_plan is not None:
            attach_sort_state(ws, sheet_def, sort_plan)

    # Sheet 标签颜色
    if sheet_def.tab_color:
        ws.sheet_properties.tabColor = sheet_def.tab_color

    # 工作表可见性
    ws.sheet_state = sheet_def.sheet_state

    # 单元格超链接与批注
    if sheet_def.hyperlinks:
        for link in sheet_def.hyperlinks:
            add_hyperlink_to_sheet(ws, link)
    if sheet_def.comments:
        for comment in sheet_def.comments:
            add_comment_to_sheet(ws, comment)

    # 图片
    if sheet_def.images:
        for image in sheet_def.images:
            add_image_to_sheet(ws, image)

    # 条件格式
    if sheet_def.conditional_formats:
        for cf in sheet_def.conditional_formats:
            add_conditional_format_to_sheet(ws, cf)

    # 数据验证
    if sheet_def.data_validations:
        for dv in sheet_def.data_validations:
            add_data_validation_to_sheet(ws, dv)

    # 超级表
    if sheet_def.tables:
        for table in sheet_def.tables:
            add_table_to_sheet(ws, table)

    # 页面设置
    if sheet_def.page_setup:
        apply_page_setup(ws, sheet_def.page_setup)

    # 定义名称
    owner = workbook if workbook is not None else ws.parent
    if sheet_def.defined_names:
        for defined_name in sheet_def.defined_names:
            add_defined_name(owner, defined_name, ws)

    # 图表
    if sheet_def.charts:
        for chart in sheet_def.charts:
            result = add_chart_to_sheet(ws, chart, owner)
            for warning in result["warnings"]:
                logger.warning("%s: %s", sheet_def.name, warning)

    # 工作表保护：放在数据/合并/筛选之后，避免保护位影响本次写入。
    if sheet_def.protection:
        apply_sheet_protection(ws, sheet_def.protection)


def create_excel_workbook(
    output_path: str,
    sheets: list[SheetDef],
    template: str | None = None,
    if_exists: str = "overwrite",
) -> CreateResult:
    """创建 Excel 工作簿"""
    from .models import CreateResult
    from .utils import atomic_save, ensure_output_dir, get_file_size, resolve_if_exists, validate_input_file, validate_output_path

    wb = None
    try:
        validate_output_path(output_path, (".xlsx", ".xlsm"))
        ensure_output_dir(output_path)
        resolved_path = resolve_if_exists(output_path, if_exists)
    except (OSError, ValueError) as e:
        return CreateResult(success=False, file_path=output_path, message=str(e))

    try:
        output_is_macro = resolved_path.lower().endswith(".xlsm")
        template_is_macro = bool(template and template.lower().endswith((".xlsm", ".xltm")))
        if output_is_macro and not template_is_macro:
            raise ValueError("生成 .xlsm 必须提供 .xlsm 或 .xltm 宏模板，以避免扩展名与内容不匹配")
        if template_is_macro and not output_is_macro:
            raise ValueError("使用宏模板时输出文件必须为 .xlsm，以避免丢失 VBA 内容")

        if template:
            validate_input_file(template, (".xlsx", ".xlsm", ".xltx", ".xltm"))
            wb = load_workbook(
                template,
                keep_vba=template.lower().endswith((".xlsm", ".xltm")),
                rich_text=True,
            )
            # 删除默认 sheet
            for name in list(wb.sheetnames):
                del wb[name]
        else:
            wb = Workbook()
            # 删除默认 sheet
            wb.remove(wb.active)

        # 先建全部工作表，图表才能引用尚未写入的其他新表
        worksheets = [
            (wb.create_sheet(title=sheet_def.name[:31]), sheet_def)  # Excel 限制 sheet 名 31 字符
            for sheet_def in sheets
        ]
        for ws, sheet_def in worksheets:
            _write_sheet(ws, sheet_def, wb)

        backup_path = atomic_save(wb.save, resolved_path)
        size = get_file_size(resolved_path)

        return CreateResult(
            success=True,
            file_path=resolved_path,
            file_size=size,
            sheets_count=len(sheets),
            message=f"Excel 工作簿创建成功: {resolved_path}",
            backup_path=backup_path,
        )
    except Exception as e:
        return CreateResult(
            success=False,
            file_path=resolved_path,
            message=f"Excel 工作簿创建失败: {str(e)}",
        )
    finally:
        if wb is not None:
            _close_workbook(wb)


def edit_excel_workbook(
    file_path: str,
    append_sheets: list[SheetDef],
) -> EditResult:
    """向已有 Excel 追加工作表"""
    from .models import EditResult
    from .utils import atomic_save, validate_input_file

    # An open Excel instance can initialise an Explorer-created placeholder
    # before it becomes a valid OOXML archive on disk.
    from .office_com import try_edit_open_excel_workbook
    com_result = try_edit_open_excel_workbook(file_path, append_sheets)
    if com_result is not None:
        return com_result

    wb = None
    try:
        validate_input_file(file_path, (".xlsx", ".xlsm"))
    except ValueError as e:
        return EditResult(
            success=False,
            file_path=file_path,
            message=str(e),
        )

    try:
        # `.xlsm` 保存时必须保留 VBA archive，否则追加工作表会把宏静默剥离。
        wb = load_workbook(
            file_path,
            keep_vba=file_path.lower().endswith(".xlsm"),
            rich_text=True,
        )
        # 先建全部工作表，图表才能引用本次追加的其他表
        worksheets = [
            (wb.create_sheet(title=sheet_def.name[:31]), sheet_def)
            for sheet_def in append_sheets
        ]
        for ws, sheet_def in worksheets:
            _write_sheet(ws, sheet_def, wb)
        count = len(worksheets)
        backup_path = atomic_save(wb.save, file_path)
        return EditResult(
            success=True,
            file_path=file_path,
            operation_count=count,
            message=f"已追加 {count} 个工作表到 Excel 文件",
            backup_path=backup_path,
        )
    except Exception as e:
        return EditResult(
            success=False,
            file_path=file_path,
            message=f"编辑 Excel 工作簿失败: {str(e)}",
        )
    finally:
        if wb is not None:
            _close_workbook(wb)
