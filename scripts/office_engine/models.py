# Vendored from mcp_office_writer/models.py (mcp-office-writer 0.6.0, 2026-11).
# Upstream path: mcp_office_writer/models.py (repository-relative).
# Modifications: (2026-11, local) extended Word/Excel/PPT models — hyperlink/field/properties,
# Excel charts/conditional formats/data validations/tables/names/comments/images/page setup,
# PPT table/chart geometry, image crop, hyperlinks, and new entry-point models.

"""
mcp-office-writer 数据模型定义
所有 Pydantic 模型集中在此，供各 writer 模块共享
"""

from __future__ import annotations

import re
from datetime import datetime, timezone
from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator, model_validator

# A1 范围：可选 Sheet 前缀，必须含具体行号（拒绝 A:C、1:10 这类整列/整行）
_A1_RANGE = re.compile(
    r"^(?:(?:'[^']+'|[^!]+)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$"
)


def _check_a1_range(value: str | None, label: str) -> str | None:
    if value is None:
        return None
    text = value.strip()
    if not text or not _A1_RANGE.match(text):
        raise ValueError(f"{label} 必须是含具体行列的 A1 范围（如 A1:C10），不支持整列/整行或空值: {value!r}")
    return text


# ── 通用类型 ──────────────────────────────────────────────────────

class FontConfig(BaseModel):
    """字体配置"""
    name: str | None = Field(None, description="字体名称，如 '宋体'、'微软雅黑'、'Arial'")
    size: int | str | None = Field(None, description="字号（国内标准如'小四''五号'，或直接 pt 数值）")
    bold: bool | None = Field(None, description="是否加粗")
    italic: bool | None = Field(None, description="是否斜体")
    color: str | None = Field(None, description="字体颜色 (HEX, 如 'FF0000')")


class BorderSpec(BaseModel):
    """单元格四边框样式（四边同款）"""
    style: Literal["thin", "medium", "thick", "dashed", "dotted", "double", "hair"] = "thin"
    color: str | None = Field(None, description="边框颜色 HEX，默认黑色")


class CellStyle(BaseModel):
    """单元格样式"""
    bold: bool | None = None
    italic: bool | None = None
    font_name: str | None = None
    font_size: int | str | None = None
    font_color: str | None = Field(None, description="字体颜色 HEX")
    bg_color: str | None = Field(None, description="背景颜色 HEX")
    number_format: str | None = Field(None, description="数字格式，如 '#,##0.00'、'0%'、'YYYY-MM-DD'")
    alignment: Literal["left", "center", "right"] | None = None
    vertical_alignment: Literal["top", "center", "bottom"] | None = None
    wrap_text: bool | None = None
    border: BorderSpec | None = Field(None, description="四周边框；None 不设置")


class PageSetup(BaseModel):
    """页面设置"""
    orientation: Literal["portrait", "landscape"] | None = None
    paper_size: Literal["A4", "Letter", "A3", "B5"] | None = "A4"
    margin_top: float | None = Field(None, description="上边距 (英寸)")
    margin_bottom: float | None = Field(None, description="下边距 (英寸)")
    margin_left: float | None = Field(None, description="左边距 (英寸)")
    margin_right: float | None = Field(None, description="右边距 (英寸)")
    columns: int | None = Field(None, ge=1, le=4, description="分栏数（1-4）")
    column_spacing: float | None = Field(None, description="分栏间距 (英寸)")
    page_border: Literal["single", "double", "dashed", "dotted"] | None = Field(
        None, description="页面边框线型")
    page_border_color: str | None = Field(None, description="页面边框颜色 HEX")
    watermark_text: str | None = Field(None, description="文字水印内容")
    watermark_color: str | None = Field(None, description="水印颜色 HEX，默认浅灰")


class HeaderFooter(BaseModel):
    """页眉页脚设置"""
    header_text: str | None = Field(None, description="页眉文字")
    footer_text: str | None = Field(None, description="页脚文字，用 {page} 代表页码")


PageNumberStyle = Literal["none", "roman", "arabic"]


class SceneProfile(BaseModel):
    """场景排版配置文件"""
    name: str = Field(..., description="场景名称标识")
    display_name: str = Field("", description="场景显示名称")
    description: str | None = Field(None, description="场景描述")

    # 字体
    font_name: str = "宋体"
    font_name_western: str = Field("Times New Roman", description="西文字体")
    heading_font: str | None = None  # None = 使用 font_name
    body_size: str = "小四"           # 正文字号（国内标准）
    heading_sizes: dict[int, str] = Field(default_factory=lambda: {1: "三号", 2: "小三", 3: "四号", 4: "小四"})

    # 段落
    line_spacing: float = 1.3
    first_line_indent: int | None = 28  # pt, ≈ 2 字符
    alignment: str = "justify"

    # 页面
    paper_size: str = "A4"
    margin_top: float = 1.0
    margin_bottom: float = 1.0
    margin_left: float = 1.18   # 3cm
    margin_right: float = 1.18

    # 页码格式
    page_number_style: PageNumberStyle = "arabic"
    show_cover: bool = False


class CoverConfig(BaseModel):
    """封面配置"""
    title: str | None = Field(None, description="封面主标题；留空时用第一个 heading 或默认标题")
    subtitle: str | None = Field(None, description="封面副标题")
    author: str | None = Field(None, description="作者/创建者")
    date: str | None = Field(None, description="日期")
    organization: str | None = Field(None, description="机构/单位名称")
    extra_lines: list[str] | None = Field(None, description="附加信息行")
    recipe_id: str | None = Field(None, description="封面配方 ID (R1-R7)，自动选择时可留空")
    doc_type: str | None = Field(None, description="文档类型（general/report/thesis 等），用于自动选配方")


class MergeRange(BaseModel):
    """表格合并单元格范围 (Excel A1 风格)"""
    range: str = Field(..., description="合并范围，如 'A1:B2'、'C1:C3'")


# ── Word 相关模型 ─────────────────────────────────────────────────

class ParagraphStyle(BaseModel):
    """段落样式"""
    font_name: str | None = None
    font_size: int | str | None = Field(None, description="字号（国内标准如'小四''五号'，或直接 pt 数值）")
    bold: bool | None = None
    italic: bool | None = None
    color: str | None = None
    alignment: Literal["left", "center", "right", "justify"] | None = None
    line_spacing: float | None = Field(None, description="行距倍数，如 1.0、1.5、2.0")
    first_line_indent: int | None = Field(None, description="首行缩进 (pt)，中文段落常用 28pt≈2字符")
    space_before: int | None = Field(None, description="段前距 (pt)")
    space_after: int | None = Field(None, description="段后距 (pt)")


class TableDef(BaseModel):
    """Word 表格定义"""
    headers: list[str] = Field(default_factory=list, description="表头行")
    rows: list[list[str]] = Field(default_factory=list, description="数据行")
    column_widths: list[float] | None = Field(None, description="列宽 (英寸)")
    header_style: CellStyle | None = Field(None, description="表头样式")
    merged_cells: list[MergeRange] | None = Field(None, description="合并单元格")


class ImageDef(BaseModel):
    """图片定义"""
    path: str = Field(..., description="图片文件路径")
    width: float | None = Field(None, description="宽度 (英寸)")
    height: float | None = Field(None, description="高度 (英寸)")


class HyperlinkDef(BaseModel):
    """Word 超链接定义"""
    text: str = Field(..., description="显示文字")
    url: str = Field(..., description="链接地址（http/https/mailto/file）")
    tooltip: str | None = Field(None, description="悬停提示")


class FieldDef(BaseModel):
    """Word 域定义（PAGE/NUMPAGES/DATE 等）"""
    field_type: Literal[
        "PAGE", "NUMPAGES", "SECTIONPAGES", "DATE", "TIME",
        "FILENAME", "AUTHOR", "TITLE", "CUSTOM",
    ] = Field("PAGE", description="域类型；CUSTOM 时用 code 提供完整域代码")
    code: str | None = Field(None, description="自定义域代码，仅 CUSTOM 类型使用，如 ' MERGEFIELD 姓名 '")
    fallback_text: str | None = Field(None, description="域结果占位文字，Word 更新域前显示")
    number_format: str | None = Field(None, description="数字格式开关，如 'roman'、'DBNUM1'；用于 PAGE/NUMPAGES")

    @model_validator(mode="after")
    def _check_custom_code(self):
        if self.field_type == "CUSTOM" and not (self.code or "").strip():
            raise ValueError("field_type=CUSTOM 时必须提供非空 code")
        return self


class BookmarkDef(BaseModel):
    """Word 书签（包裹 text 内容）"""
    name: str = Field(..., description="书签名，字母/数字/下划线，不超过 40 字符")


class SectionItem(BaseModel):
    """Word 文档的一个内容块"""
    type: Literal["heading", "paragraph", "table", "image", "page_break", "bullet_list", "numbered_list", "code", "toc", "section_break", "hyperlink", "field", "bookmark"]
    text: str | None = Field(None, description="文字内容")
    level: int | None = Field(None, description="标题层级 (1-9)，仅 heading 类型使用")
    language: str | None = Field(None, description="代码语言标签，仅 code 类型使用")
    style: ParagraphStyle | None = Field(None, description="段落样式")
    items: list[str] | None = Field(None, description="列表项，仅 bullet_list/numbered_list 使用")
    table: TableDef | None = Field(None, description="表格定义，仅 table 类型使用")
    image: ImageDef | None = Field(None, description="图片定义，仅 image 类型使用")
    toc_levels: str | None = Field(None, description="TOC 包含的标题层级，如 '1-3'，仅 toc 类型使用")
    hyperlink: HyperlinkDef | None = Field(None, description="超链接定义，仅 hyperlink 类型使用")
    field: FieldDef | None = Field(None, description="域定义，仅 field 类型使用")
    bookmark: BookmarkDef | None = Field(None, description="书签定义，仅 bookmark 类型使用")

    @model_validator(mode="after")
    def _check_payload(self):
        if self.type == "hyperlink" and self.hyperlink is None:
            raise ValueError("type=hyperlink 时必须提供 hyperlink 定义")
        if self.type == "field" and self.field is None:
            raise ValueError("type=field 时必须提供 field 定义")
        if self.type == "bookmark" and self.bookmark is None:
            raise ValueError("type=bookmark 时必须提供 bookmark 定义")
        return self


class WordProperties(BaseModel):
    """Word 文档核心属性"""
    title: str | None = None
    subject: str | None = None
    author: str | None = None
    keywords: str | None = None
    comments: str | None = None
    category: str | None = None
    last_modified_by: str | None = None
    created: str | None = Field(None, description="创建时间，ISO 8601，如 '2026-11-19T10:00:00'")
    modified: str | None = Field(None, description="修改时间，ISO 8601")

    @field_validator("created", "modified")
    @classmethod
    def _normalise_datetime(cls, value: str | None) -> str | None:
        if value is None:
            return None
        text = value.strip()
        if text.endswith("Z"):
            text = text[:-1] + "+00:00"
        try:
            parsed = datetime.fromisoformat(text)
        except ValueError as exc:
            raise ValueError(f"不是合法的 ISO 8601 时间: {value!r}") from exc
        if parsed.tzinfo is not None:
            parsed = parsed.astimezone(timezone.utc).replace(tzinfo=None)
        return parsed.isoformat()


class WordContent(BaseModel):
    """Word 文档内容"""
    sections: list[SectionItem] = Field(..., description="文档内容块列表")
    properties: WordProperties | None = Field(None, description="写入文档核心属性（标题/作者/主题等）")


# ── Excel 相关模型 ───────────────────────────────────────────────

class ColumnFormat(BaseModel):
    """Excel 列格式"""
    number_format: str | None = None
    font_name: str | None = None
    font_size: int | None = None
    font_color: str | None = None
    bg_color: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    wrap_text: bool | None = None
    width: float | None = Field(None, description="列宽 (字符数)")
    alignment: Literal["left", "center", "right"] | None = None
    vertical_alignment: Literal["top", "center", "bottom"] | None = None


class ExcelChartDef(BaseModel):
    """Excel 图表定义（插入到指定工作表）"""
    chart_type: Literal["column", "bar", "line", "pie", "doughnut", "area", "scatter", "radar"] = Field(
        "column", description="图表类型"
    )
    title: str | None = None
    data_range: str = Field(..., description="数据区域，如 'A1:C10' 或 'Sheet1!A1:C10'（首行/首列可作系列名与分类名）")
    categories_range: str | None = Field(None, description="分类轴区域，如 'A2:A10'；缺省时由 data_range 推断")
    titles_from_data: bool = Field(True, description="data_range 首行是否作为系列名")
    anchor: str = Field("E2", description="图表左上角锚点单元格，如 'E2'")
    width: float = Field(15, description="图表宽度 (cm)")
    height: float = Field(8, description="图表高度 (cm)")
    legend: Literal["right", "bottom", "top", "left", "none"] = Field("right", description="图例位置")
    x_axis_title: str | None = None
    y_axis_title: str | None = None
    style: int | None = Field(None, ge=1, le=48, description="Excel 内置图表样式编号 1-48")

    @field_validator("data_range", "categories_range")
    @classmethod
    def _valid_chart_range(cls, value: str | None) -> str | None:
        return _check_a1_range(value, "图表数据范围")


class ConditionalFormatDef(BaseModel):
    """Excel 条件格式规则"""
    range: str = Field(..., description="作用区域，如 'A2:A100'")
    type: Literal["cell_is", "formula", "color_scale", "data_bar", "icon_set"] = Field(
        "cell_is", description="规则类型"
    )
    operator: Literal[
        "greaterThan", "lessThan", "between", "equal", "notEqual",
        "greaterThanOrEqual", "lessThanOrEqual",
    ] | None = Field(None, description="cell_is 比较运算符")
    formula: list[str] | None = Field(None, description="比较值或公式列表，如 ['100'] 或 ['$A1>5']")
    fill_color: str | None = Field(None, description="满足条件时的背景色 HEX（cell_is/formula）")
    font_color: str | None = Field(None, description="满足条件时的字体色 HEX（cell_is/formula）")
    min_color: str | None = Field(None, description="色阶最小值颜色（color_scale）")
    mid_color: str | None = Field(None, description="色阶中间值颜色（color_scale）")
    max_color: str | None = Field(None, description="色阶最大值颜色（color_scale）")
    bar_color: str | None = Field(None, description="数据条颜色（data_bar）")
    icon_style: str | None = Field(None, description="图标集名称（icon_set），如 '3TrafficLights1'，默认同左")

    @field_validator("range")
    @classmethod
    def _valid_range(cls, value: str) -> str:
        return _check_a1_range(value, "条件格式范围") or value


class DataValidationDef(BaseModel):
    """Excel 数据验证规则"""
    range: str = Field(..., description="作用区域，如 'B2:B100'")
    type: Literal["list", "whole", "decimal", "date", "time", "textLength", "custom"] = Field(
        "list", description="验证类型"
    )
    formula1: str = Field(..., description="条件 1；list 类型为逗号分隔值或区域引用，如 '\"是,否\"' 或 '=$D$1:$D$5'")
    formula2: str | None = Field(None, description="条件 2（between 类验证使用）")
    allow_blank: bool = Field(True, description="是否允许空值")
    show_dropdown: bool = Field(True, description="list 类型是否显示下拉箭头")
    prompt_title: str | None = None
    prompt: str | None = None
    error_title: str | None = None
    error: str | None = None

    @field_validator("range")
    @classmethod
    def _valid_range(cls, value: str) -> str:
        return _check_a1_range(value, "数据验证范围") or value


class ExcelTableDef(BaseModel):
    """Excel 超级表（ListObject）定义"""
    range: str = Field(..., description="表格区域，如 'A1:D10'（必须包含表头行）")
    name: str = Field(..., description="表格名称（不含空格）")
    style: str = Field("TableStyleMedium2", description="表格样式名")
    show_row_stripes: bool = True
    show_totals_row: bool = False

    @field_validator("range")
    @classmethod
    def _valid_range(cls, value: str) -> str:
        return _check_a1_range(value, "表格范围") or value


class DefinedNameDef(BaseModel):
    """Excel 定义名称"""
    name: str = Field(..., description="名称，如 'SalesTotal'")
    refers_to: str = Field(..., description="引用内容，如 'Sheet1!$B$2:$B$10' 或 '=SUM(Sheet1!$B$2:$B$10)'")
    local: bool = Field(False, description="True 表示只作用于当前工作表")


class ExcelHyperlinkDef(BaseModel):
    """Excel 单元格超链接"""
    cell: str = Field(..., description="单元格，如 'A2'")
    url: str = Field(..., description="链接地址")
    tooltip: str | None = None


class ExcelCommentDef(BaseModel):
    """Excel 单元格批注"""
    cell: str = Field(..., description="单元格，如 'C3'")
    text: str = Field(..., description="批注正文")
    author: str = Field("office-writer", description="批注作者")


class ExcelImageDef(BaseModel):
    """Excel 工作表图片"""
    path: str = Field(..., description="本地图片路径")
    anchor: str = Field("A1", description="左上角锚点单元格")
    width: float | None = Field(None, description="宽度 (cm)")
    height: float | None = Field(None, description="高度 (cm)")


class ExcelPageSetupDef(BaseModel):
    """Excel 打印与页面设置"""
    orientation: Literal["portrait", "landscape"] | None = None
    paper_size: Literal["A4", "Letter", "A3", "B5"] | None = None
    fit_to_width: int | None = Field(None, description="缩放到页宽")
    fit_to_height: int | None = Field(None, description="缩放到页高")
    print_area: str | None = Field(None, description="打印区域，如 'A1:F30'")
    print_title_rows: str | None = Field(None, description="重复标题行，如 '1:1'")
    print_title_cols: str | None = Field(None, description="重复标题列，如 'A:A'")
    margin_top: float | None = Field(None, description="上边距 (英寸)")
    margin_bottom: float | None = Field(None, description="下边距 (英寸)")
    margin_left: float | None = Field(None, description="左边距 (英寸)")
    margin_right: float | None = Field(None, description="右边距 (英寸)")
    center_horizontally: bool | None = None
    center_vertically: bool | None = None


class SheetProtectionDef(BaseModel):
    """工作表保护（allow_* 为 True 表示允许该操作）"""
    password: str | None = Field(None, description="保护密码（明文，写入 OOXML hash）")
    allow_select_locked_cells: bool = Field(True, description="允许选中锁定单元格")
    allow_select_unlocked_cells: bool = Field(True, description="允许选中未锁定单元格")
    allow_format_cells: bool = False
    allow_format_columns: bool = False
    allow_format_rows: bool = False
    allow_insert_rows: bool = False
    allow_insert_columns: bool = False
    allow_delete_rows: bool = False
    allow_delete_columns: bool = False
    allow_sort: bool = False
    allow_auto_filter: bool = False


class SortKeyDef(BaseModel):
    """排序键：column 为表头名；has_header=false 时可用列字母，如 'B'"""
    column: str = Field(..., description="表头名或列字母")
    ascending: bool = Field(True, description="是否升序")


class ExcelSortDef(BaseModel):
    """按行排序（写数据后重排行，并写 sortState）"""
    range: str | None = Field(None, description="排序区域，如 'A1:D20'；默认整个已用区域")
    keys: list[SortKeyDef] = Field(..., min_length=1, description="排序键，按顺序比较")
    has_header: bool = Field(True, description="首行是否为表头")


class SheetDef(BaseModel):
    """Excel 工作表定义"""
    name: str = Field(..., description="工作表名称")
    headers: list[str] | None = Field(None, description="表头行")
    data: list[list] = Field(default_factory=list, description="数据行（支持 str/int/float/None）")
    allow_formulas: bool = Field(False, description="是否将以 = 开头的字符串写为 Excel 公式；默认按纯文本写入")
    column_widths: list[float] | None = Field(None, description="列宽 (字符数)")
    row_heights: dict[int, float] | None = Field(None, description="行高（行号→磅值），如 {\"1\": 24}")
    header_style: CellStyle | None = Field(None, description="表头样式")
    column_formats: list[ColumnFormat] | None = Field(None, description="按列格式")
    merged_cells: list[MergeRange] | None = Field(None, description="合并单元格")
    freeze: str | None = Field(None, description="冻结窗格，如 'A2' 冻结首行")
    auto_filter: bool = Field(False, description="是否启用自动筛选")
    tab_color: str | None = Field(None, description="工作表标签颜色 HEX")
    sheet_state: Literal["visible", "hidden", "veryHidden"] = Field("visible", description="工作表可见性")
    charts: list[ExcelChartDef] | None = Field(None, description="插入的图表")
    conditional_formats: list[ConditionalFormatDef] | None = Field(None, description="条件格式规则")
    data_validations: list[DataValidationDef] | None = Field(None, description="数据验证规则")
    tables: list[ExcelTableDef] | None = Field(None, description="超级表")
    defined_names: list[DefinedNameDef] | None = Field(None, description="定义名称（工作簿级或本表级）")
    hyperlinks: list[ExcelHyperlinkDef] | None = Field(None, description="单元格超链接")
    comments: list[ExcelCommentDef] | None = Field(None, description="单元格批注")
    images: list[ExcelImageDef] | None = Field(None, description="插入图片")
    page_setup: ExcelPageSetupDef | None = Field(None, description="打印与页面设置")
    protection: SheetProtectionDef | None = Field(None, description="工作表保护")
    sort: ExcelSortDef | None = Field(None, description="写数据后按行排序")


class ExcelContent(BaseModel):
    """Excel 工作簿内容"""
    sheets: list[SheetDef] = Field(..., description="工作表列表（至少 1 个）")


# ── PPT 相关模型 ──────────────────────────────────────────────────

LayoutType = Literal[
    "title", "section", "content", "two_content", "comparison",
    "title_only", "blank", "content_caption", "picture_caption"
]


class TextStyle(BaseModel):
    """PPT 文字样式"""
    font_name: str | None = None
    font_size: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    color: str | None = None
    alignment: Literal["left", "center", "right"] | None = None


class SlideTableDef(BaseModel):
    """PPT 幻灯片表格定义"""
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    x: float | None = Field(None, description="左侧位置 (英寸)，默认 0.5")
    y: float | None = Field(None, description="顶部位置 (英寸)，默认 2.0")
    width: float | None = Field(None, description="表格宽度 (英寸)，默认 7.0")
    height: float | None = Field(None, description="表格高度 (英寸)，默认 2.5")
    column_widths: list[float] | None = Field(None, description="各列宽度 (英寸)，总宽超出表格宽度时按比例缩放")
    header_bold: bool = Field(True, description="表头是否加粗")


class SlideImageDef(BaseModel):
    """PPT 幻灯片图片定义"""
    path: str | None = Field(None, description="图片文件路径（与 url 二选一）")
    url: str | None = Field(None, description="图片 URL（与 path 二选一），自动下载后插入")
    x: float | None = Field(None, description="左侧位置 (英寸)")
    y: float | None = Field(None, description="顶部位置 (英寸)")
    width: float | None = Field(None, description="宽度 (英寸)")
    height: float | None = Field(None, description="高度 (英寸)")
    crop_left: float | None = Field(None, ge=0, lt=1, description="左侧裁剪比例 (0~1)")
    crop_right: float | None = Field(None, ge=0, lt=1, description="右侧裁剪比例 (0~1)")
    crop_top: float | None = Field(None, ge=0, lt=1, description="顶部裁剪比例 (0~1)")
    crop_bottom: float | None = Field(None, ge=0, lt=1, description="底部裁剪比例 (0~1)")
    hyperlink_url: str | None = Field(None, description="点击图片打开的链接")

    @model_validator(mode="after")
    def _check_path_or_url(self):
        if not self.path and not self.url:
            raise ValueError("SlideImageDef 必须提供 path（本地文件）或 url（在线图片）之一")
        if self.path and self.url:
            raise ValueError("SlideImageDef 不能同时提供 path 和 url，请二选一")
        return self


class SlideShapeDef(BaseModel):
    """PPT 形状定义"""
    shape_type: Literal["rect", "rounded_rect", "oval", "arrow", "cylinder", "diamond"] = "rect"
    text: str | None = Field(None, description="形状内文字")
    hyperlink_url: str | None = Field(None, description="点击形状打开的链接")
    x: float | None = Field(None, description="左侧位置 (英寸)")
    y: float | None = Field(None, description="顶部位置 (英寸)")
    width: float = Field(..., description="宽度 (英寸)")
    height: float = Field(..., description="高度 (英寸)")
    bg_color: str | None = Field(None, description="填充颜色 HEX，None 且 no_fill=False 时保持默认")
    font_color: str | None = Field(None, description="文字颜色 HEX")
    font_size: int | None = Field(None, description="字号 (pt)，默认 12")
    bold: bool | None = False
    no_fill: bool = Field(False, description="无填充（透明），用于虚线分组框等")
    line_color: str | None = Field(None, description="边框颜色 HEX，None 不设置边框")
    line_width: float | None = Field(None, description="边框宽度 (pt)")
    line_dash: bool = Field(False, description="边框是否虚线")
    vertical_text: bool = Field(False, description="文字竖排（圆柱体阶段标记用）")


class SlideConnectorDef(BaseModel):
    """PPT 连接线/箭头定义（坐标单位：英寸，相对幻灯片左上角）"""
    x1: float = Field(..., description="起点 X (英寸)")
    y1: float = Field(..., description="起点 Y (英寸)")
    x2: float = Field(..., description="终点 X (英寸)")
    y2: float = Field(..., description="终点 Y (英寸)")
    color: str | None = Field(None, description="线条颜色 HEX，留空用默认")
    width: float = Field(1.5, description="线宽 (pt)")
    dash: bool = Field(False, description="是否虚线（损失反向传播等特殊箭头）")
    arrowhead: bool = Field(True, description="终点是否带箭头")
    arrowhead_style: Literal["triangle", "open", "stealth"] = Field("triangle", description="箭头样式")
    begin_arrowhead: bool = Field(False, description="起点是否带箭头")


class SlideItem(BaseModel):
    """幻灯片中的一个内容元素"""
    type: Literal["paragraph", "table", "image", "shape", "connector", "bullet_list", "numbered_list", "chart"]
    text: str | None = None
    level: int = Field(0, description="缩进层级 (0=顶级)")
    style: TextStyle | None = None
    table: SlideTableDef | None = None
    image: SlideImageDef | None = None
    shape: SlideShapeDef | None = None
    connector: SlideConnectorDef | None = Field(None, description="连接线定义，仅 connector 类型使用")
    items: list[str] | None = Field(None, description="列表项，仅 bullet_list/numbered_list 使用")
    chart: ChartDef | None = Field(None, description="图表定义，仅 chart 类型使用")


class ChartDef(BaseModel):
    """PPT 图表定义"""
    chart_type: Literal["column", "bar", "line", "pie", "doughnut", "area", "radar"] = Field(
        "column", description="图表类型"
    )
    categories: list[str] = Field(..., description="类别标签（X 轴）")
    series: list[ChartSeries] = Field(..., description="数据系列")
    has_legend: bool = Field(True, description="是否显示图例")
    title: str | None = Field(None, description="图表标题")
    x: float | None = Field(None, description="左侧位置 (英寸)，默认 0.5")
    y: float | None = Field(None, description="顶部位置 (英寸)，默认 1.5")
    width: float | None = Field(None, description="宽度 (英寸)，默认 8.0")
    height: float | None = Field(None, description="高度 (英寸)，默认 4.5")


class ChartSeries(BaseModel):
    """图表数据系列"""
    name: str = Field(..., description="系列名称")
    values: list[float] = Field(..., description="数值列表")


class ShowSettingsDef(BaseModel):
    """放映设置（写入 presentation.xml 的 p:showPr）"""
    show_type: Literal["presenter", "browse", "kiosk"] = Field("presenter", description="放映方式")
    loop: bool = Field(False, description="循环放映")
    show_animation: bool = Field(True, description="放映时播放动画")
    show_narration: bool = Field(True, description="播放旁白")
    use_timings: bool = Field(True, description="使用排练计时")
    range_start: int | None = Field(None, ge=1, description="从第 N 张开始（1 起）")
    range_end: int | None = Field(None, ge=1, description="到第 N 张结束（1 起）")


class SectionDef(BaseModel):
    """幻灯片节"""
    name: str = Field(..., description="节名称")
    slide_start: int = Field(..., ge=1, description="该节从第几张幻灯片开始（1 起）")


class SlideDef(BaseModel):
    """幻灯片定义"""
    layout: LayoutType = "content"
    title: str | None = Field(None, description="幻灯片标题")
    subtitle: str | None = Field(None, description="副标题（仅 title/section 布局）")
    body: list[SlideItem] | None = Field(None, description="主体内容")
    left: list[SlideItem] | None = Field(None, description="左侧内容（仅 two_content/comparison 布局）")
    right: list[SlideItem] | None = Field(None, description="右侧内容（仅 two_content/comparison 布局）")
    notes: str | None = Field(None, description="演讲者备注")
    bg_color: str | None = Field(None, description="幻灯片背景色 HEX")
    slide_width: float | None = Field(None, description="幻灯片宽度 (英寸)，默认 10")
    slide_height: float | None = Field(None, description="幻灯片高度 (英寸)，默认 7.5")
    footer_text: str | None = Field(None, description="页脚文字（底部居中）")
    show_slide_number: bool = Field(False, description="显示幻灯片编号（右下角）")


class PptContent(BaseModel):
    """PPT 演示文稿内容"""
    slides: list[SlideDef] = Field(..., description="幻灯片列表（至少 1 个）")
    sections: list[SectionDef] | None = Field(None, description="幻灯片节（仅创建时）")
    show_settings: ShowSettingsDef | None = Field(None, description="放映设置（仅创建时）")


# ── 框架图 ────────────────────────────────────────────────────────

class FwNode(BaseModel):
    """框架图节点"""
    id: str = Field(..., description="节点唯一标识")
    text: str = Field(..., description="节点显示文本，支持 \\n 换行")
    style: Literal[
        "rect", "rounded_rect", "ellipse", "diamond",
        "process", "decision", "data", "document", "terminator", "predefined",
        "cylinder", "parallelogram", "chevron", "right_arrow", "left_arrow",
        "callout", "cloud",
    ] = Field("rounded_rect", description="节点形状")
    fill_color: str | None = Field(None, description="填充色 HEX，留空使用阶段色")
    font_size: int = Field(12, description="字号")
    bold: bool = Field(True, description="是否加粗")


class FwEdge(BaseModel):
    """框架图连接线"""
    from_id: str = Field(..., description="起始节点 ID")
    to_id: str = Field(..., description="目标节点 ID")
    label: str | None = Field(None, description="箭头上的标签文字（可选）")
    edge_style: Literal["straight", "elbow", "curve"] = Field("straight", description="连接线样式：straight(直线)/elbow(正交)/curve(曲线)")


class FwStage(BaseModel):
    """框架图阶段分组"""
    label: str = Field(..., description="阶段标签，显示在左侧圆柱形上")
    color: str | None = Field(None, description="阶段背景色 HEX，留空使用配色方案默认")
    nodes: list[FwNode] = Field(default_factory=list, description="阶段内的节点列表")


class FrameworkDiagram(BaseModel):
    """框架图完整定义"""
    title: str = Field("", description="框架图总标题")
    subtitle: str | None = Field(None, description="框架图副标题")
    stages: list[FwStage] = Field(..., description="阶段列表，按从上到下顺序排列")
    edges: list[FwEdge] = Field(default_factory=list, description="节点间连接线，未指定的相邻节点自动生成箭头")
    color_scheme: str = Field("professional-blue", description="配色方案名称")
    node_width: float = Field(1.4, description="默认节点宽度 (英寸)")
    node_height: float = Field(0.55, description="默认节点高度 (英寸)")
    stage_gap: float = Field(0.15, description="阶段间垂直间距 (英寸)")
    arrow_color: str | None = Field(None, description="箭头颜色 HEX")
    font_cn: str = Field("宋体", description="中文字体")
    font_en: str = Field("Times New Roman", description="西文字体")
    font_size: int = Field(16, description="正文默认字号 (pt)，放不下时自动缩小")
    title_font_size: int = Field(24, description="标题字号 (pt)")


# ── 配色方案 ──────────────────────────────────────────────────────

class ColorScheme(BaseModel):
    """PPT 配色方案"""
    name: str = Field(..., description="方案名称标识")
    display_name: str = Field("", description="方案显示名称")
    description: str | None = Field(None, description="方案描述")
    title_color: str = Field(..., description="标题颜色 HEX")
    body_color: str = Field(..., description="正文颜色 HEX")
    accent_color: str = Field(..., description="强调色 HEX")
    background_color: str = Field(..., description="幻灯片背景色 HEX")
    shape_fill_color: str = Field(..., description="形状填充色 HEX")
    table_header_color: str = Field(..., description="表格表头背景色 HEX")
    title_font: str = Field("微软雅黑", description="标题字体")
    body_font: str = Field("微软雅黑", description="正文字体")
    title_size: int = Field(32, description="标题字号 (pt)")
    body_size: int = Field(18, description="正文字号 (pt)")


# ── 模板分析 ──────────────────────────────────────────────────────

class TemplateAnalysis(BaseModel):
    """PPT 模板分析结果"""
    file_path: str = Field(..., description="模板文件路径")
    slide_count: int = Field(0, description="幻灯片数量")
    slide_width: float = Field(0, description="幻灯片宽度 (英寸)")
    slide_height: float = Field(0, description="幻灯片高度 (英寸)")
    layout_names: list[str] = Field(default_factory=list, description="可用布局名称列表")
    color_schemes: list[ColorScheme] | None = Field(None, description="提取的配色方案")
    fonts_used: list[str] = Field(default_factory=list, description="模板中使用的字体")


# ── 返回结果 ──────────────────────────────────────────────────────

class CreateResult(BaseModel):
    """创建操作的返回结果"""
    success: bool
    file_path: str
    file_size: int = 0
    sheets_count: int | None = None
    slides_count: int | None = None
    message: str
    backup_path: str | None = None


class EditResult(BaseModel):
    """编辑操作的返回结果"""
    success: bool
    file_path: str
    operation_count: int = 0
    message: str
    backup_path: str | None = None


# ── PDF 相关模型 ─────────────────────────────────────────────────

class PdfTableDef(BaseModel):
    """PDF 表格定义"""
    headers: list[str] = Field(default_factory=list, description="表头行")
    rows: list[list[str]] = Field(default_factory=list, description="数据行")
    column_widths: list[float] | None = Field(None, description="列宽比例 (0~1)，如 [0.3, 0.4, 0.3]")


class PdfSectionItem(BaseModel):
    """PDF 文档的一个内容块"""
    type: Literal["heading", "paragraph", "table", "image", "page_break", "bullet_list", "numbered_list"]
    text: str | None = Field(None, description="文字内容")
    level: int | None = Field(None, description="标题层级 (1-6)，仅 heading 类型使用")
    style: ParagraphStyle | None = Field(None, description="段落样式")
    items: list[str] | None = Field(None, description="列表项，仅 bullet_list/numbered_list 使用")
    table: PdfTableDef | None = Field(None, description="表格定义，仅 table 类型使用")
    image: ImageDef | None = Field(None, description="图片定义，仅 image 类型使用")


class PdfContent(BaseModel):
    """PDF 文档内容"""
    sections: list[PdfSectionItem] = Field(..., description="文档内容块列表")


class PdfPageSetup(BaseModel):
    """PDF 页面设置"""
    paper_size: Literal["A4", "Letter", "A3", "B5"] = "A4"
    orientation: Literal["portrait", "landscape"] = "portrait"
    margin_top: float = Field(10, description="上边距 (mm)")
    margin_bottom: float = Field(10, description="下边距 (mm)")
    margin_left: float = Field(10, description="左边距 (mm)")
    margin_right: float = Field(10, description="右边距 (mm)")


class PdfFontConfig(BaseModel):
    """PDF 字体配置"""
    cn_font: str = Field("", description="中文字体名称，如 'simhei'、'simsun'，留空自动检测")
    cn_font_path: str | None = Field(None, description="自定义中文字体文件路径 (.ttf/.ttc)")
    en_font: str = Field("Helvetica", description="西文字体: Helvetica/Courier/Times")
    cn_font_size: int | str = Field("小四", description="正文字号")
    heading_font_size: int | str = Field("三号", description="标题字号")


class OfficeOperation(BaseModel):
    """A path-addressed Office document mutation used by batch editing."""

    op: Literal["set", "add", "remove", "move", "replace"]
    path: str = Field(..., description="目标路径，如 /body/p[1]、/Sheet1/A1 或 /slide[1]/shape[2]")
    element_type: str | None = Field(None, description="add 操作的元素类型，如 paragraph、sheet、slide、textbox")
    props: dict[str, Any] = Field(default_factory=dict, description="操作属性")
    target_path: str | None = Field(None, description="move 操作的目标路径")
    position: Literal["before", "after"] = Field("after", description="move 相对目标的位置")



# ── 活动文档内容替换 ─────────────────────────────────────────────

class ReplaceSection(BaseModel):
    """Word 活动文档的按标题替换节。"""

    heading_text: str = Field(..., description="定位标题文字")
    heading_match: Literal["exact", "contains"] = Field("contains", description="标题匹配方式")
    end_heading_text: str | None = Field(None, description="结束标题文字（不含）；为空则到下一个同级或更高级标题")
    keep_heading: bool = Field(True, description="是否保留标题段落")
    lines: list[str] = Field(default_factory=list, description="替换后的内容行，每行一段")
    font: FontConfig | None = Field(None, description="新内容字体")


class TextReplacement(BaseModel):
    """活动文档查找替换（Word/PPT/Excel 通用）。"""

    find: str = Field(..., description="待查找文本")
    replace: str = Field("", description="替换文本")
    match_case: bool = Field(False, description="是否区分大小写")


# ── 新增入口专用模型 ─────────────────────────────────────────────

class WordCommentDef(BaseModel):
    """Word 批注（锚定到正文文字或段落）"""
    text: str = Field(..., description="批注正文")
    author: str = Field("office-writer", description="批注作者")
    initials: str = Field("OW", description="作者缩写")
    anchor_text: str | None = Field(None, description="锚定文字：命中的第一处 run 上添加批注")
    paragraph_index: int | None = Field(None, ge=1, description="锚定到第 N 个正文段落（1 起）；与 anchor_text 二选一")
    match_case: bool = Field(False, description="anchor_text 是否区分大小写")

    @model_validator(mode="after")
    def _check_anchor(self):
        if not self.anchor_text and self.paragraph_index is None:
            raise ValueError("WordCommentDef 必须提供 anchor_text 或 paragraph_index")
        return self


class PptHyperlinkDef(BaseModel):
    """PPT 超链接（按形状或形状内文字定位）"""
    slide: int = Field(..., ge=1, description="幻灯片序号（1 起）")
    url: str = Field(..., description="链接地址")
    shape_index: int | None = Field(None, ge=1, description="形状序号（1 起，对应 inspect 路径 /slide[n]/shape[i]）")
    shape_text: str | None = Field(None, description="按形状文字包含匹配定位")
    text_match: str | None = Field(None, description="仅给形状内匹配的文字 run 加链接；留空则整个形状可点击")
    match_case: bool = Field(False, description="text_match 是否区分大小写")

    @model_validator(mode="after")
    def _check_target(self):
        if self.shape_index is None and not self.shape_text:
            raise ValueError("PptHyperlinkDef 必须提供 shape_index 或 shape_text")
        return self
