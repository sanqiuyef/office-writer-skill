# Vendored from mcp_office_writer/ppt_writer.py (mcp-office-writer 0.6.0, 2026-11).
# Upstream path: mcp_office_writer/ppt_writer.py (repository-relative).
# Modifications: cast layout geometry to int for python-pptx 1.0.2 compatibility; added read-only audit_slide_layout for out-of-bounds/overlap warnings; added reflow_slide_layout for opt-in automatic repositioning/resizing; added table/chart geometry, image crop/hyperlinks, shape hyperlinks, placeholder-column targeting, real bullet/number markers, and .potx template loading via a content-type-normalized working copy.

"""
mcp-office-writer - PowerPoint (.pptx) 写入/编辑逻辑
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import TYPE_CHECKING

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

from . import office_com
from .utils import atomic_save

if TYPE_CHECKING:
    from .models import SlideDef, SlideItem, SlideTableDef, SlideImageDef
    from .models import SlideShapeDef, TextStyle, CreateResult, EditResult
    from .models import ChartDef, FrameworkDiagram


# ── 预设配色方案 ──────────────────────────────────────────────────

_PRESET_COLOR_SCHEMES: dict[str, dict] = {
    "professional-blue": {
        "title_color": "000000", "body_color": "333333",
        "accent_color": "2563EB", "background_color": "F8FAFC",
        "shape_fill_color": "DBEAFE", "table_header_color": "1A3C6D",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "商务蓝", "description": "专业商务风格，深蓝主色配合浅灰背景",
    },
    "modern-dark": {
        "title_color": "FFFFFF", "body_color": "D1D5DB",
        "accent_color": "60A5FA", "background_color": "1F2937",
        "shape_fill_color": "374151", "table_header_color": "111827",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "暗夜模式", "description": "现代深色主题，白字深灰底，适合大屏演示",
    },
    "warm-earth": {
        "title_color": "7C2D12", "body_color": "44403C",
        "accent_color": "D97706", "background_color": "FFF7ED",
        "shape_fill_color": "FED7AA", "table_header_color": "7C2D12",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "暖色大地", "description": "温暖自然风格，橙色与棕色搭配",
    },
    "elegant-minimal": {
        "title_color": "18181B", "body_color": "52525B",
        "accent_color": "8B5CF6", "background_color": "FFFFFF",
        "shape_fill_color": "F4F4F5", "table_header_color": "18181B",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 28, "body_size": 16,
        "display_name": "极简优雅", "description": "黑白灰极简风格，紫色点缀，适合学术/设计",
    },
    "vibrant": {
        "title_color": "BE185D", "body_color": "27272A",
        "accent_color": "14B8A6", "background_color": "FFF1F2",
        "shape_fill_color": "FCE7F3", "table_header_color": "BE185D",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 34, "body_size": 20,
        "display_name": "活力缤纷", "description": "高对比度配色，玫红+青绿，适合创意/营销",
    },
    "corporate-green": {
        "title_color": "064E3B", "body_color": "374151",
        "accent_color": "10B981", "background_color": "ECFDF5",
        "shape_fill_color": "D1FAE5", "table_header_color": "064E3B",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "企业绿", "description": "稳重大气，深绿主色配合浅绿背景",
    },
    "tech-purple": {
        "title_color": "4C1D95", "body_color": "374151",
        "accent_color": "7C3AED", "background_color": "F5F3FF",
        "shape_fill_color": "EDE9FE", "table_header_color": "4C1D95",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "科技紫", "description": "科技感紫色渐变，适合互联网/AI/创新主题",
    },
    "classic-red": {
        "title_color": "991B1B", "body_color": "44403C",
        "accent_color": "DC2626", "background_color": "FEF2F2",
        "shape_fill_color": "FEE2E2", "table_header_color": "991B1B",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 34, "body_size": 20,
        "display_name": "经典中国红", "description": "庄重典雅的中国红，适合党政/党建/传统节日主题",
    },
    "fresh-nature": {
        "title_color": "166534", "body_color": "374151",
        "accent_color": "22C55E", "background_color": "F0FDF4",
        "shape_fill_color": "DCFCE7", "table_header_color": "166534",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 30, "body_size": 18,
        "display_name": "清新自然", "description": "清爽绿色系，适合环保/健康/农业/教育主题",
    },
    "ocean-blue": {
        "title_color": "000000", "body_color": "334155",
        "accent_color": "0EA5E9", "background_color": "F0F9FF",
        "shape_fill_color": "E0F2FE", "table_header_color": "0C4A6E",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 18,
        "display_name": "海洋蓝", "description": "清爽海洋蓝，适合医疗/教育/水利/科技主题",
    },
    "sunset-orange": {
        "title_color": "9A3412", "body_color": "44403C",
        "accent_color": "F97316", "background_color": "FFF7ED",
        "shape_fill_color": "FFEDD5", "table_header_color": "9A3412",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 32, "body_size": 20,
        "display_name": "日落暖橙", "description": "温暖活力的橙色系，适合餐饮/旅游/创意/生活主题",
    },
    "monochrome-gray": {
        "title_color": "1E293B", "body_color": "475569",
        "accent_color": "64748B", "background_color": "F8FAFC",
        "shape_fill_color": "F1F5F9", "table_header_color": "1E293B",
        "title_font": "微软雅黑", "body_font": "微软雅黑",
        "title_size": 28, "body_size": 16,
        "display_name": "高级灰", "description": "克制高级灰，适合设计/建筑/金融/高端品牌主题",
    },
}


ALIGN_MAP = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
}


def _hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _apply_text_style(run, style: TextStyle | None):
    if style is None:
        return
    if style.font_name:
        run.font.name = style.font_name
    if style.font_size:
        run.font.size = Pt(style.font_size)
    if style.bold is not None:
        run.font.bold = style.bold
    if style.italic is not None:
        run.font.italic = style.italic
    if style.color:
        run.font.color.rgb = _hex_to_rgb(style.color)


def _set_text_frame_alignment(tf, alignment_str: str | None):
    if alignment_str and alignment_str in ALIGN_MAP:
        tf.paragraphs[0].alignment = ALIGN_MAP[alignment_str]


def _add_paragraph_item(tf, item: SlideItem):
    """向文本框添加段落内容"""
    from .models import TextStyle

    if not item.text:
        return
    p = tf.add_paragraph()
    p.text = item.text
    p.level = item.level
    style = item.style or TextStyle()
    _apply_text_style(p.runs[0], style) if p.runs else None
    if style.alignment:
        p.alignment = ALIGN_MAP.get(style.alignment, PP_ALIGN.LEFT)


def _add_slide_table(slide, td: SlideTableDef, default_left=None, default_top=None,
                     default_width=None, default_height=None):
    """在幻灯片上添加表格；模型未给位置时使用调用方传入的栏位默认值。"""
    rows = len(td.rows) + (1 if td.headers else 0)
    cols = max(len(td.headers) if td.headers else 0,
               max((len(r) for r in td.rows), default=0))
    if rows == 0 or cols == 0:
        return

    left = Inches(td.x) if td.x is not None else (default_left or Inches(0.5))
    top = Inches(td.y) if td.y is not None else (default_top or Inches(2.0))
    width = Inches(td.width) if td.width is not None else (default_width or Inches(7.0))
    height = Inches(td.height) if td.height is not None else (default_height or Inches(2.5))
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    if td.headers:
        for ci, h in enumerate(td.headers):
            cell = table.cell(0, ci)
            cell.text = str(h)
            if td.header_bold:
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

    for ri, row_data in enumerate(td.rows, 1 if td.headers else 0):
        for ci, val in enumerate(row_data):
            if ci < cols:
                table.cell(ri, ci).text = str(val)

    if td.column_widths:
        target_width = td.width if td.width is not None else 7.0
        widths = [max(float(w), 0.1) for w in td.column_widths[:cols]]
        if len(widths) < cols:
            missing = cols - len(widths)
            remaining = max(target_width - sum(widths), 0.0)
            # 缺省列均分剩余宽度；剩余不足时仍保证每列可用的最小宽度
            fill_width = max(remaining / missing, 0.1)
            widths.extend([fill_width] * missing)
        total = sum(widths)
        if total > 0 and abs(total - target_width) > 0.01:
            # python-pptx 的表格宽度等于列宽之和，必须归一化到声明宽度。
            factor = target_width / total
            widths = [width * factor for width in widths]
        for index, column_width in enumerate(widths):
            table.columns[index].width = Inches(column_width)


def _add_slide_image(slide, img_def: SlideImageDef):
    """在幻灯片上添加图片（支持本地 path 或 URL 下载）"""
    if img_def.url:
        from .utils import download_image
        try:
            file_path = download_image(img_def.url)
        except Exception:
            # 下载失败时静默跳过（避免 PPT 创建中断）
            return
    elif img_def.path:
        file_path = img_def.path
    else:
        return

    if not os.path.exists(file_path):
        return
    left = Inches(img_def.x or 0.5)
    top = Inches(img_def.y or 1.5)
    width = Inches(img_def.width or 4.0)
    height = Inches(img_def.height) if img_def.height else None
    if height:
        picture = slide.shapes.add_picture(file_path, left, top, width, height)
    else:
        picture = slide.shapes.add_picture(file_path, left, top, width)

    # 裁剪比例直接映射到 python-pptx 的 crop_*（0~1 表示裁掉的比例）
    if img_def.crop_left is not None:
        picture.crop_left = img_def.crop_left
    if img_def.crop_right is not None:
        picture.crop_right = img_def.crop_right
    if img_def.crop_top is not None:
        picture.crop_top = img_def.crop_top
    if img_def.crop_bottom is not None:
        picture.crop_bottom = img_def.crop_bottom

    if img_def.hyperlink_url:
        picture.click_action.hyperlink.address = img_def.hyperlink_url


def _add_slide_shape(slide, shape_def: SlideShapeDef):
    """在幻灯片上添加形状"""
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.dml import MSO_LINE_DASH_STYLE

    shape_map = {
        "rect": MSO_SHAPE.RECTANGLE,
        "rounded_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
        "oval": MSO_SHAPE.OVAL,
        "arrow": MSO_SHAPE.RIGHT_ARROW,
        "cylinder": MSO_SHAPE.CAN,
        "diamond": MSO_SHAPE.DIAMOND,
    }

    shape_type = shape_map.get(shape_def.shape_type, MSO_SHAPE.RECTANGLE)
    left = Inches(shape_def.x or 0.5)
    top = Inches(shape_def.y or 1.5)
    width = Inches(shape_def.width)
    height = Inches(shape_def.height)

    shape = slide.shapes.add_shape(shape_type, left, top, width, height)

    if shape_def.no_fill:
        shape.fill.background()
    elif shape_def.bg_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _hex_to_rgb(shape_def.bg_color)

    if shape_def.line_color:
        shape.line.color.rgb = _hex_to_rgb(shape_def.line_color)
    if shape_def.line_width is not None:
        shape.line.width = Pt(shape_def.line_width)
    if shape_def.line_dash:
        shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH

    if shape_def.text:
        tf = shape.text_frame
        if shape_def.vertical_text:
            # python-pptx 无竖排 API，直接注入 bodyPr vert 属性
            from pptx.oxml.ns import qn
            body_pr = tf._txBody.find(qn("a:bodyPr"))
            if body_pr is not None:
                body_pr.set("vert", "vert270")
        # 横排文字按框宽自动换行（避免长公式/长句溢出）；竖排不换行
        tf.word_wrap = not shape_def.vertical_text
        tf.text = shape_def.text
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        for run in tf.paragraphs[0].runs:
            if shape_def.font_color:
                run.font.color.rgb = _hex_to_rgb(shape_def.font_color)
            if shape_def.bold:
                run.font.bold = True
            run.font.size = Pt(shape_def.font_size or 12)

    if shape_def.hyperlink_url:
        shape.click_action.hyperlink.address = shape_def.hyperlink_url


def _add_slide_connector(slide, connector_def):
    """在幻灯片上添加连接线/箭头"""
    from pptx.enum.shapes import MSO_CONNECTOR_TYPE
    from pptx.enum.dml import MSO_LINE_DASH_STYLE

    connector = slide.shapes.add_connector(
        MSO_CONNECTOR_TYPE.STRAIGHT,
        Inches(connector_def.x1), Inches(connector_def.y1),
        Inches(connector_def.x2), Inches(connector_def.y2),
    )
    if connector_def.color:
        connector.line.color.rgb = _hex_to_rgb(connector_def.color)
    connector.line.width = Pt(connector_def.width)
    if connector_def.dash:
        connector.line.dash_style = MSO_LINE_DASH_STYLE.DASH

    # 箭头挂在 a:ln 上，而 a:ln 位于 p:spPr 之下；直接 find(qn("a:ln")) 匹配不到，
    # 必须经 LineFormat._get_or_add_ln() 取（连接线颜色/宽度已创建该元素）。
    from pptx.oxml.ns import qn
    from lxml import etree
    line_elem = connector.line._get_or_add_ln()
    if connector_def.arrowhead:
        tail = line_elem.find(qn("a:tailEnd"))
        if tail is None:
            tail = etree.SubElement(line_elem, qn("a:tailEnd"))
        tail.set("type", {"triangle": "triangle", "open": "oval", "stealth": "stealth"}.get(
            connector_def.arrowhead_style, "triangle"))
        tail.set("w", "med")
        tail.set("len", "med")
    if connector_def.begin_arrowhead:
        head = line_elem.find(qn("a:headEnd"))
        if head is None:
            head = etree.SubElement(line_elem, qn("a:headEnd"))
        head.set("type", "triangle")
        head.set("w", "med")
        head.set("len", "med")


def _set_list_marker(paragraph, numbered: bool) -> None:
    """给段落注入真正的列表标记（项目符号或自动编号）。

    python-pptx 没有列表 API，直接写 a:pPr 下的 a:buChar / a:buAutoNum；
    a:pPr 子元素有固定顺序，必须插在 a:buBlip/a:tabLst/a:defRPr/a:extLst 之前，
    并移除布局可能带入的 a:buNone。
    """
    from lxml import etree
    from pptx.oxml.ns import qn

    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buNone", "a:buAutoNum", "a:buChar"):
        for old in pPr.findall(qn(tag)):
            pPr.remove(old)
    # 列表文本需要悬挂缩进，否则符号与正文从同一 x 开始、首字被符号覆盖。
    level = max(int(getattr(paragraph, "level", 0) or 0), 0)
    pPr.set("marL", str(342900 * (level + 1)))
    pPr.set("indent", "-342900")

    if numbered:
        marker = etree.Element(qn("a:buAutoNum"))
        marker.set("type", "arabicPeriod")
    else:
        marker = etree.Element(qn("a:buChar"))
        marker.set("char", "•")
    pPr.insert_element_before(marker, "a:buBlip", "a:tabLst", "a:defRPr", "a:extLst")


def _build_slide_body(
    slide,
    items: list[SlideItem] | None,
    default_left,
    default_top,
    default_width,
    default_height,
    placeholder_idx: int | None = 1,
):
    """向幻灯片主体区域写入内容元素"""
    if not items:
        return

    for item in items:
        if item.type == "paragraph":
            # 将段落添加到现有的标题/正文占位符；双栏时按 placeholder_idx 选栏
            text_frame = None
            if slide.placeholders:
                if placeholder_idx is not None:
                    for ph in slide.placeholders:
                        if ph.placeholder_format.idx == placeholder_idx:
                            text_frame = ph.text_frame
                            break
                if text_frame is None:
                    for ph in slide.placeholders:
                        if ph.placeholder_format.idx == 1:  # 正文
                            text_frame = ph.text_frame
                            break
                if text_frame is None and len(slide.placeholders) > 1:
                    text_frame = slide.placeholders[1].text_frame

            if text_frame:
                _add_paragraph_item(text_frame, item)
            else:
                # 没有占位符时手动加文本框
                txBox = slide.shapes.add_textbox(default_left, default_top, default_width, default_height)
                tf = txBox.text_frame
                tf.text = item.text or ""

        elif item.type == "table" and item.table:
            _add_slide_table(
                slide, item.table,
                default_left, Inches(2.0), default_width, Inches(2.5),
            )

        elif item.type == "image" and item.image:
            _add_slide_image(slide, item.image)

        elif item.type == "shape" and item.shape:
            _add_slide_shape(slide, item.shape)

        elif item.type == "connector" and item.connector:
            _add_slide_connector(slide, item.connector)

        elif item.type == "bullet_list" or item.type == "numbered_list":
            # 添加项目符号/编号列表
            txBox = slide.shapes.add_textbox(default_left, Inches(1.8), default_width, default_height)
            tf = txBox.text_frame
            tf.word_wrap = True
            for i, list_item in enumerate(item.items or []):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = list_item
                p.level = item.level
                p.alignment = PP_ALIGN.LEFT
                if item.type == "bullet_list":
                    p.level = 0
                _set_list_marker(p, numbered=item.type == "numbered_list")
                if item.style:
                    _apply_text_style(p.runs[0], item.style) if p.runs else None

        elif item.type == "chart" and item.chart:
            _add_slide_chart(
                slide, item.chart,
                default_left, default_top, default_width, default_height,
            )


def _add_slide_chart(slide, chart_def: ChartDef, default_left=None, default_top=None,
                     default_width=None, default_height=None):
    """在幻灯片上添加图表（位置与尺寸优先取模型定义，否则用调用方栏位默认值）。"""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    chart_type_map = {
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
        "doughnut": XL_CHART_TYPE.DOUGHNUT,
        "area": XL_CHART_TYPE.AREA,
        "radar": XL_CHART_TYPE.RADAR,
    }
    ctype = chart_type_map.get(chart_def.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    chart_data = CategoryChartData()
    chart_data.categories = chart_def.categories
    for series in chart_def.series:
        chart_data.add_series(series.name, series.values)

    left = Inches(chart_def.x) if chart_def.x is not None else (default_left or Inches(0.5))
    top = Inches(chart_def.y) if chart_def.y is not None else (default_top or Inches(1.5))
    width = Inches(chart_def.width) if chart_def.width is not None else (default_width or Inches(8.0))
    height = Inches(chart_def.height) if chart_def.height is not None else (default_height or Inches(4.5))
    chart_frame = slide.shapes.add_chart(ctype, left, top, width, height, chart_data)
    chart = chart_frame.chart

    if chart_def.title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = chart_def.title
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)

    if chart_def.has_legend:
        chart.has_legend = True
        from pptx.enum.chart import XL_LEGEND_POSITION
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM


def _set_slide_bg(slide, bg_color: str | None):
    """设置幻灯片背景色"""
    if bg_color is None:
        return
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = _hex_to_rgb(bg_color)


def _create_slide(prs, slide_def: SlideDef, layout_idx: int):
    """创建一张幻灯片"""

    layout = prs.slide_layouts[layout_idx]
    slide = prs.slides.add_slide(layout)

    # 设置背景
    _set_slide_bg(slide, slide_def.bg_color)

    # 标题
    if slide_def.title and slide.shapes.title:
        slide.shapes.title.text = slide_def.title

    # 副标题 (title/section 布局)
    if slide_def.subtitle:
        from pptx.enum.shapes import PP_PLACEHOLDER
        for ph in slide.placeholders:
            if ph.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                ph.text = slide_def.subtitle
                break

    # 主体内容
    if slide_def.body:
        _build_slide_body(
            slide, slide_def.body,
            Inches(0.5), Inches(1.5), Inches(8.0), Inches(4.5),
            placeholder_idx=1,
        )

    # 双栏布局：左栏落 idx=1、右栏落 idx=2（版式没有 idx=2 时回退原有逻辑）
    if slide_def.left or slide_def.right:
        _build_slide_body(
            slide, slide_def.left,
            Inches(0.5), Inches(1.5), Inches(4.0), Inches(4.5),
            placeholder_idx=1,
        )
        _build_slide_body(
            slide, slide_def.right,
            Inches(5.0), Inches(1.5), Inches(4.0), Inches(4.5),
            placeholder_idx=2,
        )

    # 演讲者备注
    if slide_def.notes:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = slide_def.notes

    return slide


_POTX_TEMPLATE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
)
_PPTX_PRESENTATION_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
)


def _load_presentation(file_path: str):
    """加载演示文稿；.potx 先转成临时 .pptx 工作副本再交给 python-pptx。

    python-pptx 的 Package.open 只接受 presentation.main 内容类型，标准 .potx
    的 [Content_Types].xml 标记为 template.main，直接加载会抛 ValueError。
    这里把工作副本中的 template.main 改写为 presentation.main，加载完成后立即
    删除临时文件，源 .potx 保持不变。
    """
    if os.path.splitext(file_path)[1].lower() != ".potx":
        return Presentation(file_path)

    import tempfile
    import zipfile

    fd, temp_path = tempfile.mkstemp(prefix=".potx_", suffix=".pptx")
    os.close(fd)
    try:
        with zipfile.ZipFile(file_path) as source, zipfile.ZipFile(
            temp_path, "w", zipfile.ZIP_DEFLATED
        ) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    text = data.decode("utf-8").replace(
                        _POTX_TEMPLATE_CONTENT_TYPE, _PPTX_PRESENTATION_CONTENT_TYPE
                    )
                    data = text.encode("utf-8")
                target.writestr(item, data)
        # python-pptx 在加载阶段就把 part 读入内存，返回后即可删除临时文件
        return Presentation(temp_path)
    finally:
        try:
            os.remove(temp_path)
        except OSError:
            pass


def create_powerpoint(
    output_path: str,
    slides: list[SlideDef],
    template: str | None = None,
    if_exists: str = "overwrite",
) -> CreateResult:
    """创建 PowerPoint 演示文稿"""
    from .models import CreateResult
    from .utils import atomic_save, ensure_output_dir, get_file_size, get_ppt_layout_index, resolve_if_exists, validate_input_file, validate_output_path

    try:
        validate_output_path(output_path, (".pptx",))
        ensure_output_dir(output_path)
        resolved_path = resolve_if_exists(output_path, if_exists)
    except (OSError, ValueError) as e:
        return CreateResult(success=False, file_path=output_path, message=str(e))

    try:
        if template:
            validate_input_file(template, (".pptx", ".potx"))
            prs = _load_presentation(template)
        else:
            prs = Presentation()

        # 支持自定义幻灯片尺寸（取第一个指定宽高的 slide_def）
        for sd in slides:
            if sd.slide_width and sd.slide_height:
                prs.slide_width = Inches(sd.slide_width)
                prs.slide_height = Inches(sd.slide_height)
                break

        for slide_def in slides:
            layout_idx = get_ppt_layout_index(slide_def.layout)
            _create_slide(prs, slide_def, layout_idx)

        backup_path = atomic_save(prs.save, resolved_path)
        size = get_file_size(resolved_path)

        return CreateResult(
            success=True,
            file_path=resolved_path,
            file_size=size,
            slides_count=len(slides),
            message=f"PowerPoint 演示文稿创建成功: {resolved_path}",
            backup_path=backup_path,
        )
    except Exception as e:
        return CreateResult(
            success=False,
            file_path=resolved_path,
            message=f"PowerPoint 演示文稿创建失败: {str(e)}",
        )


def edit_powerpoint(
    file_path: str,
    append_slides: list[SlideDef],
) -> EditResult:
    """向已有 PPT 追加幻灯片"""
    from .models import EditResult
    from .utils import atomic_save, get_ppt_layout_index, validate_input_file

    # 优先尝试 COM 活动文档编辑（实时交互查看效果）
    from .office_com import try_edit_open_powerpoint
    com_result = try_edit_open_powerpoint(file_path, append_slides)
    if com_result is not None:
        return com_result

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return EditResult(
            success=False,
            file_path=file_path,
            message=str(e),
        )

    try:
        prs = Presentation(file_path)
        count = 0
        for slide_def in append_slides:
            layout_idx = get_ppt_layout_index(slide_def.layout)
            _create_slide(prs, slide_def, layout_idx)
            count += 1
        backup_path = atomic_save(prs.save, file_path)
        return EditResult(
            success=True,
            file_path=file_path,
            operation_count=count,
            message=f"已追加 {count} 张幻灯片到 PPT",
            backup_path=backup_path,
        )
    except Exception as e:
        return EditResult(
            success=False,
            file_path=file_path,
            message=f"编辑 PowerPoint 失败: {str(e)}",
        )


# ── 配色方案工具 ──────────────────────────────────────────────────

def list_color_schemes() -> list[dict]:
    """列出所有预设配色方案及其详情。"""
    results = []
    for name, config in _PRESET_COLOR_SCHEMES.items():
        results.append({"name": name, "display_name": config["display_name"], "description": config["description"],
                        "preview": {"title_color": config["title_color"], "body_color": config["body_color"],
                                     "accent_color": config["accent_color"], "background_color": config["background_color"]}})
    return results


def apply_color_to_slides(
    file_path: str,
    scheme_name: str = "professional-blue",
) -> dict:
    """对已有 PPT 文件中的所有幻灯片应用指定配色方案。

    将根据配色方案更新每张幻灯片的标题颜色、正文颜色、背景色、
    形状填充色和表格表头颜色。
    """
    from .utils import atomic_save, validate_input_file

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return {"success": False, "file_path": file_path, "message": str(e)}

    if scheme_name not in _PRESET_COLOR_SCHEMES:
        names = ", ".join(_PRESET_COLOR_SCHEMES)
        return {"success": False, "file_path": file_path, "message": f"未知配色方案: {scheme_name}。可用: {names}"}

    scheme = _PRESET_COLOR_SCHEMES[scheme_name]
    try:
        prs = Presentation(file_path)
        slides_updated = 0

        for slide in prs.slides:
            # 背景色
            if scheme["background_color"]:
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = _hex_to_rgb(scheme["background_color"])

            for shape in slide.shapes:
                # 表格：表头行白字 + 表头背景色，数据行套正文样式
                # （GraphicFrame 没有 has_fill/has_text_frame 属性，先单独处理）
                if getattr(shape, "has_table", False):
                    table = shape.table
                    for row_index, row in enumerate(table.rows):
                        for cell in row.cells:
                            for paragraph in cell.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    if row_index == 0:
                                        run.font.color.rgb = _hex_to_rgb("FFFFFF")
                                    else:
                                        run.font.color.rgb = _hex_to_rgb(scheme["body_color"])
                                        run.font.name = scheme["body_font"]
                                        run.font.size = Pt(scheme["body_size"])
                            if row_index == 0:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = _hex_to_rgb(scheme["table_header_color"])
                    continue

                # 标题
                if shape.has_text_frame and shape == slide.shapes.title:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = _hex_to_rgb(scheme["title_color"])
                            run.font.name = scheme["title_font"]
                            run.font.size = Pt(scheme["title_size"])

                elif shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = _hex_to_rgb(scheme["body_color"])
                            run.font.name = scheme["body_font"]
                            run.font.size = Pt(scheme["body_size"])

                # 形状填充
                if getattr(shape, "has_fill", False) and not shape.has_text_frame:
                    try:
                        shape.fill.solid()
                        shape.fill.fore_color.rgb = _hex_to_rgb(scheme["shape_fill_color"])
                    except Exception:
                        pass

            slides_updated += 1

        atomic_save(prs.save, file_path)
        return {
            "success": True,
            "file_path": file_path,
            "scheme": scheme_name,
            "slides_updated": slides_updated,
            "message": f"已对 {slides_updated} 张幻灯片应用「{scheme['display_name']}」配色方案",
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "message": f"配色应用失败: {e}"}


# ── 布局优化 ──────────────────────────────────────────────────────

def optimize_slide_layout(
    file_path: str,
) -> dict:
    """对已有 PPT 中的幻灯片进行布局优化。

    自动调整内容元素的位置和间距：
    - 多个元素时均匀分布垂直间距
    - 调整文本框大小自适应内容
    - 对齐标题位置
    """
    from .utils import atomic_save, validate_input_file

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return {"success": False, "file_path": file_path, "message": str(e)}

    try:
        prs = Presentation(file_path)
        slides_optimized = 0

        for slide in prs.slides:
            shapes_with_text = []
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text.strip():
                    shapes_with_text.append(shape)

            if not shapes_with_text:
                continue

            # 按 Y 坐标排序
            shapes_with_text.sort(key=lambda s: s.top)

            # 计算可用区域
            slide_height = prs.slide_height
            top_margin = Inches(0.8)
            bottom_margin = Inches(0.5)
            usable_height = slide_height - top_margin - bottom_margin

            n = len(shapes_with_text)
            if n <= 1:
                slides_optimized += 1
                continue

            # 均匀分布垂直间距
            segment_height = usable_height / n
            for i, shape in enumerate(shapes_with_text):
                new_top = top_margin + segment_height * i

                # 调整文本框高度
                if shape.has_text_frame:
                    text_lines = len(shape.text_frame.text.split('\n'))
                    line_height = Pt(24)
                    new_height = max(Inches(0.6), line_height * max(text_lines, 1) + Pt(8))
                    shape.height = int(new_height)

                shape.top = int(new_top)

            slides_optimized += 1

        atomic_save(prs.save, file_path)
        return {
            "success": True,
            "file_path": file_path,
            "slides_optimized": slides_optimized,
            "message": f"已优化 {slides_optimized} 张幻灯片的布局",
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "message": f"布局优化失败: {e}"}


# ── 布局检查 ─────────────────────────────────────────────────────

_LAYOUT_EPSILON_EMU = 9144  # 0.01 英寸，容忍舍入误差


def _shape_box(shape) -> tuple[int, int, int, int] | None:
    """Return (left, top, width, height) in EMU, or None for unusable geometry."""
    values = (shape.left, shape.top, shape.width, shape.height)
    if any(value is None for value in values):
        return None
    left, top, width, height = (int(value) for value in values)
    if width <= 0 or height <= 0:
        return None
    return left, top, width, height


def _shape_label(shape, index: int) -> str:
    name = str(getattr(shape, "name", "") or "").strip()
    return f"#{index} {name}".strip()


def audit_slide_layout(file_path: str) -> dict:
    """只读检查 PPT 形状越界与重叠，返回 warnings，不修改文件。

    越界按幻灯片边界判断；重叠按形状外接矩形判断。检测阈值 0.01 英寸，
    用于过滤舍入误差。``--fix-spacing`` 仍走 optimize_slide_layout 的原有重排逻辑。
    """
    from .utils import validate_input_file

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return {"success": False, "file_path": file_path, "message": str(e)}

    try:
        prs = Presentation(file_path)
        slide_width = int(prs.slide_width)
        slide_height = int(prs.slide_height)
        warnings: list[dict] = []
        shape_count = 0
        slides_checked = 0

        for slide_index, slide in enumerate(prs.slides, start=1):
            slides_checked += 1
            boxes: list[tuple[str, tuple[int, int, int, int], int]] = []
            for shape_index, shape in enumerate(slide.shapes, start=1):
                box = _shape_box(shape)
                if box is None:
                    continue
                shape_count += 1
                left, top, width, height = box
                label = _shape_label(shape, shape_index)
                shape_id = int(getattr(shape, "shape_id", 0) or 0)
                overflows = []
                if left < -_LAYOUT_EPSILON_EMU:
                    overflows.append(f"左边界超出 {-left / 914400:.2f}in")
                if top < -_LAYOUT_EPSILON_EMU:
                    overflows.append(f"上边界超出 {-top / 914400:.2f}in")
                if left + width > slide_width + _LAYOUT_EPSILON_EMU:
                    overflows.append(
                        f"右边界超出 {(left + width - slide_width) / 914400:.2f}in"
                    )
                if top + height > slide_height + _LAYOUT_EPSILON_EMU:
                    overflows.append(
                        f"下边界超出 {(top + height - slide_height) / 914400:.2f}in"
                    )
                if overflows:
                    warnings.append({
                        "type": "out_of_bounds",
                        "slide": slide_index,
                        "shape": label,
                        "shape_id": shape_id,
                        "detail": "；".join(overflows),
                    })
                boxes.append((label, box, shape_id))

            for first in range(len(boxes)):
                label_a, (left_a, top_a, width_a, height_a), id_a = boxes[first]
                for second in range(first + 1, len(boxes)):
                    label_b, (left_b, top_b, width_b, height_b), id_b = boxes[second]
                    overlap_x = min(left_a + width_a, left_b + width_b) - max(left_a, left_b)
                    overlap_y = min(top_a + height_a, top_b + height_b) - max(top_a, top_b)
                    if (
                        overlap_x > _LAYOUT_EPSILON_EMU
                        and overlap_y > _LAYOUT_EPSILON_EMU
                    ):
                        warnings.append({
                            "type": "overlap",
                            "slide": slide_index,
                            "shapes": [label_a, label_b],
                            "shape_ids": [id_a, id_b],
                            "detail": (
                                f"重叠约 {overlap_x / 914400:.2f}×"
                                f"{overlap_y / 914400:.2f}in"
                            ),
                        })

        if warnings:
            message = (
                f"布局检查完成：检查 {slides_checked} 张幻灯片、{shape_count} 个形状，"
                f"发现 {len(warnings)} 条警告（未自动重排；自动重排用 --reflow，"
                "纵向间距优化用 --fix-spacing）"
            )
        else:
            message = (
                f"布局检查完成：检查 {slides_checked} 张幻灯片、{shape_count} 个形状，"
                "未发现越界或重叠"
            )
        return {
            "success": True,
            "file_path": file_path,
            "slides_checked": slides_checked,
            "shape_count": shape_count,
            "warning_count": len(warnings),
            "warnings": warnings,
            "message": message,
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "message": f"布局检查失败: {e}"}


# ── 布局自动重排 ─────────────────────────────────────────────────

_SAFE_MARGIN_EMU = 91440  # 0.10 英寸
_MIN_SHAPE_GAP_EMU = 137160  # 0.15 英寸


def _box_as_inches(box: tuple[int, int, int, int]) -> dict[str, float]:
    left, top, width, height = box
    return {
        "left": round(left / 914400, 3),
        "top": round(top / 914400, 3),
        "width": round(width / 914400, 3),
        "height": round(height / 914400, 3),
    }


def _boxes_overlap(
    first: tuple[int, int, int, int],
    second: tuple[int, int, int, int],
) -> bool:
    left_a, top_a, width_a, height_a = first
    left_b, top_b, width_b, height_b = second
    overlap_x = min(left_a + width_a, left_b + width_b) - max(left_a, left_b)
    overlap_y = min(top_a + height_a, top_b + height_b) - max(top_a, top_b)
    return overlap_x > _LAYOUT_EPSILON_EMU and overlap_y > _LAYOUT_EPSILON_EMU


def _is_out_of_bounds(box: tuple[int, int, int, int], slide_width: int, slide_height: int) -> bool:
    left, top, width, height = box
    return (
        left < -_LAYOUT_EPSILON_EMU
        or top < -_LAYOUT_EPSILON_EMU
        or left + width > slide_width + _LAYOUT_EPSILON_EMU
        or top + height > slide_height + _LAYOUT_EPSILON_EMU
    )


def _clamp_box(
    box: tuple[int, int, int, int],
    slide_width: int,
    slide_height: int,
) -> tuple[int, int, int, int]:
    left, top, width, height = box
    left = max(_SAFE_MARGIN_EMU, min(left, slide_width - _SAFE_MARGIN_EMU - width))
    top = max(_SAFE_MARGIN_EMU, min(top, slide_height - _SAFE_MARGIN_EMU - height))
    return left, top, width, height


def _box_inside(
    box: tuple[int, int, int, int],
    slide_width: int,
    slide_height: int,
) -> bool:
    left, top, width, height = box
    return (
        left >= _SAFE_MARGIN_EMU - _LAYOUT_EPSILON_EMU
        and top >= _SAFE_MARGIN_EMU - _LAYOUT_EPSILON_EMU
        and left + width <= slide_width - _SAFE_MARGIN_EMU + _LAYOUT_EPSILON_EMU
        and top + height <= slide_height - _SAFE_MARGIN_EMU + _LAYOUT_EPSILON_EMU
    )


def _fit_box(
    box: tuple[int, int, int, int],
    slide_width: int,
    slide_height: int,
) -> tuple[int, int, int, int]:
    """Scale a too-large shape proportionally, then move it inside the margin."""
    left, top, width, height = box
    available_width = max(1, slide_width - 2 * _SAFE_MARGIN_EMU)
    available_height = max(1, slide_height - 2 * _SAFE_MARGIN_EMU)
    scale = min(1.0, available_width / width, available_height / height)
    if scale < 1.0:
        width = max(1, int(width * scale))
        height = max(1, int(height * scale))
    return _clamp_box((left, top, width, height), slide_width, slide_height)


def _candidate_boxes(
    box: tuple[int, int, int, int],
    placed: list[tuple[int, int, int, int]],
) -> list[tuple[int, int, int, int]]:
    """Generate move/shrink candidates around already placed shapes."""
    left, top, width, height = box
    candidates = [box]
    overlapping = [other for other in placed if _boxes_overlap(box, other)]
    if overlapping:
        down = max(other[1] + other[3] for other in overlapping) + _MIN_SHAPE_GAP_EMU
        right = max(other[0] + other[2] for other in overlapping) + _MIN_SHAPE_GAP_EMU
        candidates.extend([
            (left, down, width, height),
            (right, top, width, height),
            (right, down, width, height),
        ])
    for other in placed:
        candidates.append((left, other[1] + other[3] + _MIN_SHAPE_GAP_EMU, width, height))
        candidates.append((other[0] + other[2] + _MIN_SHAPE_GAP_EMU, top, width, height))
    for factor in (0.9, 0.8, 0.7, 0.6, 0.5):
        scaled_width = max(1, int(width * factor))
        scaled_height = max(1, int(height * factor))
        candidates.append((left, top, scaled_width, scaled_height))
        for other in placed:
            candidates.append((
                left, other[1] + other[3] + _MIN_SHAPE_GAP_EMU,
                scaled_width, scaled_height,
            ))
            candidates.append((
                other[0] + other[2] + _MIN_SHAPE_GAP_EMU, top,
                scaled_width, scaled_height,
            ))
    return candidates


def _find_free_box(
    box: tuple[int, int, int, int],
    placed: list[tuple[int, int, int, int]],
    slide_width: int,
    slide_height: int,
) -> tuple[int, int, int, int] | None:
    """Choose the least-disruptive candidate that is inside the page and free."""
    best: tuple[int, int, int, int] | None = None
    best_score: float | None = None
    original_area = box[2] * box[3]
    for candidate in _candidate_boxes(box, placed):
        candidate = _clamp_box(candidate, slide_width, slide_height)
        if candidate == box:
            continue
        if not _box_inside(candidate, slide_width, slide_height):
            continue
        if any(_boxes_overlap(candidate, other) for other in placed):
            continue
        distance = abs(candidate[0] - box[0]) + abs(candidate[1] - box[1])
        area_loss = 1.0 - (candidate[2] * candidate[3]) / original_area
        score = distance + area_loss * 914400
        if best_score is None or score < best_score:
            best = candidate
            best_score = score
    return best


def reflow_slide_layout(file_path: str) -> dict:
    """重排越界/重叠形状并写回文件。

    - 越界：保留 0.1in 页边距，优先平移；形状大于可用区域时按比例缩小；
    - 重叠：按阅读顺序（上→下、左→右）向下/向右移动，最小间距 0.15in，
      仍放不下时尝试缩小；
    - 无法安全放置时保留原状并返回 warning；
    - 只修改 left/top/width/height，不改变字体、颜色或文本内容。
    """
    from .utils import atomic_save, validate_input_file

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return {"success": False, "file_path": file_path, "message": str(e)}

    try:
        prs = Presentation(file_path)
        slide_width = int(prs.slide_width)
        slide_height = int(prs.slide_height)
        reflow: list[dict] = []
        warnings: list[dict] = []
        shape_count = 0
        slides_checked = 0
        slides_reflowed = 0

        for slide_index, slide in enumerate(prs.slides, start=1):
            slides_checked += 1
            entries: list[tuple[object, int, tuple[int, int, int, int]]] = []
            for shape_index, shape in enumerate(slide.shapes, start=1):
                box = _shape_box(shape)
                if box is None:
                    continue
                shape_count += 1
                entries.append((shape, shape_index, box))
            if not entries:
                continue

            fitted: dict[int, tuple[int, int, int, int]] = {}
            reasons: dict[int, str] = {}
            for shape, _index, box in entries:
                if _is_out_of_bounds(box, slide_width, slide_height):
                    fitted[id(shape)] = _fit_box(box, slide_width, slide_height)
                    reasons[id(shape)] = "out_of_bounds"
                else:
                    fitted[id(shape)] = box

            order = sorted(
                entries,
                key=lambda item: (
                    fitted[id(item[0])][1],
                    fitted[id(item[0])][0],
                    int(getattr(item[0], "shape_id", 0) or 0),
                ),
            )
            placed: list[tuple[int, int, int, int]] = []
            final: dict[int, tuple[int, int, int, int]] = {}
            for shape, shape_index, box in order:
                current = fitted[id(shape)]
                if any(_boxes_overlap(current, other) for other in placed):
                    reason = (
                        "out_of_bounds+overlap"
                        if reasons.get(id(shape)) == "out_of_bounds"
                        else "overlap"
                    )
                    candidate = _find_free_box(current, placed, slide_width, slide_height)
                    if candidate is None:
                        warnings.append({
                            "type": "unplaceable",
                            "slide": slide_index,
                            "shape": _shape_label(shape, shape_index),
                            "shape_id": int(getattr(shape, "shape_id", 0) or 0),
                            "detail": "无法在保留 0.1in 边距和 0.15in 间距的前提下放置，保持原状",
                        })
                        reasons[id(shape)] = "unplaceable"
                        final[id(shape)] = box
                    else:
                        reasons[id(shape)] = reason
                        final[id(shape)] = candidate
                else:
                    final[id(shape)] = current
                placed.append(final[id(shape)])

            changed = 0
            for shape, shape_index, box in entries:
                new_box = final[id(shape)]
                if new_box == box:
                    if reasons.get(id(shape)) == "unplaceable":
                        reflow.append({
                            "shape_id": int(getattr(shape, "shape_id", 0) or 0),
                            "name": str(getattr(shape, "name", "") or ""),
                            "slide": slide_index,
                            "action": "unchanged",
                            "reason": "unplaceable",
                            "from": _box_as_inches(box),
                            "to": _box_as_inches(box),
                        })
                    continue
                shape.left = int(new_box[0])
                shape.top = int(new_box[1])
                shape.width = int(new_box[2])
                shape.height = int(new_box[3])
                moved = new_box[:2] != box[:2]
                resized = new_box[2:] != box[2:]
                action = "move_resize" if moved and resized else ("resize" if resized else "move")
                reflow.append({
                    "shape_id": int(getattr(shape, "shape_id", 0) or 0),
                    "name": str(getattr(shape, "name", "") or ""),
                    "slide": slide_index,
                    "action": action,
                    "reason": reasons.get(id(shape), "overlap"),
                    "from": _box_as_inches(box),
                    "to": _box_as_inches(new_box),
                })
                changed += 1
            if changed:
                slides_reflowed += 1

        changed_count = sum(1 for item in reflow if item["action"] != "unchanged")
        if changed_count:
            atomic_save(prs.save, file_path)
        audit = audit_slide_layout(file_path)
        remaining = audit.get("warnings", []) if audit.get("success") else []

        if changed_count:
            message = (
                f"已重排 {changed_count} 个形状（{slides_reflowed} 张幻灯片）"
            )
        else:
            message = "未发现可安全重排的越界或重叠形状"
        if remaining:
            message += f"；仍有 {len(remaining)} 条布局警告"
        return {
            "success": True,
            "file_path": file_path,
            "slides_checked": slides_checked,
            "slides_reflowed": slides_reflowed,
            "shape_count": shape_count,
            "reflow_count": changed_count,
            "reflow": reflow,
            "warnings": warnings,
            "remaining_warnings": remaining,
            "remaining_warning_count": len(remaining),
            "message": message,
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "message": f"布局重排失败: {e}"}


# ── 占位符清理 ───────────────────────────────────────────────────

def clean_unused_placeholders(
    file_path: str,
) -> dict:
    """清理 PPT 文件中未使用的空白占位符（如空标题框、空文本框）。

    移除每张幻灯片上没有文字内容的占位符形状，使幻灯片更加整洁。
    """
    from .utils import atomic_save, validate_input_file

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as e:
        return {"success": False, "file_path": file_path, "message": str(e)}

    try:
        prs = Presentation(file_path)
        total_removed = 0

        for slide in prs.slides:
            shapes_to_remove = []
            for shape in slide.shapes:
                if shape.is_placeholder:
                    if not shape.has_text_frame or not shape.text_frame.text.strip():
                        shapes_to_remove.append(shape)

            for shape in shapes_to_remove:
                sp = shape._element
                sp.getparent().remove(sp)
                total_removed += 1

        atomic_save(prs.save, file_path)
        return {
            "success": True,
            "file_path": file_path,
            "placeholders_removed": total_removed,
            "message": f"已清理 {total_removed} 个未使用的占位符",
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "message": f"占位符清理失败: {e}"}


# ── 模板分析 ──────────────────────────────────────────────────────

def analyze_powerpoint_template(
    file_path: str,
) -> dict:
    """分析已有 PPT 模板，提取布局、配色和字体信息。

    返回模板的幻灯片尺寸、可用布局名称、使用到的字体列表，
    以及可提取的配色方案建议。
    """
    from .utils import validate_input_file

    try:
        validate_input_file(file_path, (".pptx", ".potx"))
    except ValueError as e:
        return {"success": False, "file_path": file_path, "message": str(e)}

    try:
        prs = _load_presentation(file_path)
        layout_names = [layout.name for layout in prs.slide_layouts]
        fonts_found: set[str] = set()
        color_hexes: set[str] = set()

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.font.name:
                                fonts_found.add(run.font.name)
                            try:
                                color_hexes.add(str(run.font.color.rgb))
                            except Exception:
                                pass

        return {
            "success": True,
            "file_path": file_path,
            "slide_count": len(prs.slides),
            "slide_width": round(prs.slide_width / 914400, 2),  # EMU → 英寸
            "slide_height": round(prs.slide_height / 914400, 2),
            "layout_names": layout_names,
            "fonts_used": sorted(fonts_found),
            "color_hexes_found": sorted(color_hexes),
            "message": f"模板分析完成: {len(layout_names)} 种布局, {len(fonts_found)} 种字体, {len(color_hexes)} 种颜色",
        }
    except Exception as e:
        return {"success": False, "file_path": file_path, "message": f"模板分析失败: {e}"}


# ── 框架图生成 ────────────────────────────────────────────────────

def create_framework_diagram(
    output_path: str,
    diagram: FrameworkDiagram,
    if_exists: str = "overwrite",
) -> CreateResult:
    """自动生成 PPT 框架图（阶段标签 + 节点 + 箭头连接）。

    根据 FrameworkDiagram 定义自动布局并渲染为单个幻灯片：
    - 左侧圆柱形阶段标签
    - 阶段背景矩形容器
    - 节点按指定形状（rect/rounded_rect/ellipse/diamond）渲染
    - 相邻节点自动生成箭头，也可通过 edges 指定显式连接
    - 支持 12 套配色方案
    """
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR_TYPE

    from .models import CreateResult
    from .utils import atomic_save, ensure_output_dir, resolve_if_exists, validate_output_path

    try:
        validate_output_path(output_path, (".pptx",))
        ensure_output_dir(output_path)
        resolved_path = resolve_if_exists(output_path, if_exists)
    except (OSError, ValueError) as e:
        return CreateResult(success=False, file_path=output_path, message=str(e))

    try:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide_layout = prs.slide_layouts[6]  # blank
        slide = prs.slides.add_slide(slide_layout)

        scheme = _PRESET_COLOR_SCHEMES.get(diagram.color_scheme, _PRESET_COLOR_SCHEMES["professional-blue"])
        scheme_bg = scheme["background_color"]
        scheme_title = scheme["title_color"]
        scheme_accent = scheme["accent_color"]
        scheme_body = scheme["body_color"]
        # 连接线用中性灰弱化（PptMcp 设计规范：箭头应比内容弱，中性灰优于主色）
        arrow_color = diagram.arrow_color or "94A3B8"

        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = _hex_to_rgb(scheme_bg)

        # ── 标题 ──────────────────────────────────────────────
        title_font = diagram.font_cn
        body_font_cn = diagram.font_cn
        body_font_en = diagram.font_en
        title_fz = diagram.title_font_size

        accent_bar_y = 0.62
        if diagram.title:
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.15), Inches(11.5), Inches(0.45))
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            _set_mixed_font_text(p, diagram.title, title_font, body_font_en, title_fz, True,
                               _hex_to_rgb(scheme_title))
            p.alignment = PP_ALIGN.CENTER
            title_y_offset = Inches(0.7)
        else:
            title_y_offset = Inches(0.2)

        if diagram.subtitle:
            sub_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.5), Inches(0.3))
            sub_tf = sub_title.text_frame
            p = sub_tf.paragraphs[0]
            p.text = diagram.subtitle
            p.alignment = PP_ALIGN.CENTER
            run = p.runs[0]
            run.font.size = Pt(12)
            run.font.color.rgb = _hex_to_rgb(scheme_body)
            run.font.name = scheme["body_font"]
            title_y_offset = Inches(0.95)
            accent_bar_y = 0.9

        # 标题下 accent 短横条（居中，60% 宽）
        if diagram.title:
            bar_w = 11.5 * 0.6
            bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.8 + (11.5 - bar_w) / 2), Inches(accent_bar_y),
                Inches(bar_w), Inches(0.04),
            )
            bar.fill.solid()
            bar.fill.fore_color.rgb = _hex_to_rgb(scheme_accent)
            bar.line.fill.background()

        # ── 构建节点索引 ───────────────────────────────────────
        node_id_to_pos: dict[str, dict] = {}  # id → (shape, x, y, w, h)
        stage_count = len(diagram.stages)
        if stage_count == 0:
            atomic_save(prs.save, resolved_path)
            return CreateResult(success=True, file_path=resolved_path, file_size=0, slides_count=1, message="空框架图已创建")

        # 可用区域
        slide_width_in = 13.333
        content_left = 1.15  # 阶段标签左边
        content_right = slide_width_in - 0.3
        content_width = content_right - content_left

        # 阶段之间 gap
        usable_height = 7.5 - title_y_offset.inches - 0.3
        stage_unit_h = usable_height / stage_count - diagram.stage_gap

        shape_map = {
            # 基础形状
            "rect": MSO_SHAPE.RECTANGLE,
            "rounded_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
            "ellipse": MSO_SHAPE.OVAL,
            "diamond": MSO_SHAPE.DIAMOND,
            # Flowchart 流程图标准形状
            "process": MSO_SHAPE.FLOWCHART_PROCESS,
            "decision": MSO_SHAPE.FLOWCHART_DECISION,
            "data": MSO_SHAPE.FLOWCHART_DATA,
            "document": MSO_SHAPE.FLOWCHART_DOCUMENT,
            "terminator": MSO_SHAPE.FLOWCHART_TERMINATOR,
            "predefined": MSO_SHAPE.FLOWCHART_PREDEFINED_PROCESS,
            # Basic Shapes
            "cylinder": MSO_SHAPE.CAN,
            "parallelogram": MSO_SHAPE.PARALLELOGRAM,
            # Block Arrows
            "chevron": MSO_SHAPE.CHEVRON,
            "right_arrow": MSO_SHAPE.RIGHT_ARROW,
            "left_arrow": MSO_SHAPE.LEFT_ARROW,
            # Callouts
            "callout": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
            "cloud": MSO_SHAPE.CLOUD,
        }

        for si, stage in enumerate(diagram.stages):
            stage_y = title_y_offset.inches + si * (stage_unit_h + diagram.stage_gap)
            stage_color = stage.color or scheme_accent

            # ── 阶段背景容器（极淡同色系底 + 同色系细边 + 柔和阴影） ──
            bg = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(content_left + 0.55), Inches(stage_y),
                Inches(content_width - 0.55), Inches(stage_unit_h),
            )
            bg.fill.solid()
            bg.fill.fore_color.rgb = _hex_to_rgb(_mix_with_white(stage_color, 0.96))
            bg.line.color.rgb = _hex_to_rgb(_mix_with_white(stage_color, 0.88))
            bg.line.width = Pt(0.75)
            _add_shape_shadow(bg, blur_pt=10, dist_pt=2, opacity_pct=25)

            # ── 阶段标签（实心主色胶囊 + 白字） ────────────────
            pill_w, pill_h = 1.25, 0.38
            pill = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(0.25), Inches(stage_y + stage_unit_h / 2 - pill_h / 2),
                Inches(pill_w), Inches(pill_h),
            )
            pill.adjustments[0] = 0.5  # 胶囊形
            pill.fill.solid()
            pill.fill.fore_color.rgb = _hex_to_rgb(stage_color)
            pill.line.color.rgb = _hex_to_rgb(_mix_with_black(stage_color, 0.2))
            pill.line.width = Pt(0.75)
            _add_shape_shadow(pill)
            pill_tf = pill.text_frame
            pill_tf.word_wrap = False
            pill_tf.margin_left = pill_tf.margin_right = Inches(0.02)
            pill_p = pill_tf.paragraphs[0]
            pill_p.alignment = PP_ALIGN.CENTER
            _set_mixed_font_text(pill_p, stage.label, body_font_cn, body_font_en,
                                 12, True, _hex_to_rgb("FFFFFF"))

            nodes = stage.nodes
            if not nodes:
                continue

            # ── 节点水平排列 ──────────────────────────────────
            node_count = len(nodes)
            node_w = diagram.node_width
            node_h = diagram.node_height
            total_nodes_w = node_count * node_w + (node_count - 1) * 0.3
            start_x = content_left + 0.6 + max(0, (content_width - 0.6 - total_nodes_w) / 2)
            node_cy = stage_y + stage_unit_h / 2 - node_h / 2

            prev_shape = None
            for ni, node in enumerate(nodes):
                nx = start_x + ni * (node_w + 0.3)
                n_color = node.fill_color or stage_color
                shape_type = shape_map.get(node.style, MSO_SHAPE.ROUNDED_RECTANGLE)

                shape = slide.shapes.add_shape(
                    shape_type,
                    Inches(nx), Inches(node_cy),
                    Inches(node_w), Inches(node_h),
                )
                # 节点：主色填充（不再提亮）+ 同色系深边框 + 白字
                shape.fill.solid()
                shape.fill.fore_color.rgb = _hex_to_rgb(n_color)
                shape.line.color.rgb = _hex_to_rgb(_mix_with_black(n_color, 0.25))
                shape.line.width = Pt(1.0)
                _add_shape_shadow(shape)

                tf = shape.text_frame
                tf.word_wrap = True
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # 放不下时自动缩小字号
                tf.margin_left = tf.margin_right = Inches(0.04)
                tf.margin_top = tf.margin_bottom = Inches(0.02)
                # 处理 \n 换行，每条行用中英文混排字体，深底统一白字
                # 兼容真实换行符（MCP JSON 转义）与字面反斜杠+n 两种写法
                lines = node.text.replace("\\n", "\n").split("\n")
                for li, line in enumerate(lines):
                    if not line.strip() and li != 0:
                        # 空行保留段落结构；首行空则跳过（无意义）
                        continue
                    if li == 0:
                        p = tf.paragraphs[0]
                    else:
                        p = tf.add_paragraph()
                    p.alignment = PP_ALIGN.CENTER
                    _set_mixed_font_text(
                        p, line, body_font_cn, body_font_en,
                        node.font_size or 14, node.bold,
                        _hex_to_rgb("FFFFFF"),
                    )

                # 步骤编号徽章（阶段内节点 ≥2 时）：白色小圆 + 主色 Consolas 数字
                if node_count >= 2 and node.style != "terminator":
                    badge_d = 0.24
                    badge = slide.shapes.add_shape(
                        MSO_SHAPE.OVAL,
                        Inches(nx - 0.08), Inches(node_cy - 0.08),
                        Inches(badge_d), Inches(badge_d),
                    )
                    badge.fill.solid()
                    badge.fill.fore_color.rgb = _hex_to_rgb("FFFFFF")
                    badge.line.color.rgb = _hex_to_rgb(n_color)
                    badge.line.width = Pt(1.0)
                    _add_shape_shadow(badge, blur_pt=4, dist_pt=1.5, opacity_pct=35)
                    btf = badge.text_frame
                    btf.margin_left = btf.margin_right = btf.margin_top = btf.margin_bottom = 0
                    bp = btf.paragraphs[0]
                    bp.alignment = PP_ALIGN.CENTER
                    br = bp.add_run()
                    br.text = str(ni + 1)
                    br.font.size = Pt(9)
                    br.font.bold = True
                    br.font.color.rgb = _hex_to_rgb(n_color)
                    br.font.name = "Consolas"

                node_id_to_pos[node.id] = {"shape": shape, "x": nx, "y": node_cy, "w": node_w, "h": node_h}

                # ── 节点间自动箭头（同阶段内，始终 STRAIGHT） ──
                if prev_shape is not None:
                    prev = node_id_to_pos[nodes[ni - 1].id]
                    _add_connector_arrow(
                        slide, arrow_color,
                        MSO_CONNECTOR_TYPE.STRAIGHT,
                        Inches(prev["x"] + prev["w"]), Inches(prev["y"] + prev["h"] / 2),
                        Inches(nx), Inches(node_cy + node_h / 2),
                    )

                prev_shape = shape

        # ── 显式 edges（跨阶段连接，支持 edge_style） ────
        connector_style = {
            "straight": MSO_CONNECTOR_TYPE.STRAIGHT,
            "elbow": MSO_CONNECTOR_TYPE.ELBOW,
            "curve": MSO_CONNECTOR_TYPE.CURVE,
        }
        for edge in diagram.edges:
            if edge.from_id not in node_id_to_pos or edge.to_id not in node_id_to_pos:
                continue
            f = node_id_to_pos[edge.from_id]
            t = node_id_to_pos[edge.to_id]
            fx = f["x"] + f["w"] / 2
            fy = f["y"] + f["h"]
            tx = t["x"] + t["w"] / 2
            ty = t["y"]
            ctype = connector_style.get(edge.edge_style, MSO_CONNECTOR_TYPE.STRAIGHT)
            _add_connector_arrow(slide, arrow_color, ctype,
                                 Inches(fx), Inches(fy), Inches(tx), Inches(ty))

            if edge.label:
                lx = (fx + tx) / 2 - 0.5
                ly = (fy + ty) / 2 - 0.15
                label_box = slide.shapes.add_textbox(Inches(lx), Inches(ly), Inches(1.0), Inches(0.3))
                ltf = label_box.text_frame
                p = ltf.paragraphs[0]
                p.text = edge.label
                p.alignment = PP_ALIGN.CENTER
                run = p.runs[0]
                run.font.size = Pt(9)
                run.font.italic = True
                run.font.color.rgb = _hex_to_rgb(arrow_color)

        atomic_save(prs.save, resolved_path)
        import os as _os
        size = _os.path.getsize(resolved_path) if _os.path.exists(resolved_path) else 0
        return CreateResult(
            success=True,
            file_path=resolved_path,
            file_size=size,
            slides_count=1,
            message=f"框架图生成成功: {len(diagram.stages)} 个阶段, {sum(len(s.nodes) for s in diagram.stages)} 个节点",
        )
    except Exception as e:
        return CreateResult(success=False, file_path=output_path, message=f"框架图生成失败: {e}")


def _set_mixed_font_text(paragraph, text: str, font_cn: str, font_en: str,
                         font_size: int, bold: bool = False,
                         color_rgb=None) -> None:
    """将文本按中英文拆分后写入段落，中文用 font_cn，西文用 font_en。

    使用正则 [⏎-￭] 块匹配中文字符，非中文段落使用西文字体。
    保留原始 run 用于首段，后续添加的 run 使用 add_run()。
    """
    import re
    from pptx.util import Pt

    # 清空段落默认文本
    paragraph.clear()

    # 按中文/非中文边界拆分
    # 中文字符 Unicode 范围：基本汉字 4E00-9FFF, 扩展A 3400-4DBF, 标点等
    pattern = re.compile(r'([\u4e00-\u9fff\u3400-\u4dbf\u3000-\u303f\uff00-\uffef]+)')
    segments = pattern.split(text)

    for seg in segments:
        if not seg:
            continue
        run = paragraph.add_run()
        run.text = seg
        run.font.size = Pt(font_size)
        run.font.bold = bold
        # 判断是否中文字段
        if pattern.match(seg):
            run.font.name = font_cn
        else:
            run.font.name = font_en
        if color_rgb is not None:
            run.font.color.rgb = color_rgb


def _add_connector_arrow(slide, arrow_color_hex: str, connector_type, x1, y1, x2, y2) -> None:
    """创建一条带箭头末端的连接线。"""
    from pptx.oxml.ns import qn
    from lxml import etree
    connector = slide.shapes.add_connector(connector_type, x1, y1, x2, y2)
    connector.line.color.rgb = _hex_to_rgb(arrow_color_hex)
    connector.line.width = Pt(1.5)
    # a:ln 位于 p:spPr 之下，直接 find(qn("a:ln")) 匹配不到，必须经 LineFormat 取
    tail_end = connector.line._get_or_add_ln()
    tail = tail_end.find(qn("a:tailEnd"))
    if tail is None:
        tail = etree.SubElement(tail_end, qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("len", "med")


def _lighten(hex_color: str, amount: int = 30) -> str:
    """将 HEX 颜色提亮指定量。"""
    h = hex_color.lstrip("#")
    r = min(255, int(h[0:2], 16) + amount)
    g = min(255, int(h[2:4], 16) + amount)
    b = min(255, int(h[4:6], 16) + amount)
    return f"{r:02X}{g:02X}{b:02X}"


def _mix_with_white(hex_color: str, weight: float = 0.9) -> str:
    """将 HEX 颜色向白色混合，weight=0.9 表示 90% 白 + 10% 原色（极淡同色系底）。"""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    w = max(0.0, min(1.0, weight))
    return (
        f"{round(r + (255 - r) * w):02X}"
        f"{round(g + (255 - g) * w):02X}"
        f"{round(b + (255 - b) * w):02X}"
    )


def _mix_with_black(hex_color: str, weight: float = 0.25) -> str:
    """将 HEX 颜色向黑色混合，weight=0.25 生成同色系深边框色。"""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    w = max(0.0, min(1.0, weight))
    return (
        f"{round(r * (1 - w)):02X}"
        f"{round(g * (1 - w)):02X}"
        f"{round(b * (1 - w)):02X}"
    )


def _add_shape_shadow(shape, blur_pt: float = 8.0, dist_pt: float = 3.0,
                      opacity_pct: int = 50) -> None:
    """给 python-pptx 形状加柔和外阴影（手写 OOXML a:effectLst/a:outerShdw）。

    python-pptx 无原生 shadow API（官方 scope 明确 out-of-scope），需在
    p:spPr 下手写 effectLst。参数对齐 PptxGenJS ShadowProps 的语义。
    """
    from lxml import etree
    from pptx.oxml.ns import qn

    sp_pr = shape._element.spPr
    # 移除可能已存在的 effectLst，避免重复阴影
    for old in sp_pr.findall(qn("a:effectLst")):
        sp_pr.remove(old)
    effect_lst = etree.SubElement(sp_pr, qn("a:effectLst"))
    outer = etree.SubElement(effect_lst, qn("a:outerShdw"))
    outer.set("blurRad", str(int(blur_pt * 12700)))
    outer.set("dist", str(int(dist_pt * 12700)))
    outer.set("dir", "5400000")  # 正下方向
    outer.set("rotWithShape", "0")
    srgb = etree.SubElement(outer, qn("a:srgbClr"))
    srgb.set("val", "000000")
    alpha = etree.SubElement(srgb, qn("a:alpha"))
    alpha.set("val", str(int(opacity_pct * 1000)))


# ── 单形状绘制（ppt_manage_shapes）──────────────────────────────

_PPT_SHAPE_PRESET_MAP = {
    "rect": MSO_SHAPE.RECTANGLE,
    "rounded_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "oval": MSO_SHAPE.OVAL,
    "arrow": MSO_SHAPE.RIGHT_ARROW,
    "cylinder": MSO_SHAPE.CAN,
    "diamond": MSO_SHAPE.DIAMOND,
}


def add_shape_to_presentation(file_path: str, slide_number: int, shape_def) -> dict[str, object]:
    """向指定幻灯片添加一个形状；文件已打开时优先走活动文档 COM。"""
    com_result = office_com.try_add_open_powerpoint_shape(
        file_path, slide_number, shape_def)
    if com_result is not None:
        return com_result
    return _add_shape_file(file_path, slide_number, shape_def)


def list_slide_shapes(file_path: str, slide_number: int) -> dict[str, object]:
    """列出指定幻灯片的形状；文件已打开时优先走活动文档 COM。"""
    com_result = office_com.try_list_open_powerpoint_shapes(file_path, slide_number)
    if com_result is not None:
        return com_result
    return _list_shapes_file(file_path, slide_number)


def _add_shape_file(file_path: str, slide_number: int, shape_def) -> dict[str, object]:
    path = Path(file_path)
    presentation = _load_presentation(str(path))
    slide = _slide_at(presentation, slide_number)
    shape_type = _PPT_SHAPE_PRESET_MAP.get(shape_def.shape_type, MSO_SHAPE.RECTANGLE)
    left = Inches(shape_def.x if shape_def.x is not None else 0.5)
    top = Inches(shape_def.y if shape_def.y is not None else 1.5)
    width = Inches(shape_def.width)
    height = Inches(shape_def.height)
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    _apply_shape_style(shape, shape_def)
    backup_path = atomic_save(lambda tmp: presentation.save(tmp), str(path))
    return {
        "success": True,
        "action": "add",
        "mode": "file",
        "file_path": str(path.resolve()),
        "slide": slide_number,
        "shape_index": len(slide.shapes),
        "shape_name": shape.name,
        "backup_path": backup_path,
        "message": f"已向第 {slide_number} 张幻灯片添加形状（文件级写入）",
    }


def _list_shapes_file(file_path: str, slide_number: int) -> dict[str, object]:
    path = Path(file_path)
    presentation = _load_presentation(str(path))
    slide = _slide_at(presentation, slide_number)
    items = []
    for index, shape in enumerate(slide.shapes, start=1):
        box = _shape_box(shape)
        preset = None
        try:
            sp_pr = shape._element.find(qn("p:spPr"))
            if sp_pr is not None:
                geom = sp_pr.find(qn("a:prstGeom"))
                if geom is not None:
                    preset = geom.get("prst")
        except Exception:
            pass
        text = ""
        try:
            if shape.has_text_frame:
                text = shape.text_frame.text
        except Exception:
            pass
        items.append({
            "index": index,
            "shape_id": shape.shape_id,
            "name": shape.name,
            "type": shape._element.tag.split("}")[-1],
            "preset": preset,
            "x": round(box[0] / 914400, 4) if box else None,
            "y": round(box[1] / 914400, 4) if box else None,
            "width": round(box[2] / 914400, 4) if box else None,
            "height": round(box[3] / 914400, 4) if box else None,
            "text": text,
        })
    return {
        "success": True,
        "action": "list",
        "mode": "file",
        "file_path": str(path.resolve()),
        "slide": slide_number,
        "count": len(items),
        "shapes": items,
    }


def _slide_at(presentation, slide_number: int):
    count = len(presentation.slides)
    if slide_number < 1 or slide_number > count:
        raise ValueError(f"slide={slide_number} 超出范围（共 {count} 张）")
    return presentation.slides[slide_number - 1]


def _apply_shape_style(shape, shape_def) -> None:
    if shape_def.no_fill:
        shape.fill.background()
    elif shape_def.bg_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(shape_def.bg_color.upper())
    if shape_def.line_color:
        shape.line.color.rgb = RGBColor.from_string(shape_def.line_color.upper())
    if shape_def.line_width is not None:
        shape.line.width = Pt(shape_def.line_width)
    if shape_def.line_dash:
        shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    if shape_def.text:
        text_frame = shape.text_frame
        text_frame.word_wrap = not shape_def.vertical_text
        text_frame.text = shape_def.text
        paragraph = text_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER
        if shape_def.vertical_text:
            try:
                text_frame._txBody.bodyPr.set("vert", "vert270")
            except Exception:
                pass
        for run in paragraph.runs:
            if shape_def.font_color:
                run.font.color.rgb = RGBColor.from_string(shape_def.font_color.upper())
            run.font.size = Pt(shape_def.font_size) if shape_def.font_size else Pt(12)
            if shape_def.bold:
                run.font.bold = True
    if shape_def.hyperlink_url:
        try:
            shape.click_action.hyperlink.address = shape_def.hyperlink_url
        except Exception:
            pass