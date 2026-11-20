#!/usr/bin/env python3
"""重新生成 templates/ppt/ 下的 6 套 16:9 PPT 模板。

每套主题产出两个文件：
- `<slug>.pptx`：带 8 页示例的成品模板，可直接打开改文字；
- `<slug>.potx`：0 页版式模板，供 create_powerpoint --template 套用。

约束：11 个标准版式的顺序与名称保持不变，create_powerpoint 按固定索引映射。
修改模板文案/配色后运行本脚本，再用 convert_to_pdf + render_document_pages 复核。

用法:
  D:/anaconda3/python.exe -X utf8 tools/build_ppt_templates.py
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.util import Inches, Pt

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# ── 主题定义 ──────────────────────────────────────────────────────

THEMES = [
    {
        "slug": "business-report",
        "name": "商务汇报",
        "colors": {
            "primary": "1F3864", "secondary": "2E75B6", "accent": "ED7D31",
            "light": "F2F2F2", "muted": "767676", "text": "262626",
            "header_text": "FFFFFF", "stripe": "F2F5F9",
        },
        "fonts": {"title": "微软雅黑", "body": "微软雅黑"},
        "decor": "bar",
        "demo": {
            "cover": ("2026 年度经营分析汇报", "财务管理中心  ·  2026 年 11 月"),
            "toc": ("目录", ["01  经营概况", "02  核心指标", "03  问题与对策", "04  下一步计划"]),
            "section": ("01  经营概况", "全年营收与利润保持双增长"),
            "content": ("经营概况：规模与质量同步提升", [
                "全年营业收入 12.8 亿元，同比增长 18.4%",
                "毛利率 34.2%，较上年提升 2.1 个百分点",
                "经营性现金流净额 1.9 亿元，连续三年为正",
                "客户续约率 92%，核心客户流失率降至 3% 以内",
            ]),
            "chart": ("核心指标：季度营收稳步上行", ["Q1", "Q2", "Q3", "Q4"],
                      [("2025", (2.4, 2.7, 2.9, 3.2)), ("2026", (2.8, 3.1, 3.3, 3.6))]),
            "table": ("区域经营对比", ["区域", "营收（亿元）", "同比", "毛利率"],
                      [["华东", "4.6", "+26.1%", "36.8%"],
                       ["华南", "3.1", "+8.4%", "31.2%"],
                       ["华北", "2.9", "+12.7%", "33.5%"],
                       ["西部", "2.2", "+19.3%", "34.1%"]]),
            "two_col": ("优势与短板", "优势",
                        ["产品结构持续优化", "华东区域增长 26%", "交付周期缩短 15%"],
                        "待改进", ["华南区域低于预期", "应收账款周转放缓", "新客户获客成本上升"]),
            "end": ("谢谢聆听", "财务管理中心  ·  2026 年 11 月"),
        },
    },
    {
        "slug": "consulting",
        "name": "咨询提案",
        "colors": {
            "primary": "1A1A1A", "secondary": "00529B", "accent": "F2A900",
            "light": "F5F7FA", "muted": "6B7280", "text": "1F2937",
            "header_text": "FFFFFF", "stripe": "F0F4FA",
        },
        "fonts": {"title": "微软雅黑", "body": "微软雅黑"},
        "decor": "card",
        "demo": {
            "cover": ("智能制造转型：三年行动方案", "战略咨询项目组  ·  2026 年 11 月"),
            "toc": ("目录", ["01  现状诊断", "02  目标蓝图", "03  实施路径", "04  投入与收益"]),
            "section": ("01  现状诊断", "产线自动化率低于行业均值 18 个百分点"),
            "content": ("核心结论：转型窗口期只剩 24 个月", [
                "行业头部企业自动化率已达 72%，本公司为 54%",
                "人工成本占比 23%，高于行业均值 9 个百分点",
                "订单交付周期 21 天，行业标杆为 12 天",
                "若不转型，预计 2028 年毛利率下降 4.5 个百分点",
            ]),
            "chart": ("目标：三年内自动化率提升至 80%", ["2026", "2027", "2028", "2029"],
                      [("自动化率", (54, 63, 72, 80)), ("行业均值", (60, 66, 71, 75))]),
            "table": ("实施路径与里程碑", ["阶段", "重点举措", "周期", "责任部门"],
                      [["一期", "产线数字化改造", "6 个月", "制造部"],
                       ["二期", "MES 系统上线", "9 个月", "信息部"],
                       ["三期", "智能排产与预测", "12 个月", "计划部"]]),
            "two_col": ("投入与收益", "投入",
                        ["设备改造 2.4 亿元", "系统建设 0.6 亿元", "人员培训 0.2 亿元"],
                        "收益", ["年人工成本下降 0.8 亿元", "交付周期缩短至 13 天", "三年 ROI 约 118%"]),
            "end": ("谢谢", "战略咨询项目组"),
        },
    },
    {
        "slug": "academic-defense",
        "name": "学术答辩",
        "colors": {
            "primary": "003366", "secondary": "1F7A8C", "accent": "D97706",
            "light": "F4F7FB", "muted": "6B7280", "text": "1F2937",
            "header_text": "FFFFFF", "stripe": "EEF3F9",
        },
        "fonts": {"title": "华文中宋", "body": "微软雅黑"},
        "decor": "rule",
        "demo": {
            "cover": ("基于深度学习的遥感影像变化检测方法研究",
                      "答辩人：张  三　　导师：李四  教授　　2026 年 11 月"),
            "toc": ("汇报提纲", ["01  研究背景与意义", "02  国内外研究现状",
                                "03  研究方法与设计", "04  实验结果与分析", "05  结论与展望"]),
            "section": ("01  研究背景与意义", "变化检测是遥感应用的核心环节"),
            "content": ("研究问题：复杂场景下的误检与漏检", [
                "传统方法依赖人工特征，跨时相泛化能力不足",
                "现有深度模型在小目标与边界区域误检率偏高",
                "多源影像辐射差异导致伪变化干扰",
                "标注样本稀缺，模型训练容易过拟合",
            ]),
            "chart": ("实验结果：所提方法精度领先", ["OA", "F1", "IoU", "Kappa"],
                      [("基线模型", (0.842, 0.801, 0.734, 0.792)),
                       ("本文方法", (0.913, 0.887, 0.821, 0.868))]),
            "table": ("与主流方法对比", ["方法", "OA", "F1", "参数量"],
                      [["FC-EF", "0.842", "0.801", "1.35 M"],
                       ["Bi-SRNet", "0.876", "0.849", "5.2 M"],
                       ["本文方法", "0.913", "0.887", "3.8 M"]]),
            "two_col": ("研究贡献与不足", "主要贡献",
                        ["提出跨时相特征对齐模块", "构建边界感知损失函数", "在两个公开数据集上取得最优结果"],
                        "研究不足", ["极端天气样本覆盖不足", "模型推理速度有待优化", "尚未验证多光谱场景"]),
            "end": ("敬请各位老师批评指正", "2026 年 11 月"),
        },
    },
    {
        "slug": "academic-seminar",
        "name": "开题与组会",
        "colors": {
            "primary": "1B4332", "secondary": "40916C", "accent": "B7791F",
            "light": "F4F7F4", "muted": "6B7280", "text": "1F2937",
            "header_text": "FFFFFF", "stripe": "EEF5EF",
        },
        "fonts": {"title": "华文中宋", "body": "微软雅黑"},
        "decor": "rule",
        "demo": {
            "cover": ("课题开题报告", "汇报人：张  三　　指导教师：李四  教授　　2026 年 11 月"),
            "toc": ("汇报提纲", ["01  研究背景", "02  研究问题", "03  研究设计",
                                "04  进度安排", "05  预期成果"]),
            "section": ("01  研究背景", "面向真实场景的模型鲁棒性问题"),
            "content": ("研究问题与价值", [
                "现有方法在分布偏移场景下性能下降明显",
                "缺乏统一的多源数据融合框架",
                "研究成果可支撑行业标准制定",
                "预期形成 2 项发明专利与 1 套开源工具",
            ]),
            "chart": ("前期实验结果", ["基线", "改进 A", "改进 B", "本文方案"],
                      [("准确率", (0.78, 0.83, 0.85, 0.89)),
                       ("召回率", (0.74, 0.80, 0.83, 0.87))]),
            "table": ("研究进度安排", ["阶段", "时间", "主要任务", "交付物"],
                      [["第一阶段", "2026.11–2027.03", "文献调研与方案设计", "综述与设计文档"],
                       ["第二阶段", "2027.04–2027.10", "算法实现与实验", "原型系统"],
                       ["第三阶段", "2027.11–2028.04", "论文撰写与结题", "学位论文"]]),
            "two_col": ("可行性分析", "已有基础",
                        ["完成预实验并验证核心假设", "具备数据集与算力条件", "已有 1 篇相关论文发表"],
                        "潜在风险", ["数据标注周期较长", "跨学科协作成本高", "实验进度可能受设备影响"]),
            "end": ("请各位老师指导", "2026 年 11 月"),
        },
    },
    {
        "slug": "government-report",
        "name": "政务汇报",
        "colors": {
            "primary": "A61B29", "secondary": "8C1C13", "accent": "C9A063",
            "light": "FBF7F0", "muted": "6B5B4B", "text": "2B2B2B",
            "header_text": "FFFFFF", "stripe": "F8F1E7",
        },
        "fonts": {"title": "华文中宋", "body": "仿宋"},
        "decor": "bar",
        "demo": {
            "cover": ("2026 年度重点工作情况汇报", "XX 市 XX 局  ·  2026 年 11 月"),
            "toc": ("汇报提纲", ["01  工作概况", "02  主要做法", "03  工作成效", "04  下一步打算"]),
            "section": ("01  工作概况", "各项重点任务总体进展顺利"),
            "content": ("工作概况：四项重点任务全面推进", [
                "全年召开专题会议 26 次，形成制度文件 8 项",
                "完成重点项目 15 个，投资完成率 108%",
                "开展专项督查 12 轮，问题整改率 100%",
                "群众满意度测评 96.4 分，同比提升 3.2 分",
            ]),
            "chart": ("主要指标完成情况", ["一季度", "二季度", "三季度", "四季度"],
                      [("计划进度", (25, 50, 75, 100)), ("实际进度", (27, 54, 79, 108))]),
            "table": ("重点任务完成情况", ["任务名称", "年度目标", "完成情况", "完成率"],
                      [["项目建设", "12 个", "13 个", "108%"],
                       ["制度完善", "6 项", "8 项", "133%"],
                       ["问题整改", "100%", "100%", "100%"]]),
            "two_col": ("做法与成效", "主要做法",
                        ["建立周调度、月通报机制", "实行清单化管理、销号式推进", "强化部门协同与督导问责"],
                        "工作成效", ["重点指标超额完成", "群众满意度稳步提升", "干部作风持续改进"]),
            "end": ("汇报完毕，请予指导", "XX 市 XX 局  ·  2026 年 11 月"),
        },
    },
    {
        "slug": "government-briefing",
        "name": "政策宣讲",
        "colors": {
            "primary": "C00000", "secondary": "8B0000", "accent": "B8860B",
            "light": "FAF7F2", "muted": "6B5B4B", "text": "2B2B2B",
            "header_text": "FFFFFF", "stripe": "F9F2EC",
        },
        "fonts": {"title": "华文中宋", "body": "仿宋"},
        "decor": "card",
        "demo": {
            "cover": ("《XX 管理办法》政策解读", "XX 市 XX 局  ·  2026 年 11 月"),
            "toc": ("解读提纲", ["01  出台背景", "02  核心内容", "03  主要变化", "04  落实要求"]),
            "section": ("01  出台背景", "回应基层关切，补齐制度短板"),
            "content": ("政策背景与总体考虑", [
                "原办法实施已满 5 年，部分条款与现行要求不衔接",
                "基层反映办事环节多、材料重复提交",
                "国家层面出台新规，需要地方配套落实",
                "坚持问题导向、便民导向和风险可控原则",
            ]),
            "chart": ("新旧办法办事环节对比", ["申请", "受理", "审核", "办结"],
                      [("原办法（环节数）", (4, 3, 5, 3)), ("新办法（环节数）", (2, 1, 3, 1))]),
            "table": ("主要变化一览", ["事项", "原规定", "新规定", "变化"],
                      [["办理时限", "20 个工作日", "10 个工作日", "压缩 50%"],
                       ["申请材料", "12 项", "8 项", "精简 33%"],
                       ["办理方式", "线下为主", "全程网办", "新增"]]),
            "two_col": ("落实要求", "各部门",
                        ["组织专题学习培训", "对照新规梳理办事指南", "及时更新系统与表单"],
                        "执行要点", ["不得增设额外条件", "严格执行时限要求", "做好政策宣传解读"]),
            "end": ("谢谢", "XX 市 XX 局  ·  2026 年 11 月"),
        },
    },
]

SLIDE_W, SLIDE_H = 13.333, 7.5

GEOMETRY = {
    0: {0: (0.9, 2.15, 8.6, 1.55), 1: (0.9, 3.95, 8.6, 0.75)},
    1: {0: (0.9, 0.5, 11.5, 0.85), 1: (0.9, 1.7, 11.5, 5.0)},
    2: {0: (0.9, 2.55, 11.5, 1.15), 1: (0.9, 3.9, 11.5, 0.8)},
    3: {0: (0.9, 0.5, 11.5, 0.85), 1: (0.9, 1.7, 5.5, 5.0), 2: (6.9, 1.7, 5.5, 5.0)},
    4: {
        0: (0.9, 0.5, 11.5, 0.85),
        1: (0.9, 1.55, 5.5, 0.55), 2: (0.9, 2.2, 5.5, 4.4),
        3: (6.9, 1.55, 5.5, 0.55), 4: (6.9, 2.2, 5.5, 4.4),
    },
    5: {0: (0.9, 0.5, 11.5, 0.85)},
    7: {0: (0.9, 0.5, 7.5, 0.85), 1: (0.9, 1.7, 7.5, 4.7), 2: (8.9, 1.7, 3.5, 4.7)},
    8: {0: (0.9, 0.5, 11.5, 0.85), 1: (1.6, 1.65, 10.1, 4.3), 2: (1.6, 6.1, 10.1, 0.7)},
}


# ── 基础写入工具 ──────────────────────────────────────────────────

def _rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def _style_run(run, size: int, color: str, bold: bool = False, font: str = "微软雅黑") -> None:
    run.font.size = Pt(size)
    run.font.color.rgb = _rgb(color)
    run.font.bold = bold
    run.font.name = font


def _bullet_none(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    if p_pr.find(f"{{{A_NS}}}buNone") is None:
        p_pr.append(parse_xml(f"<a:buNone {nsdecls('a')}/>"))


def _rect_xml(left: float, top: float, width: float, height: float,
              color: str, shape_id: int, alpha: int | None = None,
              line: str | None = None) -> str:
    fill = f'<a:solidFill><a:srgbClr val="{color}"/>'
    if alpha is not None:
        fill = (
            f'<a:solidFill><a:srgbClr val="{color}">'
            f'<a:alpha val="{alpha}"/></a:srgbClr>'
        )
    fill += "</a:solidFill>"
    outline = (
        f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
        if line else "<a:ln><a:noFill/></a:ln>"
    )
    return (
        f'<p:sp {nsdecls("p", "a")}>'
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="decor-{shape_id}"/>'
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        "<p:spPr>"
        f'<a:xfrm><a:off x="{int(left * 914400)}" y="{int(top * 914400)}"/>'
        f'<a:ext cx="{int(width * 914400)}" cy="{int(height * 914400)}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f"{fill}{outline}"
        "</p:spPr></p:sp>"
    )


def _ring_xml(left: float, top: float, size: float, color: str,
              shape_id: int, line_width: float = 1.6, alpha: int = 30000) -> str:
    return (
        f'<p:sp {nsdecls("p", "a")}>'
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="decor-ring-{shape_id}"/>'
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        "<p:spPr>"
        f'<a:xfrm><a:off x="{int(left * 914400)}" y="{int(top * 914400)}"/>'
        f'<a:ext cx="{int(size * 914400)}" cy="{int(size * 914400)}"/></a:xfrm>'
        '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
        "<a:noFill/>"
        f'<a:ln w="{int(line_width * 12700)}">'
        f'<a:solidFill><a:srgbClr val="{color}"><a:alpha val="{alpha}"/>'
        "</a:srgbClr></a:solidFill></a:ln>"
        "</p:spPr></p:sp>"
    )


def _set_theme(prs: Presentation, colors: dict[str, str], fonts: dict[str, str]) -> None:
    theme_part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)
    ns = {"a": A_NS}
    scheme = root.find(".//a:clrScheme", ns)
    mapping = {
        "dk1": "000000", "lt1": "FFFFFF", "dk2": colors["primary"],
        "lt2": colors["light"], "accent1": colors["secondary"],
        "accent2": colors["accent"], "accent3": colors["muted"],
        "accent4": colors["primary"], "accent5": colors["secondary"],
        "accent6": colors["accent"], "hlink": colors["secondary"],
        "folHlink": colors["muted"],
    }
    for slot, value in mapping.items():
        node = scheme.find(f"a:{slot}", ns)
        if node is None:
            continue
        for child in list(node):
            node.remove(child)
        etree.SubElement(node, f"{{{A_NS}}}srgbClr").set("val", value)
    for kind, key in (("majorFont", "title"), ("minorFont", "body")):
        font = root.find(f".//a:fontScheme/a:{kind}", ns)
        if font is None:
            continue
        for tag in ("latin", "ea"):
            node = font.find(f"a:{tag}", ns)
            if node is not None:
                node.set("typeface", fonts[key])
        hans = font.find("a:font[@script='Hans']", ns)
        if hans is not None:
            hans.set("typeface", fonts[key])
    theme_part._blob = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _lst_xml(level: int, align: str, size: int, color: str,
             bold: bool = False, bullet: bool = True) -> str:
    fill = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
    bold_attr = ' b="1"' if bold else ""
    bullet_xml = "" if bullet else "<a:buNone/>"
    return (
        f'<a:lvl{level}pPr {nsdecls("a")} algn="{align}">'
        f'{bullet_xml}<a:defRPr sz="{size}"{bold_attr}>{fill}</a:defRPr>'
        f"</a:lvl{level}pPr>"
    )


def _set_lst_style(shape, levels: list[tuple[int, str, int, str, bool, bool]]) -> None:
    """写入版式占位符的 lstStyle——新幻灯片继承的是它，而不是占位提示段落。

    元组顺序：level, align, size, color, bold, bullet。
    """
    tx_body = shape.text_frame._txBody
    lst = tx_body.find(f"{{{A_NS}}}lstStyle")
    if lst is None:
        lst = parse_xml(f'<a:lstStyle {nsdecls("a")}/>')
        tx_body.insert(1, lst)
    for child in list(lst):
        lst.remove(child)
    for level, align, size, color, bold, bullet in levels:
        lst.append(parse_xml(_lst_xml(level, align, size, color, bold, bullet)))


def _style_layouts(prs: Presentation, theme: dict) -> None:
    colors = theme["colors"]
    dark_pages = theme["decor"] == "card"
    body_sizes = [(1, 1800), (2, 1600), (3, 1500), (4, 1400)]
    for layout_idx, layout in enumerate(prs.slide_layouts):
        fill = layout.background.fill
        fill.solid()
        fill.fore_color.rgb = _rgb("FFFFFF")
        geometry = GEOMETRY.get(layout_idx, {})
        for ph in layout.placeholders:
            ph_idx = ph.placeholder_format.idx
            box = geometry.get(ph_idx)
            if box:
                ph.left, ph.top, ph.width, ph.height = (Inches(v) for v in box)
            align = "ctr" if layout_idx == 2 else "l"
            dark = dark_pages and layout_idx in (0, 2)
            if ph_idx == 0:
                ph.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                size = 4000 if layout_idx == 0 else 3600 if layout_idx == 2 else 3000
                _set_lst_style(ph, [(1, align, size,
                                     colors["header_text"] if dark else colors["primary"],
                                     True, False)])
            elif ph_idx in (1, 2, 3, 4):
                color = colors["header_text"] if dark else colors["text"]
                if layout_idx in (0, 2):
                    # 封面/章节页副标题：无项目符号
                    _set_lst_style(ph, [(1, align, 1700, color, False, False)])
                else:
                    _set_lst_style(ph, [(level, "l", size, color, False, True)
                                        for level, size in body_sizes])
            for p in ph.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER if layout_idx == 2 else PP_ALIGN.LEFT


def _decorate(prs: Presentation, theme: dict) -> None:
    colors = theme["colors"]
    layouts = prs.slide_layouts
    style = theme.get("decor", "bar")
    items: list[tuple[object, str]] = []

    def rect(layout, left, top, width, height, color, shape_id, **kw):
        items.append((layout, _rect_xml(left, top, width, height, color, shape_id, **kw)))

    def ring(layout, left, top, size, color, shape_id, **kw):
        items.append((layout, _ring_xml(left, top, size, color, shape_id, **kw)))

    if style == "bar":
        rect(layouts[0], 0, 0, 0.3, SLIDE_H, colors["primary"], 910)
        ring(layouts[0], 9.9, 0.95, 3.2, colors["secondary"], 911, line_width=2.4, alpha=24000)
        rect(layouts[0], 10.2, 4.55, 1.5, 0.06, colors["accent"], 913)
        rect(layouts[2], 0, 0, 0.3, SLIDE_H, colors["primary"], 914)
        ring(layouts[2], 8.5, 1.45, 4.5, colors["secondary"], 915, line_width=2.0, alpha=20000)
    elif style == "rule":
        rect(layouts[0], 0, 0, SLIDE_W, 0.28, colors["primary"], 920)
        rect(layouts[0], 0.9, 1.95, 3.2, 0.07, colors["accent"], 921)
        rect(layouts[2], 0, 0, SLIDE_W, 0.28, colors["primary"], 922)
        rect(layouts[2], 0.9, 3.82, 3.2, 0.07, colors["accent"], 923)
    else:  # card
        rect(layouts[0], 0, 0, SLIDE_W, SLIDE_H, colors["primary"], 930)
        rect(layouts[0], 0.6, 1.95, 0.09, 3.4, colors["accent"], 931)
        rect(layouts[2], 0, 0, SLIDE_W, SLIDE_H, colors["primary"], 933)

    # 内容页标题下的强调线 + 页脚分隔线 + 两栏分隔线
    for idx in (1, 3, 4, 5):
        rect(layouts[idx], 0.9, 1.45, 1.5, 0.06, colors["accent"], 940 + idx)
    for idx in (1, 3, 4, 5, 7, 8):
        rect(layouts[idx], 0.9, 7.05, 11.5, 0.01, "D9D9D9", 950 + idx)
    rect(layouts[3], 6.64, 1.8, 0.012, 4.85, "DCE3EC", 960)

    # 按添加顺序插入到占位符之前，保证背景在下、强调元素在上
    positions: dict[int, int] = {}
    for layout, xml in items:
        pos = positions.get(id(layout), 2)
        layout.shapes._spTree.insert(pos, parse_xml(xml))
        positions[id(layout)] = pos + 1


def _place_header(slide, text: str, theme: dict, align=PP_ALIGN.LEFT,
                  size: int = 30) -> None:
    colors = theme["colors"]
    for ph in slide.placeholders:
        if ph.placeholder_format.idx != 0:
            continue
        ph.text = text
        ph.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in ph.text_frame.paragraphs:
            p.alignment = align
            for run in p.runs:
                _style_run(run, size, colors["primary"], bold=True,
                           font=theme["fonts"]["title"])


def _fill_body(shape, blocks: list[tuple[str, bool]], theme: dict) -> None:
    """blocks: [(text, is_heading)]；heading 段不加项目符号。"""
    colors = theme["colors"]
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, (text, heading) in enumerate(blocks):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text
        p.level = 0
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(12 if not heading else 8)
        p.line_spacing = 1.15
        if heading:
            _bullet_none(p)
        for run in p.runs:
            _style_run(run, 18 if not heading else 19, colors["primary"] if heading
                       else colors["text"], bold=heading,
                       font=theme["fonts"]["body"])


def _add_basic_slide(prs: Presentation, layout_idx: int, theme: dict,
                     title: str, notes: str, size: int = 30):
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    _place_header(slide, title, theme, size=size)
    slide.notes_slide.notes_text_frame.text = notes
    return slide


# ── 演示页 ────────────────────────────────────────────────────────

def _add_demo(prs: Presentation, theme: dict) -> None:
    demo = theme["demo"]
    colors = theme["colors"]
    title_font = theme["fonts"]["title"]
    body_font = theme["fonts"]["body"]

    # 1 封面
    title, subtitle = demo["cover"]
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = title
            for p in ph.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                for run in p.runs:
                    _style_run(run, 38, colors["header_text"] if theme["decor"] == "card"
                               else colors["primary"], bold=True, font=title_font)
        elif idx == 1:
            ph.text = subtitle
            for p in ph.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                for run in p.runs:
                    _style_run(run, 17, colors["header_text"] if theme["decor"] == "card"
                               else colors["muted"], font=body_font)
    slide.notes_slide.notes_text_frame.text = (
        "封面页：替换主标题、副标题（单位/汇报人/日期）；深色封面主题的标题为白色。")

    # 2 目录
    toc_title, items = demo["toc"]
    slide = _add_basic_slide(prs, 1, theme, toc_title,
                             "目录页：4–6 个条目，编号与正文章节页一致。")
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 1:
            _fill_body(ph, [(item, False) for item in items], theme)

    # 3 章节页
    section_title, section_sub = demo["section"]
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = section_title
            for p in ph.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                for run in p.runs:
                    _style_run(run, 36, colors["header_text"] if theme["decor"] == "card"
                               else colors["primary"], bold=True, font=title_font)
        elif idx == 1:
            ph.text = section_sub
            for p in ph.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                for run in p.runs:
                    _style_run(run, 18, colors["header_text"] if theme["decor"] == "card"
                               else colors["muted"], font=body_font)
    slide.notes_slide.notes_text_frame.text = "章节页：编号 + 章节标题 + 一句结论式副标题。"

    # 4 内容页
    content_title, bullets = demo["content"]
    slide = _add_basic_slide(prs, 1, theme, content_title,
                             "内容页：3–5 条要点，每条一个结论；避免整段文字。")
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 1:
            _fill_body(ph, [(line, False) for line in bullets], theme)

    # 5 图表页
    chart_title, categories, series = demo["chart"]
    slide = _add_basic_slide(prs, 5, theme, chart_title,
                             "数据页：替换为实际数据；标题直接写结论。")
    data = CategoryChartData()
    data.categories = categories
    for name, values in series:
        data.add_series(name, values)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1.0), Inches(1.7),
        Inches(11.3), Inches(5.0), data,
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.font.size = Pt(12)
    chart.font.name = body_font

    # 6 表格页
    table_title, headers, rows = demo["table"]
    slide = _add_basic_slide(prs, 5, theme, table_title,
                             "表格页：表头深色底白字，数据行浅色斑马纹。")
    shape = slide.shapes.add_table(
        len(rows) + 1, len(headers), Inches(1.0), Inches(1.7),
        Inches(11.3), Inches(0.6 * (len(rows) + 1)),
    )
    table = shape.table
    for col, header in enumerate(headers):
        cell = table.cell(0, col)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(colors["primary"])
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.LEFT if col == 0 else PP_ALIGN.CENTER
            for run in p.runs:
                _style_run(run, 15, "FFFFFF", bold=True, font=body_font)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    for row_idx, row in enumerate(rows, start=1):
        for col, value in enumerate(row):
            cell = table.cell(row_idx, col)
            cell.text = value
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(
                colors["stripe"] if row_idx % 2 else "FFFFFF")
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER if col else PP_ALIGN.LEFT
                for run in p.runs:
                    _style_run(run, 14, colors["text"], font=body_font)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 7 两栏页
    two_title, left_title, left_items, right_title, right_items = demo["two_col"]
    slide = _add_basic_slide(prs, 3, theme, two_title,
                             "两栏页：左栏现状/优势，右栏问题/对策或补充说明。")
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 1:
            _fill_body(ph, [(left_title, True)] + [(item, False) for item in left_items], theme)
        elif idx == 2:
            _fill_body(ph, [(right_title, True)] + [(item, False) for item in right_items], theme)

    # 8 结尾
    end_title, end_sub = demo["end"]
    slide = prs.slides.add_slide(prs.slide_layouts[2])
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = end_title
            for p in ph.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    _style_run(run, 40, colors["header_text"] if theme["decor"] == "card"
                               else colors["primary"], bold=True, font=title_font)
        elif idx == 1:
            ph.text = end_sub
            for p in ph.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    _style_run(run, 17, colors["header_text"] if theme["decor"] == "card"
                               else colors["muted"], font=body_font)
    slide.notes_slide.notes_text_frame.text = "结尾页：致谢、联系方式或 Q&A。"


# ── 保存 ──────────────────────────────────────────────────────────

def _build(theme: dict, with_demo: bool) -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    _set_theme(prs, theme["colors"], theme["fonts"])
    _style_layouts(prs, theme)
    _decorate(prs, theme)
    if with_demo:
        _add_demo(prs, theme)
    return prs


def _to_potx(pptx_path: Path, potx_path: Path) -> None:
    shutil.copyfile(pptx_path, potx_path)
    tmp = potx_path.with_suffix(".zip")
    with zipfile.ZipFile(potx_path, "r") as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                    b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
                )
            dst.writestr(item, data)
    tmp.replace(potx_path)


def main() -> int:
    out_dir = Path(__file__).resolve().parent.parent / "templates" / "ppt"
    out_dir.mkdir(parents=True, exist_ok=True)
    for theme in THEMES:
        demo_prs = _build(theme, with_demo=True)
        demo_path = out_dir / f"{theme['slug']}.pptx"
        demo_prs.save(demo_path)
        layout_prs = _build(theme, with_demo=False)
        layout_path = out_dir / f"{theme['slug']}.pptx.tmp"
        layout_prs.save(layout_path)
        potx_path = out_dir / f"{theme['slug']}.potx"
        _to_potx(layout_path, potx_path)
        layout_path.unlink()
        print(f"{theme['slug']}: {demo_path.name} ({demo_path.stat().st_size} B, "
              f"{len(demo_prs.slides)} 页), {potx_path.name} ({potx_path.stat().st_size} B)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
