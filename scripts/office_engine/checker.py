# Vendored from mcp_office_writer/checker.py (mcp-office-writer 0.6.0, 2026-11).
# Upstream path: mcp_office_writer/checker.py (repository-relative).
# Modifications: added deterministic checks (installed fonts, heading numbering, PAGE/TOC fields, table column/merge structure, empty-paragraph ratio, mixed CJK/Latin fonts) with stable check ids; existing checks kept for backward compatibility.
"""
mcp-office-writer 后处理验证模块
参考 document-skills 的 postcheck.py，对生成的 Office 文档进行质量检查
"""
from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)


@dataclass
class CheckResult:
    """单项检查结果"""
    check_name: str
    passed: bool
    severity: str = "info"  # info / warn / error
    message: str = ""
    details: str = ""
    check_id: str = ""


@dataclass
class ValidationReport:
    """完整验证报告"""
    file_path: str
    checks: list[CheckResult] = field(default_factory=list)
    total: int = 0
    passed: int = 0
    warnings: int = 0
    errors: int = 0

    @property
    def summary(self) -> str:
        parts = [f"验证报告: {self.file_path}"]
        parts.append(f"  总计: {self.total} | 通过: {self.passed} | 警告: {self.warnings} | 错误: {self.errors}")
        for c in self.checks:
            icon = "✅" if c.passed else ("⚠️" if c.severity == "warn" else "❌")
            parts.append(f"  {icon} {c.check_name}: {c.message}")
        return "\n".join(parts)

    def to_dict(self) -> dict:
        return {
            "file_path": self.file_path,
            "total": self.total,
            "passed": self.passed,
            "warnings": self.warnings,
            "errors": self.errors,
            "checks": [
                {"id": c.check_id or c.check_name, "name": c.check_name,
                 "passed": c.passed, "severity": c.severity,
                 "message": c.message, "details": c.details}
                for c in self.checks
            ],
        }


def _add_check(report: ValidationReport, name: str, passed: bool,
               severity: str = "info", message: str = "", details: str = "",
               check_id: str = ""):
    """添加检查结果"""
    report.checks.append(CheckResult(name, passed, severity, message, details, check_id))
    report.total += 1
    if passed:
        report.passed += 1
    elif severity == "error":
        report.errors += 1
    else:
        report.warnings += 1


def _safe_line_spacing(pf) -> float | None:
    """安全提取段落行距数值，兼容枚举值"""
    ls = pf.line_spacing
    if ls is None:
        return None
    if isinstance(ls, (int, float)):
        return round(float(ls), 2)
    # 枚举值（如 WD_LINE_SPACING.ONE_POINT_FIVE），无法 round
    return None


def _has_page_break(run) -> bool:
    """检测 run 中是否包含分页符（安全 XML 检查）"""
    try:
        for child in run._element:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "br":
                if child.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type") == "page":
                    return True
        return False
    except Exception:
        return False


def _has_drawing(run) -> bool:
    """检测 run 中是否包含图片（安全 XML 检查）"""
    try:
        for child in run._element:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "drawing":
                return True
        return False
    except Exception:
        return False


def _get_heading_level(p) -> int | None:
    """安全提取段落标题层级"""
    style = p.style
    if not style or not style.name:
        return None
    m = re.search(r"(?:Heading|heading)\s*(\d+)", style.name)
    if m:
        return int(m.group(1))
    return None


# ── 确定性检查辅助 ───────────────────────────────────────────────

def _story_roots(doc) -> list:
    """Return body plus every header/footer story XML root."""
    roots = [doc.element.body]
    for section in doc.sections:
        for attr in (
            "header", "footer", "first_page_header", "first_page_footer",
            "even_page_header", "even_page_footer",
        ):
            try:
                part = getattr(section, attr)
                element = getattr(part, "_element", None)
                if element is not None:
                    roots.append(element)
            except Exception:
                continue
    return roots


def _collect_explicit_fonts(doc) -> list[str]:
    """Collect literal (non-theme) font names used in body/header/footer runs."""
    from docx.oxml.ns import qn

    names: set[str] = set()
    for root in _story_roots(doc):
        for rfonts in root.iter(qn("w:rFonts")):
            for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                value = rfonts.get(qn(f"w:{attr}"))
                if value and not value.startswith("+"):
                    names.add(value.strip())
    return sorted(names, key=str.lower)


def _font_install_issues(doc) -> tuple[list[str], list[str]]:
    """Return (missing fonts, explicit fonts) for the current document."""
    fonts = _collect_explicit_fonts(doc)
    if os.name != "nt":
        return [], fonts
    from .office_convert import _font_key, installed_font_keys

    installed = installed_font_keys()
    missing = [name for name in fonts if _font_key(name) not in installed]
    return missing, fonts


def _field_instructions(doc) -> list[str]:
    """Return PAGE/TOC field instructions from body, headers and footers."""
    from docx.oxml.ns import qn

    instructions: list[str] = []
    for root in _story_roots(doc):
        for element in root.iter(qn("w:fldSimple")):
            instructions.append(element.get(qn("w:instr")) or "")
        for element in root.iter(qn("w:instrText")):
            instructions.append(element.text or "")
    return instructions


def _toc_marked_for_update(doc) -> bool:
    """Whether a TOC field is dirty or the document asks Word to update fields."""
    from docx.oxml.ns import qn

    truthy = {"true", "1", "on"}
    for root in _story_roots(doc):
        for element in root.iter(qn("w:fldSimple")):
            instruction = element.get(qn("w:instr")) or ""
            if re.search(r"\bTOC\b", instruction, re.IGNORECASE):
                if (element.get(qn("w:dirty")) or "").lower() in truthy:
                    return True
        for element in root.iter(qn("w:fldChar")):
            if (element.get(qn("w:dirty")) or "").lower() in truthy:
                return True
    try:
        settings = doc.settings.element
        for element in settings.iter(qn("w:updateFields")):
            if (element.get(qn("w:val")) or "").lower() in truthy:
                return True
    except Exception:
        pass
    return False


_CN_DIGITS = {
    "零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
    "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
}


def _cn_number(text: str) -> int | None:
    """Parse Chinese numerals up to 999 (一、十二、二十三、一百零五)."""
    if not text:
        return None
    if "百" in text:
        head, _, tail = text.partition("百")
        hundreds = _CN_DIGITS.get(head, 1) if head else 1
        rest = _cn_number(tail) if tail else 0
        return hundreds * 100 + (rest or 0)
    if "十" in text:
        head, _, tail = text.partition("十")
        tens = _CN_DIGITS.get(head, 1) if head else 1
        ones = _CN_DIGITS.get(tail, 0) if tail else 0
        return tens * 10 + ones
    value = 0
    for char in text:
        if char not in _CN_DIGITS:
            return None
        value = value * 10 + _CN_DIGITS[char]
    return value or None


def _heading_number(text: str) -> tuple[int, ...] | None:
    """Extract a manual heading number such as ``1.2``, ``一、`` or ``第一章``."""
    value = (text or "").strip()
    if not value:
        return None
    match = re.match(r"^(\d+(?:\.\d+)*)(?:\s*[.、．)）]|\s|$)", value)
    if match:
        return tuple(int(part) for part in match.group(1).split("."))
    match = re.match(r"^[（(]([一二三四五六七八九十百零〇两]+)[)）]", value)
    if match:
        number = _cn_number(match.group(1))
        return (number,) if number else None
    match = re.match(r"^第([一二三四五六七八九十百零〇两]+)[章节篇部分]", value)
    if match:
        number = _cn_number(match.group(1))
        return (number,) if number else None
    match = re.match(r"^([一二三四五六七八九十百零〇两]+)[、.．]\s*", value)
    if match:
        number = _cn_number(match.group(1))
        return (number,) if number else None
    return None


def _heading_numbering_issues(
    heading_levels: list[tuple[int, str]],
) -> tuple[list[str], int]:
    """Check that same-parent headings of the same level increment by one."""
    issues: list[str] = []
    last_by_level: dict[int, tuple[tuple[int, ...], str]] = {}
    numbered = 0
    for level, text in heading_levels:
        number = _heading_number(text)
        if not number:
            continue
        numbered += 1
        previous = last_by_level.get(level)
        if previous:
            previous_number, previous_text = previous
            if number[:-1] == previous_number[:-1]:
                if number[-1] != previous_number[-1] + 1:
                    issues.append(
                        f"「{previous_text[:20]}」→「{text[:20]}」"
                        f"编号 {'.'.join(map(str, previous_number))}"
                        f"→{'.'.join(map(str, number))} 不连续"
                    )
        last_by_level[level] = (number, text)
    return issues, numbered


def _table_column_issues(table, table_index: int) -> list[str]:
    """Check grid width, per-row column counts and vertical-merge consistency."""
    from docx.oxml.ns import qn

    issues: list[str] = []
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    grid_columns = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
    rows = tbl.findall(qn("w:tr"))
    first_total: int | None = None
    active_merges: set[int] = set()

    for row_index, row in enumerate(rows, start=1):
        cells = row.findall(qn("w:tc"))
        total = 0
        next_merges: set[int] = set()
        for cell in cells:
            span = 1
            vertical_merge = None
            tc_pr = cell.find(qn("w:tcPr"))
            if tc_pr is not None:
                grid_span = tc_pr.find(qn("w:gridSpan"))
                if grid_span is not None:
                    try:
                        span = int(grid_span.get(qn("w:val")) or "1")
                    except ValueError:
                        span = 1
                        issues.append(
                            f"表 {table_index} 第 {row_index} 行 gridSpan 不是整数"
                        )
                    if span < 1:
                        issues.append(
                            f"表 {table_index} 第 {row_index} 行 gridSpan={span} 非法"
                        )
                        span = 1
                merge = tc_pr.find(qn("w:vMerge"))
                if merge is not None:
                    vertical_merge = (merge.get(qn("w:val")) or "continue").lower()
            if grid_columns and span > grid_columns:
                issues.append(
                    f"表 {table_index} 第 {row_index} 行 gridSpan={span} "
                    f"超过网格 {grid_columns} 列"
                )
            if vertical_merge == "continue" and total not in active_merges:
                issues.append(
                    f"表 {table_index} 第 {row_index} 行第 {total + 1} 列"
                    "垂直合并缺少 restart"
                )
            if vertical_merge in {"continue", "restart"}:
                for column in range(total, total + span):
                    next_merges.add(column)
            total += span

        if first_total is None:
            first_total = total
        elif total != first_total:
            issues.append(
                f"表 {table_index} 第 {row_index} 行 {total} 列，"
                f"与首行 {first_total} 列不一致"
            )
        if grid_columns and total != grid_columns:
            issues.append(
                f"表 {table_index} 第 {row_index} 行 {total} 列，"
                f"与表格网格 {grid_columns} 列不一致"
            )
        active_merges = next_merges
    return issues


def _empty_paragraph_stats(doc) -> tuple[int, int, int]:
    """Return (total body paragraphs, empty count, longest empty run)."""
    paragraphs = doc.paragraphs
    flags: list[bool] = []
    for paragraph in paragraphs:
        if paragraph.text.strip():
            flags.append(False)
            continue
        has_visual_content = any(
            _has_drawing(run) or _has_page_break(run) for run in paragraph.runs
        )
        flags.append(not has_visual_content)
    empty_count = sum(flags)
    longest = current = 0
    for flag in flags:
        current = current + 1 if flag else 0
        longest = max(longest, current)
    return len(paragraphs), empty_count, longest


_CJK_PATTERN = re.compile(r"[\u3400-\u9fff\uf900-\ufaff]")
_LATIN_PATTERN = re.compile(r"[A-Za-z]")


def _iter_content_paragraphs(doc):
    """Yield body, table-cell, header and footer paragraphs."""
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    for section in doc.sections:
        for attr in (
            "header", "footer", "first_page_header", "first_page_footer",
            "even_page_header", "even_page_footer",
        ):
            try:
                part = getattr(section, attr)
            except Exception:
                continue
            yield from part.paragraphs


def _run_explicit_fonts(run) -> dict[str, str]:
    from docx.oxml.ns import qn

    r_pr = run._element.find(qn("w:rPr"))
    if r_pr is None:
        return {}
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        return {}
    fonts: dict[str, str] = {}
    for attr in ("ascii", "hAnsi", "eastAsia"):
        value = r_fonts.get(qn(f"w:{attr}"))
        if value:
            fonts[attr] = value
    return fonts


def _mixed_font_issues(doc) -> list[str]:
    """Detect multiple fonts per script in one paragraph or script-only runs."""
    issues: list[str] = []
    for paragraph in _iter_content_paragraphs(doc):
        east_asian: set[str] = set()
        latin: set[str] = set()
        for run in paragraph.runs:
            fonts = _run_explicit_fonts(run)
            if not fonts:
                continue
            if fonts.get("eastAsia"):
                east_asian.add(fonts["eastAsia"])
            for key in ("ascii", "hAnsi"):
                if fonts.get(key):
                    latin.add(fonts[key])
            text = run.text or ""
            has_cjk = bool(_CJK_PATTERN.search(text))
            has_latin = bool(_LATIN_PATTERN.search(text))
            if has_cjk and not has_latin and not fonts.get("eastAsia") and fonts.get("ascii"):
                issues.append(f"中文 run 只设置西文字体 {fonts['ascii']}")
            if has_latin and not has_cjk and not (fonts.get("ascii") or fonts.get("hAnsi")) and fonts.get("eastAsia"):
                issues.append(f"英文 run 只设置中文字体 {fonts['eastAsia']}")
        if len(east_asian) > 1:
            issues.append("同段中文字体混用: " + "、".join(sorted(east_asian)))
        if len(latin) > 1:
            issues.append("同段英文字体混用: " + "、".join(sorted(latin)))
    return list(dict.fromkeys(issues))[:20]


def validate_word_document(file_path: str) -> ValidationReport:
    """对 Word 文档进行后处理质量检查"""
    report = ValidationReport(file_path=file_path)

    if not os.path.isfile(file_path):
        _add_check(report, "文件存在", False, "error", "文件不存在", check_id="file_exists")
        return report

    try:
        from .utils import validate_input_file
        validate_input_file(file_path, (".docx",))
    except ValueError as e:
        _add_check(report, "输入文件安全检查", False, "error", str(e), check_id="input_safe")
        return report

    try:
        from docx import Document

        doc = Document(file_path)

        # ── 检查 1: 文件基础信息 ──
        file_size = os.path.getsize(file_path)
        size_ok = file_size > 0
        _add_check(report, "文件大小检查", size_ok, "error" if not size_ok else "info",
                   f"{file_size} 字节" if size_ok else "空文件", check_id="file_size")

        para_count = len(doc.paragraphs)
        _add_check(report, "段落数量", True, "info", f"{para_count} 个段落",
                   check_id="paragraph_count")

        # ── 检查 2: 空白页检查 ──
        empty_para_count = sum(1 for p in doc.paragraphs if not p.text.strip())
        page_break_count = 0
        for p in doc.paragraphs:
            for run in p.runs:
                if _has_page_break(run):
                    page_break_count += 1

        if empty_para_count > 5 and page_break_count > 2:
            _add_check(report, "空白页检测", False, "warn",
                       f"检测到 {empty_para_count} 个空段落和 {page_break_count} 个分页符，可能存在空白页",
                       f"空段落: {empty_para_count}, 分页符: {page_break_count}",
                       check_id="blank_pages")
        else:
            _add_check(report, "空白页检测", True, "info",
                       f"空段落: {empty_para_count}, 分页符: {page_break_count}",
                       check_id="blank_pages")

        # ── 检查 3: 行间距一致性 ──
        line_spacings = set()
        for p in doc.paragraphs:
            pf = p.paragraph_format
            if pf.line_spacing is not None and p.text.strip():
                ls = _safe_line_spacing(pf)
                if ls is not None:
                    line_spacings.add(ls)
        if len(line_spacings) > 3:
            _add_check(report, "行间距一致性", False, "warn",
                       f"检测到 {len(line_spacings)} 种不同行间距: {sorted(line_spacings)}",
                       check_id="line_spacing")
        else:
            _add_check(report, "行间距一致性", True, "info",
                       f"行间距: {sorted(line_spacings) if line_spacings else '默认'}",
                       check_id="line_spacing")

        # ── 检查 4: 表格完整性 ──
        table_count = len(doc.tables)
        table_issues = 0
        for ti, table in enumerate(doc.tables):
            if len(table.rows) == 0:
                table_issues += 1
                continue
            has_content = False
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        has_content = True
                        break
            if not has_content:
                table_issues += 1
        if table_issues > 0:
            _add_check(report, "表格完整性", False, "warn",
                       f"共 {table_count} 个表格，{table_issues} 个可能存在空表问题",
                       check_id="table_integrity")
        else:
            _add_check(report, "表格完整性", True, "info", f"共 {table_count} 个表格",
                       check_id="table_integrity")

        # ── 检查 5: 标题层级连续（报告所有跳跃，不只是第一个）──
        heading_levels = []
        for p in doc.paragraphs:
            level = _get_heading_level(p)
            if level is not None:
                heading_levels.append((level, p.text.strip()[:40]))

        if heading_levels:
            jumps_found = False
            for i in range(1, len(heading_levels)):
                prev_lvl = heading_levels[i - 1][0]
                curr_lvl = heading_levels[i][0]
                if curr_lvl > prev_lvl + 1:
                    _add_check(report, "标题层级连续性", False, "warn",
                               f"标题从 H{prev_lvl} 跳转到 H{curr_lvl}（第 {i + 1} 个标题）",
                               check_id="heading_levels")
                    jumps_found = True
            if not jumps_found:
                levels_only = [level for level, _ in heading_levels]
                _add_check(report, "标题层级连续性", True, "info",
                           f"标题层级: H{min(levels_only)}-H{max(levels_only)}",
                           check_id="heading_levels")
        else:
            _add_check(report, "标题层级连续性", True, "info", "无标题（可能为纯文档）",
                       check_id="heading_levels")

        # ── 检查 6: 图片嵌入检测 ──
        image_count = 0
        for p in doc.paragraphs:
            for run in p.runs:
                if _has_drawing(run):
                    image_count += 1
        _add_check(report, "图片嵌入检查", True, "info", f"检测到 {image_count} 个图片",
                   check_id="images")

        # ── 检查 7: section/分节设置 ──
        section_count = len(doc.sections)
        if section_count > 1:
            _add_check(report, "分节设置", True, "info", f"共 {section_count} 个节",
                       check_id="sections")
        else:
            _add_check(report, "分节设置", True, "info", "单节文档", check_id="sections")

        # ── 检查 8: 页眉页脚 ──
        hf_found = False
        for section in doc.sections:
            if section.header.paragraphs and any(p.text.strip() for p in section.header.paragraphs):
                hf_found = True
                break
            if section.footer.paragraphs and any(p.text.strip() for p in section.footer.paragraphs):
                hf_found = True
                break
        if hf_found:
            _add_check(report, "页眉页脚", True, "info", "已检测到页眉页脚内容",
                       check_id="header_footer")
        else:
            _add_check(report, "页眉页脚", True, "info", "无页眉页脚（可能是封面/简历等）",
                       check_id="header_footer")

        # ── 检查 9: 遗留占位标记 ──
        marker_pattern = re.compile(r"\b(?:TODO|TBD|DRAFT)\b|待(?:完善|补充|确认)", re.IGNORECASE)
        text_parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            text_parts.extend(cell.text for row in table.rows for cell in row.cells)
        for section in doc.sections:
            text_parts.extend(p.text for p in section.header.paragraphs)
            text_parts.extend(p.text for p in section.footer.paragraphs)
        markers = sorted({match.group(0) for text in text_parts for match in marker_pattern.finditer(text)})
        if markers:
            _add_check(report, "遗留占位标记", False, "warn", f"发现: {', '.join(markers)}",
                       check_id="placeholders")
        else:
            _add_check(report, "遗留占位标记", True, "info", "未发现 TODO/TBD/DRAFT 或待完善标记",
                       check_id="placeholders")

        # ── 检查 10: 字体是否本机已安装 ──
        try:
            missing_fonts, explicit_fonts = _font_install_issues(doc)
            if os.name != "nt":
                _add_check(report, "字体安装检查", True, "info",
                           "非 Windows 环境，跳过字体安装检查",
                           check_id="font_installed")
            elif not explicit_fonts:
                _add_check(report, "字体安装检查", True, "info",
                           "未显式指定字体（继承样式/主题）",
                           check_id="font_installed")
            elif missing_fonts:
                _add_check(report, "字体安装检查", False, "warn",
                           f"字体未安装: {', '.join(missing_fonts)}",
                           f"文档显式引用 {len(explicit_fonts)} 种字体: "
                           f"{', '.join(explicit_fonts)}",
                           check_id="font_installed")
            else:
                _add_check(report, "字体安装检查", True, "info",
                           f"字体均已安装（{len(explicit_fonts)} 种）",
                           check_id="font_installed")
        except Exception as exc:
            _add_check(report, "字体安装检查", False, "warn",
                       f"字体检查失败: {exc}", check_id="font_installed")

        # ── 检查 11: 标题编号连续性 ──
        try:
            numbering_issues, numbered_count = _heading_numbering_issues(heading_levels)
            if numbering_issues:
                _add_check(report, "标题编号连续性", False, "warn",
                           f"标题编号不连续（{len(numbering_issues)} 处）",
                           "；".join(numbering_issues[:8]),
                           check_id="heading_numbering")
            elif numbered_count:
                _add_check(report, "标题编号连续性", True, "info",
                           f"标题编号连续（{numbered_count} 个带编号标题）",
                           check_id="heading_numbering")
            else:
                _add_check(report, "标题编号连续性", True, "info",
                           "未检测到标题编号", check_id="heading_numbering")
        except Exception as exc:
            _add_check(report, "标题编号连续性", False, "warn",
                       f"标题编号检查失败: {exc}", check_id="heading_numbering")

        # ── 检查 12: PAGE 页码域 ──
        try:
            instructions = _field_instructions(doc)
            has_page_field = any(
                re.search(r"\bPAGE\b", instruction, re.IGNORECASE)
                for instruction in instructions
            )
            if has_page_field:
                _add_check(report, "页码域检查", True, "info", "检测到 PAGE 域",
                           check_id="page_number_field")
            else:
                _add_check(report, "页码域检查", False, "warn",
                           "未检测到 PAGE 域（页码可能为静态文本或缺失）",
                           check_id="page_number_field")
        except Exception as exc:
            _add_check(report, "页码域检查", False, "warn",
                       f"页码域检查失败: {exc}", check_id="page_number_field")

        # ── 检查 13: TOC 目录域 ──
        try:
            has_toc_field = any(
                re.search(r"\bTOC\b", instruction, re.IGNORECASE)
                for instruction in instructions
            )
            heading_count = len(heading_levels)
            if has_toc_field:
                if _toc_marked_for_update(doc):
                    _add_check(report, "目录域检查", True, "info",
                               "目录域存在且标记为需更新",
                               check_id="toc_field")
                else:
                    _add_check(report, "目录域检查", False, "warn",
                               "目录域存在但未标记为需更新（dirty/updateFields）",
                               "Word 打开后不会自动刷新目录",
                               check_id="toc_field")
            elif heading_count >= 3:
                _add_check(report, "目录域检查", False, "warn",
                           "未检测到目录域（TOC）",
                           f"文档有 {heading_count} 个标题；长文档建议插入 TOC 域",
                           check_id="toc_field")
            else:
                _add_check(report, "目录域检查", True, "info",
                           f"标题较少（{heading_count} 个），未要求目录域",
                           check_id="toc_field")
        except Exception as exc:
            _add_check(report, "目录域检查", False, "warn",
                       f"目录域检查失败: {exc}", check_id="toc_field")

        # ── 检查 14: 表格列数与合并单元格结构 ──
        try:
            column_issues: list[str] = []
            for table_index, table in enumerate(doc.tables, start=1):
                column_issues.extend(_table_column_issues(table, table_index))
            if column_issues:
                _add_check(report, "表格列数/合并结构", False, "warn",
                           f"表格列数或合并单元格异常（{len(column_issues)} 处）",
                           "；".join(column_issues[:8]),
                           check_id="table_columns")
            else:
                _add_check(report, "表格列数/合并结构", True, "info",
                           f"表格列数与合并结构一致（{len(doc.tables)} 个表格）",
                           check_id="table_columns")
        except Exception as exc:
            _add_check(report, "表格列数/合并结构", False, "warn",
                       f"表格结构检查失败: {exc}", check_id="table_columns")

        # ── 检查 15: 空段落/连续空行比例 ──
        try:
            total_paras, empty_count, longest_empty_run = _empty_paragraph_stats(doc)
            empty_ratio = (empty_count / total_paras) if total_paras else 0.0
            suspicious = longest_empty_run >= 3 or (
                total_paras >= 10 and empty_ratio > 0.3
            )
            details = (
                f"空段落 {empty_count}/{total_paras}（{empty_ratio:.0%}），"
                f"最长连续空段落 {longest_empty_run}"
            )
            if suspicious:
                _add_check(report, "空段落比例", False, "warn",
                           "空段落或连续空行偏多", details,
                           check_id="empty_paragraph_ratio")
            else:
                _add_check(report, "空段落比例", True, "info", details,
                           check_id="empty_paragraph_ratio")
        except Exception as exc:
            _add_check(report, "空段落比例", False, "warn",
                       f"空段落检查失败: {exc}", check_id="empty_paragraph_ratio")

        # ── 检查 16: 中英文字体混用 ──
        try:
            mixed_fonts = _mixed_font_issues(doc)
            if mixed_fonts:
                _add_check(report, "中英文字体混用", False, "warn",
                           f"检测到 {len(mixed_fonts)} 处中英文字体混用",
                           "；".join(mixed_fonts[:8]),
                           check_id="mixed_fonts")
            else:
                _add_check(report, "中英文字体混用", True, "info",
                           "同段/同 run 未发现中英文字体混用",
                           check_id="mixed_fonts")
        except Exception as exc:
            _add_check(report, "中英文字体混用", False, "warn",
                       f"中英文字体检查失败: {exc}", check_id="mixed_fonts")

    except Exception as e:
        _add_check(report, "文档解析", False, "error", f"无法解析文档: {e}", check_id="parse")

    logger.info("文档验证完成: %s", report.summary.replace("\n", " | "))
    return report
