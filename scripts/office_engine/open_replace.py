"""Replace content inside Office documents that are already open (Word/PPT/Excel).

用于「修改优先于重生成」：在用户已打开的 Word/WPS、PowerPoint/WPS 演示、
Excel/WPS 表格上按标题替换整节内容或做查找替换并保存，不重建文档。

借鉴 anthropics/skills 的 docx/pptx「unpack → edit XML → zip」保真思路，
但活动文档场景必须走 COM：只有应用自己写回才能保留 MathType/OLE、修订、
批注、域等二进制对象。
"""

from __future__ import annotations

import re
from typing import Any

from .models import EditResult, FontConfig, ReplaceSection, TextReplacement
from .office_com import (
    ActiveOfficeDocumentError,
    _active_applications,
    _com_apartment,
    _excel_workbook,
    _ppt_presentation,
    _word_document,
    com_is_available,
)
from .utils import cn_size_to_pt


# ── Word ────────────────────────────────────────────────────────

def _word_text(paragraph: Any) -> str:
    try:
        return str(paragraph.Range.Text).strip()
    except Exception:
        return ""


def _word_find_paragraph(document: Any, text: str, match: str = "contains") -> int | None:
    for index in range(1, document.Paragraphs.Count + 1):
        current = _word_text(document.Paragraphs(index))
        if match == "exact":
            if current == text.strip():
                return index
        elif text.strip() in current:
            return index
    return None


def _word_outline_level(paragraph: Any) -> int:
    try:
        return int(paragraph.OutlineLevel)
    except Exception:
        return 10


def _apply_word_font(target: Any, font: FontConfig | None) -> None:
    if font is None:
        return
    if font.name:
        target.Font.NameFarEast = font.name
        target.Font.NameAscii = "Times New Roman" if font.name in ("宋体", "黑体", "楷体", "仿宋") else font.name
        target.Font.NameOther = target.Font.NameAscii
    if font.size is not None:
        target.Font.Size = cn_size_to_pt(font.size)
    if font.bold is not None:
        target.Font.Bold = font.bold
    if font.italic is not None:
        target.Font.Italic = font.italic


def _replace_word_section(document: Any, section: ReplaceSection) -> None:
    heading_index = _word_find_paragraph(document, section.heading_text, section.heading_match)
    if heading_index is None:
        raise ActiveOfficeDocumentError(f"未找到标题: {section.heading_text!r}")
    heading = document.Paragraphs(heading_index)
    level = _word_outline_level(heading)

    end_index = document.Paragraphs.Count + 1
    if section.end_heading_text:
        found = _word_find_paragraph(document, section.end_heading_text, "contains")
        if found is None or found <= heading_index:
            raise ActiveOfficeDocumentError(f"未找到结束标题: {section.end_heading_text!r}")
        end_index = found
    else:
        for index in range(heading_index + 1, document.Paragraphs.Count + 1):
            if _word_outline_level(document.Paragraphs(index)) <= level:
                end_index = index
                break

    start = heading.Range.End if section.keep_heading else heading.Range.Start
    end = document.Paragraphs(end_index).Range.Start if end_index <= document.Paragraphs.Count else document.Content.End
    if end > start:
        document.Range(start, end).Delete()

    if section.lines:
        payload = "\r".join(section.lines) + "\r"
        rng = document.Range(start, start)
        rng.InsertAfter(payload)
        _apply_word_font(document.Range(start, start + len(payload)), section.font)


def _replace_word_texts(document: Any, replacements: list[TextReplacement]) -> int:
    changed = 0
    for item in replacements:
        finder = document.Content.Find
        finder.ClearFormatting()
        finder.Replacement.ClearFormatting()
        finder.Text = item.find
        finder.Replacement.Text = item.replace
        finder.MatchCase = item.match_case
        finder.Forward = True
        finder.Wrap = 1  # wdFindContinue
        if finder.Execute(Replace=2):  # wdReplaceAll
            changed += 1
    return changed


def try_replace_open_word_content(
    file_path: str,
    sections: list[ReplaceSection],
    replacements: list[TextReplacement],
) -> EditResult | None:
    if not com_is_available():
        return None
    try:
        with _com_apartment():
            application = None
            prog_id = None
            document = None
            for candidate, candidate_prog_id in _active_applications(("Word.Application", "kwps.Application")):
                candidate_document = _word_document(candidate, file_path)
                if candidate_document is not None:
                    application, prog_id, document = candidate, candidate_prog_id, candidate_document
                    break
            if document is None or application is None or prog_id is None:
                return None
            try:
                if bool(document.ReadOnly):
                    raise ActiveOfficeDocumentError("活动 Word 文档为只读，无法保存")
                undo_record = getattr(application, "UndoRecord", None)
                if undo_record is not None:
                    undo_record.StartCustomRecord("mcp-office-writer replace")
                try:
                    for section in sections:
                        _replace_word_section(document, section)
                    text_ops = _replace_word_texts(document, replacements)
                    document.Save()
                except Exception:
                    try:
                        if undo_record is not None:
                            undo_record.EndCustomRecord()
                            application.Undo()
                    except Exception:
                        pass
                    raise
                if undo_record is not None:
                    undo_record.EndCustomRecord()
                total = len(sections) + text_ops
                return EditResult(
                    success=True,
                    file_path=file_path,
                    operation_count=total,
                    message=f"已通过 {prog_id} 替换 {len(sections)} 节、{text_ops} 处文本并保存",
                )
            except Exception as exc:
                return EditResult(success=False, file_path=file_path, message=f"活动 Word 文档替换失败: {exc}")
    except Exception as exc:
        return EditResult(success=False, file_path=file_path, message=f"活动 Word 文档操作失败: {exc}")


# ── PowerPoint ─────────────────────────────────────────────────

def _replace_ppt_texts(presentation: Any, replacements: list[TextReplacement]) -> int:
    changed = 0
    for slide_index in range(1, presentation.Slides.Count + 1):
        slide = presentation.Slides(slide_index)
        for shape_index in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(shape_index)
            try:
                if not sk_has_text(shape):
                    continue
                text_range = shape.TextFrame.TextRange
                original = str(text_range.Text)
                updated = original
                for item in replacements:
                    if not item.find:
                        continue
                    if item.match_case:
                        updated = updated.replace(item.find, item.replace)
                    else:
                        updated = re.sub(re.escape(item.find), item.replace.replace("\\", "\\\\"), updated, flags=re.IGNORECASE)
                if updated != original:
                    text_range.Text = updated
                    changed += 1
            except Exception:
                continue
    return changed


def sk_has_text(shape: Any) -> bool:
    try:
        return bool(shape.HasTextFrame) and bool(shape.TextFrame.HasText)
    except Exception:
        return False


def try_replace_open_powerpoint_content(
    file_path: str,
    replacements: list[TextReplacement],
) -> EditResult | None:
    if not com_is_available():
        return None
    try:
        with _com_apartment():
            application = None
            prog_id = None
            presentation = None
            for candidate, candidate_prog_id in _active_applications(("PowerPoint.Application", "kwpp.Application")):
                candidate_presentation = _ppt_presentation(candidate, file_path)
                if candidate_presentation is not None:
                    application, prog_id, presentation = candidate, candidate_prog_id, candidate_presentation
                    break
            if presentation is None or application is None or prog_id is None:
                return None
            try:
                changed = _replace_ppt_texts(presentation, replacements)
                try:
                    presentation.Save()
                except Exception as save_exc:
                    # PowerPoint 对某些文件（共享锁/外部生成）会拒绝原路保存；
                    # 先尝试同路径 SaveAs，再失败则明确报错，提示用户手工保存。
                    try:
                        presentation.SaveAs(presentation.FullName)
                    except Exception:
                        raise ActiveOfficeDocumentError(
                            f"PowerPoint 保存被拒绝（文件可能被其他进程占用），请在 PowerPoint 中按 Ctrl+S 手工保存：{save_exc}")
                return EditResult(
                    success=True,
                    file_path=file_path,
                    operation_count=changed,
                    message=f"已通过 {prog_id} 替换 {changed} 处文本并保存",
                )
            except Exception as exc:
                return EditResult(success=False, file_path=file_path, message=f"活动 PowerPoint 替换失败: {exc}")
    except Exception as exc:
        return EditResult(success=False, file_path=file_path, message=f"活动 PowerPoint 操作失败: {exc}")


# ── Excel ──────────────────────────────────────────────────────

def _replace_excel_values(workbook: Any, replacements: list[TextReplacement]) -> int:
    changed = 0
    for sheet_index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets(sheet_index)
        for item in replacements:
            if not item.find:
                continue
            try:
                used = sheet.UsedRange
                if int(used.Cells.Count) <= 0:
                    continue
                if used.Replace(What=item.find, Replacement=item.replace, LookAt=2, SearchOrder=1, MatchCase=item.match_case):
                    changed += 1
            except Exception:
                continue
    return changed


def try_replace_open_excel_content(
    file_path: str,
    replacements: list[TextReplacement],
) -> EditResult | None:
    if not com_is_available():
        return None
    try:
        with _com_apartment():
            application = None
            prog_id = None
            workbook = None
            for candidate, candidate_prog_id in _active_applications(("Excel.Application", "ket.Application")):
                candidate_workbook = _excel_workbook(candidate, file_path)
                if candidate_workbook is not None:
                    application, prog_id, workbook = candidate, candidate_prog_id, candidate_workbook
                    break
            if workbook is None or application is None or prog_id is None:
                return None
            try:
                changed = _replace_excel_values(workbook, replacements)
                workbook.Save()
                return EditResult(
                    success=True,
                    file_path=file_path,
                    operation_count=changed,
                    message=f"已通过 {prog_id} 替换 {changed} 处文本并保存",
                )
            except Exception as exc:
                return EditResult(success=False, file_path=file_path, message=f"活动 Excel 替换失败: {exc}")
    except Exception as exc:
        return EditResult(success=False, file_path=file_path, message=f"活动 Excel 操作失败: {exc}")


# ── 分派 ───────────────────────────────────────────────────────

WORD_EXT = {".doc", ".docx", ".dot", ".dotx", ".docm"}
PPT_EXT = {".ppt", ".pptx", ".pot", ".potx", ".pptm"}
EXCEL_EXT = {".xls", ".xlsx", ".xlt", ".xltx", ".xlsm"}


def try_replace_open_document_content(
    file_path: str,
    sections: list[ReplaceSection] | None = None,
    replacements: list[TextReplacement] | None = None,
) -> EditResult | None:
    """根据扩展名分派到 Word/PPT/Excel 的活动文档替换。"""
    sections = sections or []
    replacements = replacements or []
    suffix = file_path.lower().rsplit(".", 1)[-1] if "." in file_path else ""
    suffix = "." + suffix
    if suffix in WORD_EXT:
        return try_replace_open_word_content(file_path, sections, replacements)
    if suffix in PPT_EXT:
        if sections:
            return EditResult(success=False, file_path=file_path, message="PPTX 不支持 replace_sections，请用 replace_texts")
        return try_replace_open_powerpoint_content(file_path, replacements)
    if suffix in EXCEL_EXT:
        if sections:
            return EditResult(success=False, file_path=file_path, message="XLSX 不支持 replace_sections，请用 replace_texts")
        return try_replace_open_excel_content(file_path, replacements)
    return EditResult(success=False, file_path=file_path, message=f"不支持的文件类型: {suffix}")
