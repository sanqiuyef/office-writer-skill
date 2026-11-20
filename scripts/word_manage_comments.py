#!/usr/bin/env python3
"""管理 Word 批注：添加批注或列出已有批注。"""

from __future__ import annotations

import argparse

from docx import Document
from docx.oxml.ns import qn

import _common as common
import _native as native
from office_engine import utils
from office_engine.models import WordCommentDef

TOOL = "word_manage_comments"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "添加 Word 批注（按 anchor_text 或 paragraph_index 锚定）或列出已有批注；写文件自动生成 .bak 备份。",
    )
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--list", dest="list", action="store_true", default=None)
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    want_list = common.bool_arg(args, params, "list", False)
    additions = common.model_list_arg(params, "add", WordCommentDef, required=False)

    try:
        utils.validate_input_file(file_path, (".docx",))
    except ValueError as exc:
        raise common.ToolRuntimeError(str(exc)) from exc
    try:
        doc = Document(file_path)
    except Exception as exc:
        raise common.ToolRuntimeError(f"无法打开 Word 文档: {file_path}（{exc}）") from exc

    def _resolve_runs(comment: WordCommentDef):
        """返回批注锚定的 runs；定位不到时直接报错。"""
        if comment.anchor_text:
            needle = comment.anchor_text if comment.match_case else comment.anchor_text.lower()
            for paragraph in doc.paragraphs:
                for run in paragraph.runs:
                    haystack = run.text if comment.match_case else run.text.lower()
                    if needle in haystack:
                        return [run]
            raise common.UsageError(f"未找到包含锚定文字的 run：{comment.anchor_text!r}")
        index = comment.paragraph_index
        if index is None or index < 1 or index > len(doc.paragraphs):
            raise common.UsageError(
                f"paragraph_index 超出范围：{index}（文档共 {len(doc.paragraphs)} 个正文段落）"
            )
        paragraph = doc.paragraphs[index - 1]
        if not paragraph.runs:
            raise common.UsageError(f"第 {index} 个正文段落没有 run，无法添加批注")
        return list(paragraph.runs)

    resolved = [(comment, _resolve_runs(comment)) for comment in additions]
    for comment, runs in resolved:
        doc.add_comment(
            runs=runs,
            text=comment.text,
            author=comment.author,
            initials=comment.initials,
        )

    added = len(resolved)
    backup_path = utils.atomic_save(doc.save, file_path) if added else None

    def _anchor_text(comment_id: int) -> str | None:
        """从正文 XML 提取 commentRangeStart/End 之间的文字。"""
        capturing = False
        pieces: list[str] = []
        for element in doc.element.body.iter():
            if element.tag == qn("w:commentRangeStart"):
                if element.get(qn("w:id")) == str(comment_id):
                    capturing = True
                continue
            if element.tag == qn("w:commentRangeEnd"):
                if element.get(qn("w:id")) == str(comment_id):
                    break
                continue
            if capturing and element.tag == qn("w:t") and element.text:
                pieces.append(element.text)
        return "".join(pieces) or None

    include_comments = want_list or not additions
    comments = [
        {
            "comment_id": comment.comment_id,
            "author": comment.author,
            "initials": comment.initials,
            "text": comment.text,
            "timestamp": comment.timestamp.isoformat() if comment.timestamp else None,
            "anchor_text": _anchor_text(comment.comment_id),
        }
        for comment in doc.comments
    ] if include_comments else []

    return {
        "success": True,
        "file_path": file_path,
        "added": added,
        "comments": comments,
        "backup_path": backup_path,
    }, 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
