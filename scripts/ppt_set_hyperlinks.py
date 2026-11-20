#!/usr/bin/env python3
"""为已有 PPT 的形状或形状内文字设置超链接。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.models import PptHyperlinkDef
from office_engine.utils import atomic_save, validate_input_file

TOOL = "ppt_set_hyperlinks"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "按形状序号或形状文字定位，为整个形状或匹配的文字 run 设置超链接。",
    )
    parser.add_argument("--file-path", dest="file_path")
    return parser


def run(args):
    from pptx import Presentation

    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    links = common.model_list_arg(params, "links", PptHyperlinkDef)

    try:
        validate_input_file(file_path, (".pptx",))
    except ValueError as exc:
        raise common.ToolRuntimeError(str(exc)) from exc

    def contains(haystack: str, needle: str, match_case: bool) -> bool:
        if match_case:
            return needle in haystack
        return needle.casefold() in haystack.casefold()

    prs = Presentation(file_path)
    applied = 0

    for link in links:
        if link.slide > len(prs.slides):
            raise common.UsageError(
                f"link.slide={link.slide} 超出范围：文件共有 {len(prs.slides)} 张幻灯片"
            )
        slide = prs.slides[link.slide - 1]

        shape = None
        if link.shape_index is not None:
            if link.shape_index > len(slide.shapes):
                raise common.UsageError(
                    f"第 {link.slide} 张幻灯片只有 {len(slide.shapes)} 个形状，"
                    f"shape_index={link.shape_index} 越界"
                )
            shape = slide.shapes[link.shape_index - 1]
        else:
            for candidate in slide.shapes:
                if not getattr(candidate, "has_text_frame", False):
                    continue
                if contains(candidate.text_frame.text, link.shape_text, link.match_case):
                    shape = candidate
                    break
            if shape is None:
                raise common.UsageError(
                    f"第 {link.slide} 张幻灯片中找不到包含文字 {link.shape_text!r} 的形状"
                )

        if link.text_match:
            matched = False
            if getattr(shape, "has_text_frame", False):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if contains(run.text, link.text_match, link.match_case):
                            run.hyperlink.address = link.url
                            matched = True
            if not matched:
                raise common.UsageError(
                    f"第 {link.slide} 张幻灯片的目标形状中找不到匹配文字 "
                    f"{link.text_match!r}"
                )
        else:
            if not hasattr(shape, "click_action"):
                raise common.UsageError(
                    f"第 {link.slide} 张幻灯片的形状 #{link.shape_index or '?'} "
                    "不支持整体点击超链接"
                )
            shape.click_action.hyperlink.address = link.url
        applied += 1

    backup_path = atomic_save(prs.save, file_path)
    return {
        "success": True,
        "file_path": file_path,
        "applied": applied,
        "backup_path": backup_path,
    }, 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
