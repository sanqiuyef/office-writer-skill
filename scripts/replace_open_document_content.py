#!/usr/bin/env python3
"""在已打开的 Word/PPT/Excel（含 WPS）文档上替换内容，不重建文档。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.models import EditResult, ReplaceSection, TextReplacement
from office_engine.office_com import com_is_available
from office_engine.open_replace import try_replace_open_document_content

TOOL = "replace_open_document_content"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(
        TOOL,
        "在已打开的 Word/PPT/Excel（含 WPS）文档上按标题替换整节或查找替换并保存；不重生成文档。",
    )
    parser.add_argument("--file-path", dest="file_path")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    sections = common.model_list_arg(params, "replace_sections", ReplaceSection, required=False)
    replacements = common.model_list_arg(params, "replace_texts", TextReplacement, required=False)
    if not sections and not replacements:
        raise common.UsageError("至少提供 replace_sections 或 replace_texts 之一")
    if not com_is_available():
        result = EditResult(
            success=False,
            file_path=file_path,
            message="Windows COM 不可用；请在安装 pywin32 的 Windows 环境中使用活动文档编辑",
        )
    else:
        timeout_result = EditResult(
            success=False,
            file_path=file_path,
            message=f"活动文档操作超时（>{common.COM_OPERATION_TIMEOUT}s），可能的应用未响应",
        )
        result = common.with_timeout(
            try_replace_open_document_content,
            common.COM_OPERATION_TIMEOUT,
            timeout_result,
            file_path,
            sections,
            replacements,
        ) or EditResult(
            success=False,
            file_path=file_path,
            message="目标文件未在活动的 Word/PPT/Excel（或 WPS）实例中打开",
        )
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
