#!/usr/bin/env python3
"""编辑已在 Word/WPS 中打开的活动文档。"""

from __future__ import annotations

import argparse

import _common as common
import _native as native
from office_engine.models import EditResult, FontConfig, SectionItem
from office_engine.office_com import com_is_available, try_edit_open_word_document

TOOL = "edit_open_word_document"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "编辑已在 Word/WPS 中打开的活动文档并由该应用保存。")
    parser.add_argument("--file-path", dest="file_path")
    return parser


def run(args):
    params = common.parse_params(args)
    if native.use_csharp(TOOL):
        return native.run_entry_from_args(TOOL, args, params)
    file_path = common.arg(args, params, "file_path", required=True)
    append_sections = common.model_list_arg(params, "append_sections", SectionItem)
    default_font = common.model_arg(params, "default_font", FontConfig, required=False)
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
            message=f"活动 Word 文档操作超时（>{common.COM_OPERATION_TIMEOUT}s），可能的应用未响应",
        )
        result = common.with_timeout(
            try_edit_open_word_document,
            common.COM_OPERATION_TIMEOUT,
            timeout_result,
            file_path,
            append_sections,
            default_font,
        ) or EditResult(
            success=False,
            file_path=file_path,
            message="目标文件未在活动 Word/WPS 实例中打开",
        )
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
