#!/usr/bin/env python3
"""向已打开的 PowerPoint 演示文稿追加流程图幻灯片。"""

from __future__ import annotations

import argparse

import _common as common
from office_engine.models import EditResult, FrameworkDiagram
from office_engine.office_com import com_is_available, try_add_flowchart_slide

TOOL = "ppt_add_flowchart_slide"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "向已打开的 PowerPoint 演示文稿末尾追加一张流程图幻灯片（COM 实时绘制）。")
    parser.add_argument("--file-path", dest="file_path")
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    diagram = common.model_arg(params, "diagram", FrameworkDiagram)
    if not com_is_available():
        result = EditResult(
            success=False,
            file_path=file_path,
            message="Windows COM 不可用；请在安装 pywin32 的 Windows 环境中使用活动演示文稿编辑",
        )
    else:
        timeout_result = EditResult(
            success=False,
            file_path=file_path,
            message=f"活动 PowerPoint 流程图操作超时（>{common.COM_OPERATION_TIMEOUT}s），可能的应用未响应",
        )
        result = common.with_timeout(
            try_add_flowchart_slide,
            common.COM_OPERATION_TIMEOUT,
            timeout_result,
            file_path,
            diagram,
        ) or EditResult(
            success=False,
            file_path=file_path,
            message="目标文件未在活动 PowerPoint 实例中打开",
        )
    common.ensure_success(result)
    return common.result_data(result), 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
