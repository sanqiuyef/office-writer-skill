#!/usr/bin/env python3
"""以原子事务批量编辑 DOCX/XLSX/PPTX。"""

from __future__ import annotations

import argparse
import shutil
from datetime import datetime
from pathlib import Path

import _common as common
from office_engine.document_ops import batch_edit_office_document
from office_engine.models import OfficeOperation

TOOL = "batch_edit_office_document"


def build_parser() -> argparse.ArgumentParser:
    parser = common.json_parser(TOOL, "以原子事务批量编辑 DOCX/XLSX/PPTX（set/add/remove/move/replace）。")
    parser.add_argument("--file-path", dest="file_path")
    parser.add_argument("--output-path", dest="output_path")
    parser.add_argument(
        "--no-backup",
        dest="no_backup",
        action="store_true",
        help="省略 --output-path 原地写回时不生成 *.bak-YYYYmmdd-HHMMSS 备份",
    )
    return parser


def run(args):
    params = common.parse_params(args)
    file_path = common.arg(args, params, "file_path", required=True)
    output_path = common.arg(args, params, "output_path")
    operations = common.model_list_arg(params, "operations", OfficeOperation)
    no_backup = bool(getattr(args, "no_backup", False))
    backup_path = None
    source = Path(file_path)
    supported = {".docx", ".xlsx", ".xlsm", ".pptx"}
    if output_path is None and not no_backup and source.is_file() and source.suffix.lower() in supported:
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        candidate = source.with_name(f"{source.name}.bak-{timestamp}")
        counter = 1
        while candidate.exists():
            candidate = source.with_name(f"{source.name}.bak-{timestamp}-{counter}")
            counter += 1
        try:
            shutil.copy2(source, candidate)
        except OSError as exc:
            raise common.ToolRuntimeError(f"无法创建原地编辑备份 {candidate}: {exc}") from exc
        backup_path = str(candidate)
    result = batch_edit_office_document(file_path, operations, output_path)
    common.ensure_success(result)
    data = common.result_data(result)
    if backup_path:
        data["backup_path"] = backup_path
    return data, 0


def main(argv=None):
    return common.cli_main(TOOL, build_parser(), run, argv)


if __name__ == "__main__":
    raise SystemExit(main())
