# Vendored from mcp_office_writer/office_convert.py (mcp-office-writer 0.6.0, 2026-11).
# Upstream path: mcp_office_writer/office_convert.py (repository-relative).
# Modifications: xlsx/xls render uses sheet + A1 range and omits --pages, matching dsoffice's spreadsheet interface; added monospace font fallback reporting (font_substitutions) on convert; resolve dsoffice before deleting existing outputs; added OFFICE_CONVERT_BACKEND selection with Microsoft Office/WPS COM fallback for PDF conversion, page rendering and Excel recalculation, including PyMuPDF rendering and process cleanup; existing file/directory targets are backed up (or renamed aside) before replacement instead of being deleted.
"""Office 文档转换、渲染与公式重算。

三种后端按 ``OFFICE_CONVERT_BACKEND`` 选择（默认 ``auto``）：

- ``dsoffice``：DeepSeek dsh-libreoffice-kit 内置精简 LibreOffice 引擎；
- ``com``：Windows 上通过 DispatchEx 启动独立的 Microsoft Office / WPS 实例，
  导出 PDF 或重算 Excel 公式，用完 Quit，不残留进程；
- ``none``：禁用转换/渲染/重算，错误信息列出三种启用方式。

``auto`` 先找 dsoffice，缺失或执行失败时回退 COM；渲染在缺少 dsoffice 时先转
临时 PDF，再用 PyMuPDF 逐页出 PNG，保持 ``--dpi``/``--pages`` 语义。

CLI 输出为单行 JSON（失败时 {"code":"failed","error":...}，退出码非 0）。
dsoffice 拒绝覆盖已存在的输出文件/目录，本模块按 if_exists 语义先行处理。
"""
from __future__ import annotations

import ctypes
import json
import logging
import os
import re
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

from .utils import (
    backup_target,
    ensure_output_dir,
    resolve_if_exists,
    validate_input_file,
    validate_output_path,
)

logger = logging.getLogger(__name__)

# dsoffice convert/render 的输入格式（render 另支持直接渲染 PDF）
CONVERT_INPUT_SUFFIXES = (".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".rtf")
RENDER_INPUT_SUFFIXES = CONVERT_INPUT_SUFFIXES + (".pdf",)
RECALC_INPUT_SUFFIXES = (".xlsx", ".ods")

DSOFFICE_ENV_VAR = "MCP_OFFICE_DSOFFICE"
# 兼容测试/外部包装器使用的别名；MCP_OFFICE_DSOFFICE 优先。
DSOFFICE_LEGACY_ENV_VAR = "MODULE_OFFICE_DSOFFICE"
BACKEND_ENV_VAR = "OFFICE_CONVERT_BACKEND"
BACKEND_MODES = ("auto", "dsoffice", "com", "none")
DEFAULT_DPI = 144
MIN_DPI, MAX_DPI = 72, 600

DSOFFICE_INSTALL_HINT = (
    "未找到 dsoffice CLI。请先安装：npm i -g dsoffice"
    "（要求 Node.js >= 22.19；Windows x64 需 MSVC v14 运行库），"
    "检测方法：在终端运行 which dsoffice（PowerShell 用 Get-Command dsoffice），"
    f"或通过环境变量 {DSOFFICE_ENV_VAR}（兼容 {DSOFFICE_LEGACY_ENV_VAR}）"
    "指定 dsoffice 可执行文件路径"
)

BACKEND_INSTALL_HINT = (
    "未找到可用的 Office 转换后端。可任选一种方式启用：\n"
    "1) dsoffice：npm i -g dsoffice（Node.js >= 22.19；Windows x64 需 MSVC v14 运行库），"
    "安装后运行 which dsoffice 检测，或设 MCP_OFFICE_DSOFFICE 指向可执行文件；\n"
    "2) Microsoft Office COM：Windows 安装 Microsoft Office（Word/Excel/PowerPoint）"
    "与 pywin32，并设 OFFICE_CONVERT_BACKEND=com（或 auto）；\n"
    "3) WPS Office COM：Windows 安装 WPS Office（KWps.Application/KET.Application/"
    "KWPP.Application）与 pywin32，并设 OFFICE_CONVERT_BACKEND=com（或 auto）。"
)

# 字体回退表：dsoffice/LibreOffice 报告缺失字体后，按类别给出本机实际可用的替代。
# 当前只覆盖等宽字体（代码块最常见的 Courier 缺失场景）。
MONOSPACE_FALLBACKS: tuple[str, ...] = (
    "Consolas",
    "Cascadia Mono",
    "DejaVu Sans Mono",
)
_MONOSPACE_TOKENS = (
    "courier", "consolas", "cascadia", "dejavusansmono",
    "lucidaconsole", "monaco", "menlo", "sourcecodepro", "firacode",
)
# Windows 字体文件名与注册表显示名不一致时的补充映射。
_FONT_FILE_NAMES = {
    "consola": "Consolas",
    "consolab": "Consolas",
    "consolai": "Consolas",
    "consolaz": "Consolas",
    "cascadiamono": "Cascadia Mono",
    "cascadiacode": "Cascadia Code",
    "dejavusansmono": "DejaVu Sans Mono",
    "lucon": "Lucida Console",
    "simsun": "SimSun",
    "simsunb": "SimSun-ExtB",
    "simsunextg": "SimSun-ExtG",
    "simhei": "SimHei",
    "simkai": "KaiTi",
    "simfang": "FangSong",
    "simli": "LiSu",
    "simyou": "YouYuan",
    "msyh": "Microsoft YaHei",
    "msyhbd": "Microsoft YaHei Bold",
    "msyhl": "Microsoft YaHei Light",
    "deng": "DengXian",
    "dengb": "DengXian Bold",
    "dengl": "DengXian Light",
    "stsong": "STSong",
    "stfangsong": "STFangsong",
    "stkaiti": "STKaiti",
    "stxihei": "STXihei",
    "stzhongs": "STZhongsong",
    "stxingka": "STXingkai",
}
# 中文 Word 常用字体名与 Windows 注册表英文名的对应关系。
_FONT_ALIASES = {
    "宋体": "SimSun",
    "新宋体": "NSimSun",
    "黑体": "SimHei",
    "微软雅黑": "Microsoft YaHei",
    "微软雅黑 Light": "Microsoft YaHei Light",
    "等线": "DengXian",
    "等线 Light": "DengXian Light",
    "楷体": "KaiTi",
    "楷体_GB2312": "KaiTi_GB2312",
    "仿宋": "FangSong",
    "仿宋_GB2312": "FangSong_GB2312",
    "隶书": "LiSu",
    "幼圆": "YouYuan",
    "华文宋体": "STSong",
    "华文仿宋": "STFangsong",
    "华文楷体": "STKaiti",
    "华文细黑": "STXihei",
    "华文中宋": "STZhongsong",
    "华文行楷": "STXingkai",
}
_INSTALLED_FONT_KEYS: set[str] | None = None


def _font_key(name: str) -> str:
    """Normalise a font name for comparison (case/spacing/parentheses insensitive)."""
    name = re.sub(r"\([^)]*\)", "", name)
    return re.sub(r"[^0-9a-z\u3400-\u9fff\uf900-\ufaff]+", "", name.lower())


_FONT_ALIAS_KEYS = {
    _font_key(chinese): _font_key(english)
    for chinese, english in _FONT_ALIASES.items()
}


def _split_registry_font_name(name: str) -> list[str]:
    """Split combined registry names such as ``SimSun & NSimSun (TrueType)``."""
    cleaned = re.sub(r"\([^)]*\)\s*$", "", name).strip()
    parts = re.split(r"\s*(?:&|,|，|、|\band\b)\s*", cleaned)
    result = [cleaned] if cleaned else []
    result.extend(part for part in parts if part and part not in result)
    return result


def _font_names_from_registry() -> set[str]:
    if os.name != "nt":
        return set()
    try:
        import winreg
    except ImportError:
        return set()
    names: set[str] = set()
    subkey = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(root, subkey) as key:
                for index in range(winreg.QueryInfoKey(key)[1]):
                    raw, value, _ = winreg.EnumValue(key, index)
                    for name in _split_registry_font_name(str(raw)):
                        names.add(name)
                    if value:
                        stem = Path(str(value)).stem
                        names.add(_FONT_FILE_NAMES.get(stem.lower(), stem))
        except OSError:
            continue
    return names


def _font_names_from_dirs() -> set[str]:
    dirs: list[Path] = []
    system_root = os.environ.get("WINDIR") or os.environ.get("SystemRoot")
    if system_root:
        dirs.append(Path(system_root) / "Fonts")
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        dirs.append(Path(local_app_data) / "Microsoft" / "Windows" / "Fonts")
    names: set[str] = set()
    for directory in dirs:
        if not directory.is_dir():
            continue
        try:
            entries = list(directory.iterdir())
        except OSError:
            continue
        for entry in entries:
            if entry.is_file() and entry.suffix.lower() in {".ttf", ".ttc", ".otf", ".tte"}:
                stem = entry.stem
                names.add(_FONT_FILE_NAMES.get(stem.lower(), stem))
    return names


def installed_font_keys() -> set[str]:
    """Return normalised names of fonts installed on this machine (cached)."""
    global _INSTALLED_FONT_KEYS
    if _INSTALLED_FONT_KEYS is None:
        names = _font_names_from_registry() | _font_names_from_dirs()
        keys = {_font_key(name) for name in names if _font_key(name)}
        for chinese_key, english_key in _FONT_ALIAS_KEYS.items():
            if chinese_key in keys or english_key in keys:
                keys.add(chinese_key)
                keys.add(english_key)
        _INSTALLED_FONT_KEYS = keys
    return _INSTALLED_FONT_KEYS


def _is_monospace_font(name: str) -> bool:
    key = _font_key(name)
    return any(token in key for token in _MONOSPACE_TOKENS)


def plan_font_substitutions(
    missing_fonts: list[str] | tuple[str, ...],
    available_fonts: set[str] | list[str] | tuple[str, ...] | None = None,
) -> list[dict[str, str]]:
    """Map missing monospace fonts to the first installed fallback.

    ``available_fonts`` is mainly for tests; when omitted the machine's font
    registry/directories are inspected. Unmatched fonts are left in
    ``missing_fonts`` so callers still see the original warning.
    """
    available = (
        installed_font_keys()
        if available_fonts is None
        else {_font_key(name) for name in available_fonts}
    )
    substitutions: list[dict[str, str]] = []
    seen: set[str] = set()
    for missing in missing_fonts:
        if not isinstance(missing, str) or not missing:
            continue
        if missing in seen:
            continue
        seen.add(missing)
        if not _is_monospace_font(missing):
            continue
        for candidate in MONOSPACE_FALLBACKS:
            if _font_key(candidate) in available:
                substitutions.append({"missing": missing, "substitute": candidate})
                break
    return substitutions


def find_dsoffice() -> str:
    """定位 dsoffice 可执行文件，找不到时抛出带安装指引的 RuntimeError。"""
    env_name = DSOFFICE_ENV_VAR
    custom = os.environ.get(DSOFFICE_ENV_VAR)
    if not custom:
        env_name = DSOFFICE_LEGACY_ENV_VAR
        custom = os.environ.get(DSOFFICE_LEGACY_ENV_VAR)
    if custom:
        if Path(custom).is_file():
            return custom
        raise RuntimeError(f"{env_name} 指定的路径不存在: {custom}")
    resolved = shutil.which("dsoffice")
    if resolved:
        return resolved
    raise RuntimeError(DSOFFICE_INSTALL_HINT)


def backend_mode() -> str:
    """Read and validate ``OFFICE_CONVERT_BACKEND``."""
    raw = (os.environ.get(BACKEND_ENV_VAR) or "auto").strip().lower()
    if raw not in BACKEND_MODES:
        raise RuntimeError(
            f"{BACKEND_ENV_VAR} 必须是 {'/'.join(BACKEND_MODES)} 之一，收到 {raw!r}"
        )
    return raw


def _com_is_available() -> bool:
    from .office_com import com_is_available

    return com_is_available()


def _backend_unavailable(operation: str, dsoffice_error: str | None = None) -> RuntimeError:
    message = f"{operation}失败：{BACKEND_INSTALL_HINT}"
    if dsoffice_error:
        message += f"\ndsoffice 检测结果：{dsoffice_error}"
    return RuntimeError(message)


def _resolve_dsoffice(mode: str) -> tuple[str | None, str | None]:
    if mode not in ("auto", "dsoffice"):
        return None, None
    try:
        return find_dsoffice(), None
    except RuntimeError as exc:
        if mode == "dsoffice":
            raise
        return None, str(exc)


def _run_dsoffice(cli_path: str, args: list[str], timeout: float) -> dict:
    """执行 dsoffice 子命令并解析其单行 JSON 输出。

    Returns:
        CLI 返回的 JSON dict（成功时含 outputPath/outputDir 等字段）。

    Raises:
        RuntimeError: 进程失败、超时或输出不是合法 JSON。
    """
    cmd = [cli_path, *args]
    # npm 在 Windows 上安装的是 dsoffice.cmd shim；subprocess 直接启动
    # .cmd 会挂起，必须经 cmd.exe /c 转发。
    if os.name == "nt" and Path(cli_path).suffix.lower() in {".cmd", ".bat"}:
        cmd = [os.environ.get("COMSPEC", "cmd.exe"), "/d", "/c", *cmd]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="backslashreplace",
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"dsoffice 执行超时（>{timeout}s）: {' '.join(cmd)}") from e
    except OSError as e:
        raise RuntimeError(f"dsoffice 无法启动: {e}") from e

    stdout = (proc.stdout or "").strip()
    try:
        payload = json.loads(stdout) if stdout else None
    except ValueError as e:
        stderr = (proc.stderr or "").strip()
        raise RuntimeError(
            f"dsoffice 输出无法解析（exit={proc.returncode}）: {stdout[:300] or stderr[:300]}"
        ) from e
    if not isinstance(payload, dict):
        # 刻意转成 RuntimeError：调用方只关心"CLI 输出异常"一种失败形态
        raise RuntimeError(  # noqa: TRY004
            f"dsoffice 输出不是 JSON 对象（exit={proc.returncode}）: {stdout[:300]}"
        )

    if proc.returncode != 0 or payload.get("code") == "failed":
        raise RuntimeError(f"dsoffice 转换失败: {payload.get('error') or payload}")
    return payload


def _resolve_target(path: Path, if_exists: str) -> Path:
    """Resolve if_exists semantics without deleting an existing target."""
    return Path(resolve_if_exists(str(path), if_exists))


def _temp_file_near(target: Path, prefix: str) -> Path:
    """Create a unique, currently non-existent file path next to ``target``."""
    fd, name = tempfile.mkstemp(prefix=prefix, suffix=target.suffix, dir=target.parent)
    os.close(fd)
    os.unlink(name)
    return Path(name)


def _aside_path(target: Path) -> Path:
    """Pick a non-existent ``*.bak-YYYYmmdd-HHMMSS`` sibling for the old target."""
    stamp = time.strftime("%Y%m%d-%H%M%S")
    candidate = target.with_name(f"{target.name}.bak-{stamp}")
    counter = 1
    while candidate.exists():
        candidate = target.with_name(f"{target.name}.bak-{stamp}-{counter}")
        counter += 1
    return candidate


def _replace_file(temp: Path, target: Path) -> None:
    """Atomically replace a file target; keep the old file as a timestamped backup."""
    if target.is_dir():
        os.replace(target, _aside_path(target))
    elif target.is_file() and target.stat().st_size > 0:
        backup_target(str(target))
    os.replace(temp, target)


def _replace_dir(temp: Path, target: Path) -> None:
    """Replace a directory target; rename the old one aside instead of deleting it."""
    if target.exists() or target.is_symlink():
        os.replace(target, _aside_path(target))
    os.replace(temp, target)


def _validate_pages(pages: str) -> str:
    """校验 --pages 参数：'all' 或 '1,3' 样式的页码列表。"""
    pages = (pages or "all").strip()
    if pages.lower() == "all":
        return "all"
    if not re.fullmatch(r"\d+(?:\s*,\s*\d+)*", pages):
        raise ValueError(f"pages 参数无效: {pages!r}，应为 'all' 或 '1,3,5' 样式")
    return ",".join(p.strip() for p in pages.split(","))


# ── COM 后端：Microsoft Office / WPS ─────────────────────────────

COM_KIND_BY_SUFFIX = {
    ".docx": "word",
    ".doc": "word",
    ".rtf": "word",
    ".xlsx": "excel",
    ".xls": "excel",
    ".pptx": "ppt",
    ".ppt": "ppt",
}
COM_OFFICE_PROGIDS = {
    "word": "Word.Application",
    "excel": "Excel.Application",
    "ppt": "PowerPoint.Application",
}
COM_WPS_PROGIDS = {
    "word": "KWps.Application",
    "excel": "KET.Application",
    "ppt": "KWPP.Application",
}
_WORD_EXPORT_PDF = 17
_EXCEL_EXPORT_PDF = 0
_PPT_SAVE_AS_PDF = 32


def _com_candidates(kind: str) -> list[tuple[str, str]]:
    """Office first, then WPS; WPS may register the standard ProgID itself."""
    office = COM_OFFICE_PROGIDS[kind]
    wps = COM_WPS_PROGIDS[kind]
    candidates = [("office", office), ("wps", wps)]
    if wps.lower() != office.lower():
        candidates.append(("wps-fallback", office))
    return candidates


def _com_app_info(app) -> dict[str, str]:
    info: dict[str, str] = {}
    for attr in ("Name", "Version", "Path", "Build"):
        try:
            info[attr.lower()] = str(getattr(app, attr))
        except Exception:
            continue
    return info


def _window_process_id(window) -> int | None:
    try:
        import win32process

        hwnd = int(getattr(window, "Hwnd"))
        if hwnd:
            return int(win32process.GetWindowThreadProcessId(hwnd)[1])
    except Exception:
        pass
    return None


def _document_process_id(document) -> int | None:
    if document is None:
        return None
    try:
        return _window_process_id(document.ActiveWindow)
    except Exception:
        pass
    try:
        return _window_process_id(document.Windows(1))
    except Exception:
        return None


def _com_process_id(app) -> int | None:
    for attr in ("ProcessID", "ProcessId"):
        try:
            value = int(getattr(app, attr))
            if value > 0:
                return value
        except Exception:
            continue
    try:
        import win32process

        hwnd = int(getattr(app, "Hwnd"))
        if hwnd:
            return int(win32process.GetWindowThreadProcessId(hwnd)[1])
    except Exception:
        pass
    try:
        return _window_process_id(app.ActiveWindow)
    except Exception:
        return None


def _pid_alive(pid: int) -> bool:
    if os.name != "nt":
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False
    process_query_limited_information = 0x1000
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(process_query_limited_information, False, int(pid))
    if not handle:
        return False
    try:
        exit_code = ctypes.c_ulong()
        if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
            return False
        return exit_code.value == 259  # STILL_ACTIVE
    finally:
        kernel32.CloseHandle(handle)


def _terminate_pid(pid: int) -> bool:
    if os.name != "nt":
        return False
    try:
        proc = subprocess.run(
            ["taskkill", "/PID", str(pid), "/F", "/T"],
            capture_output=True,
            timeout=15,
            check=False,
        )
        return proc.returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False


def _quit_com_app(app, pid: int | None) -> str:
    """Quit a COM application and wait; force-terminate only as a last resort."""
    cleanup = "quit"
    try:
        app.Quit()
    except Exception as exc:  # noqa: BLE001 - cleanup must not mask the real result
        cleanup = f"quit_error: {exc}"
    if pid:
        deadline = time.time() + 5.0
        while time.time() < deadline:
            if not _pid_alive(pid):
                return cleanup
            time.sleep(0.2)
        if _terminate_pid(pid):
            cleanup += "; terminated"
    return cleanup


def _close_quietly(obj, *args) -> None:
    if obj is None:
        return
    try:
        obj.Close(*args)
    except Exception:
        pass


def _export_with_com(
    kind: str,
    progid: str,
    src: Path,
    target: Path,
    sheet: str | None,
    cell_range: str | None,
) -> dict:
    """DispatchEx an Office/WPS app, export one document to PDF, then quit."""
    import win32com.client

    app = None
    document = None
    pid = None
    info: dict[str, object] = {"prog_id": progid, "kind": kind}
    try:
        app = win32com.client.DispatchEx(progid)
        pid = _com_process_id(app)
        info.update(_com_app_info(app))
        if kind == "word":
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            document = app.Documents.Open(
                str(src), ReadOnly=True, AddToRecentFiles=False
            )
            pid = pid or _com_process_id(app) or _document_process_id(document)
            document.ExportAsFixedFormat(str(target), _WORD_EXPORT_PDF)
        elif kind == "excel":
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.DisplayAlerts = False
            except Exception:
                pass
            document = app.Workbooks.Open(str(src), ReadOnly=True, UpdateLinks=0)
            pid = pid or _com_process_id(app) or _document_process_id(document)
            if sheet:
                worksheet = document.Worksheets(sheet)
                if cell_range:
                    worksheet.PageSetup.PrintArea = cell_range.replace("$", "")
                worksheet.ExportAsFixedFormat(_EXCEL_EXPORT_PDF, str(target))
            else:
                document.ExportAsFixedFormat(_EXCEL_EXPORT_PDF, str(target))
        else:
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            document = app.Presentations.Open(
                str(src), ReadOnly=True, WithWindow=False
            )
            pid = pid or _com_process_id(app) or _document_process_id(document)
            document.SaveAs(str(target), _PPT_SAVE_AS_PDF)
    finally:
        if kind == "word":
            _close_quietly(document, False)
        elif kind == "excel":
            _close_quietly(document, False)
        else:
            _close_quietly(document)
        if app is not None:
            info["cleanup"] = _quit_com_app(app, pid)

    if not target.is_file() or target.stat().st_size <= 0:
        raise RuntimeError(f"{progid} 未生成有效 PDF: {target}")
    info["engine"] = "wps" if "wps" in str(info.get("path", "")).lower() else "office"
    return info


def _convert_with_com(
    src: Path,
    target: Path,
    sheet: str | None = None,
    cell_range: str | None = None,
) -> dict:
    """Convert an Office file through the Office/WPS COM fallback chain."""
    if not _com_is_available():
        raise RuntimeError(BACKEND_INSTALL_HINT)
    kind = COM_KIND_BY_SUFFIX.get(src.suffix.lower())
    if kind is None:
        raise ValueError(f"COM 转换不支持的文件格式: {src.suffix}")

    import pythoncom

    errors: list[str] = []
    pythoncom.CoInitialize()
    try:
        for stage, progid in _com_candidates(kind):
            attempt = _temp_file_near(target, f".com-{stage}-")
            try:
                info = _export_with_com(kind, progid, src, attempt, sheet, cell_range)
                info["stage"] = stage
                _replace_file(attempt, target)
                return info
            except Exception as exc:  # noqa: BLE001 - try the next COM provider
                errors.append(f"{stage}/{progid}: {exc}")
                if attempt.exists():
                    attempt.unlink()
    finally:
        pythoncom.CoUninitialize()
    raise RuntimeError("COM 转换失败（已尝试 Microsoft Office 与 WPS）：" + "；".join(errors))


def _recalculate_with_com(src: Path, target: Path) -> dict:
    """Recalculate an xlsx through Excel/WPS COM and write cached values."""
    if not _com_is_available():
        raise RuntimeError(BACKEND_INSTALL_HINT)
    if src.suffix.lower() != ".xlsx" or target.suffix.lower() != ".xlsx":
        raise RuntimeError("COM 公式重算回退只支持 .xlsx 输入与输出（.ods 请安装 dsoffice）")

    import win32com.client
    import pythoncom

    errors: list[str] = []
    pythoncom.CoInitialize()
    try:
        for stage, progid in _com_candidates("excel"):
            attempt = _temp_file_near(target, f".recalc-{stage}-")
            shutil.copy2(src, attempt)
            app = None
            workbook = None
            pid = None
            info: dict[str, object] = {"prog_id": progid, "kind": "excel"}
            try:
                app = win32com.client.DispatchEx(progid)
                pid = _com_process_id(app)
                info.update(_com_app_info(app))
                try:
                    app.Visible = False
                except Exception:
                    pass
                try:
                    app.DisplayAlerts = False
                except Exception:
                    pass
                workbook = app.Workbooks.Open(str(attempt))
                pid = pid or _com_process_id(app) or _document_process_id(workbook)
                try:
                    app.CalculateFullRebuild()
                    info["calculation"] = "CalculateFullRebuild"
                except Exception:
                    app.Calculate()
                    info["calculation"] = "Calculate"
                workbook.Save()
            except Exception as exc:  # noqa: BLE001 - try the next COM provider
                errors.append(f"{stage}/{progid}: {exc}")
                if attempt.exists():
                    attempt.unlink()
                continue
            finally:
                _close_quietly(workbook, False)
                if app is not None:
                    info["cleanup"] = _quit_com_app(app, pid)
            try:
                if not attempt.is_file() or attempt.stat().st_size <= 0:
                    raise RuntimeError(f"{progid} 未生成有效工作簿")
                info["stage"] = stage
                info["engine"] = (
                    "wps" if "wps" in str(info.get("path", "")).lower() else "office"
                )
                _replace_file(attempt, target)
                return info
            except Exception as exc:  # noqa: BLE001 - try the next COM provider
                errors.append(f"{stage}/{progid}: {exc}")
                if attempt.exists():
                    attempt.unlink()
    finally:
        pythoncom.CoUninitialize()
    raise RuntimeError("COM 公式重算失败（已尝试 Microsoft Excel 与 WPS）：" + "；".join(errors))


# ── PyMuPDF 渲染 ─────────────────────────────────────────────────

def _parse_page_numbers(pages: str, total: int) -> list[int]:
    pages = _validate_pages(pages)
    if pages.lower() == "all":
        return list(range(1, total + 1))
    numbers: list[int] = []
    for item in pages.split(","):
        number = int(item)
        if number < 1 or number > total:
            raise RuntimeError(f"页码 {number} 超出范围（共 {total} 页）")
        if number not in numbers:
            numbers.append(number)
    return numbers


def _remap_image_paths(images: list, target_dir: Path) -> list:
    """Point image paths at the published output directory, not the temp dir."""
    remapped: list = []
    for item in images:
        if not isinstance(item, dict):
            remapped.append(item)
            continue
        updated = dict(item)
        path = updated.get("path")
        if path:
            updated["path"] = str(target_dir / Path(str(path)).name)
        remapped.append(updated)
    return remapped


def _render_pdf_with_pymupdf(
    pdf_path: Path,
    output_dir: Path,
    pages: str,
    dpi: int,
) -> dict:
    """Render PDF pages to PNG with PyMuPDF, honouring pages/dpi semantics."""
    try:
        import pymupdf
    except ImportError:  # older PyMuPDF exposes only the fitz alias
        import fitz as pymupdf

    zoom = max(MIN_DPI, min(MAX_DPI, int(dpi))) / 72.0
    document = pymupdf.open(str(pdf_path))
    try:
        total = document.page_count
        selected = _parse_page_numbers(pages, total)
        output_dir.mkdir(parents=True, exist_ok=True)
        images: list[dict] = []
        matrix = pymupdf.Matrix(zoom, zoom)
        for index, page_number in enumerate(selected, start=1):
            page = document.load_page(page_number - 1)
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            image_path = output_dir / f"page-{page_number:04d}.png"
            pixmap.save(str(image_path))
            images.append({
                "index": index,
                "page": page_number,
                "width": pixmap.width,
                "height": pixmap.height,
                "path": str(image_path),
            })
        return {
            "page_count": len(images),
            "page_count_reported": total,
            "images": images,
            "missing_fonts": [],
        }
    finally:
        document.close()


def _render_with_dsoffice(
    cli_path: str,
    src: Path,
    target: Path,
    pages: str,
    dpi: int,
    sheet: str | None,
    cell_range: str | None,
) -> dict:
    """Run dsoffice render into a temporary directory, then publish it."""
    temp_dir = Path(tempfile.mkdtemp(prefix=f".{target.name}-", dir=target.parent))
    temp_dir.rmdir()  # dsoffice requires a non-existent output directory
    args = [
        "render",
        "--input", str(src.resolve()),
        "--output-dir", str(temp_dir.resolve()),
        "--dpi", str(max(MIN_DPI, min(MAX_DPI, int(dpi)))),
    ]
    if src.suffix.lower() in {".xlsx", ".xls"}:
        if not (sheet and cell_range):
            raise ValueError("Excel 渲染需要同时指定 sheet 和 cell_range（如 --sheet 数据 --cell-range A1:D20）")
        args += ["--sheet", sheet, "--range", cell_range]
    else:
        args += ["--pages", _validate_pages(pages)]
        if sheet:
            args += ["--sheet", sheet]
        if cell_range:
            if not sheet:
                raise ValueError("cell_range 需要与 sheet 同时指定")
            args += ["--range", cell_range]
    try:
        payload = _run_dsoffice(cli_path, args, timeout=300)
        if not temp_dir.is_dir():
            raise RuntimeError(f"dsoffice 未生成输出目录: {temp_dir}")
        _replace_dir(temp_dir, target)
        payload["images"] = _remap_image_paths(payload.get("images") or [], target)
        return payload
    finally:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)


# ── 公开能力 ─────────────────────────────────────────────────────

def convert_to_pdf(
    input_path: str,
    output_path: str | None = None,
    if_exists: str = "overwrite",
    sheet: str | None = None,
) -> dict:
    """将 Office 文档转换为 PDF。

    支持 .docx/.xlsx/.pptx 及 97-2003 的 .doc/.xls/.ppt、.rtf。
    后端顺序由 OFFICE_CONVERT_BACKEND 控制：auto 先 dsoffice，失败/缺失回退
    Microsoft Office COM 再 WPS COM；com 强制 COM；dsoffice 强制 CLI；
    none 禁用。Excel 可用 sheet 指定单个工作表（COM 回退同时支持 cell_range，
    但本入口不暴露）。
    返回 success/output_path/backend/missing_fonts/font_substitutions。
    """
    validate_input_file(input_path, CONVERT_INPUT_SUFFIXES)
    mode = backend_mode()
    if mode == "none":
        raise _backend_unavailable("PDF 转换")
    dsoffice_path, dsoffice_error = _resolve_dsoffice(mode)
    if not dsoffice_path and not _com_is_available():
        raise _backend_unavailable("PDF 转换", dsoffice_error)

    src = Path(input_path)
    requested = Path(output_path) if output_path else src.with_suffix(".pdf")
    validate_output_path(str(requested), (".pdf",))
    target = _resolve_target(requested, if_exists)
    ensure_output_dir(str(target))

    fallback: list[dict[str, str]] = []
    backend = "dsoffice" if dsoffice_path else "com"
    payload: dict = {}
    if backend == "dsoffice":
        temp = _temp_file_near(target, ".dsoffice-")
        args = [
            "convert",
            "--input", str(src.resolve()),
            "--output", str(temp.resolve()),
        ]
        if sheet:
            args += ["--sheet", sheet]
        try:
            payload = _run_dsoffice(dsoffice_path, args, timeout=300)
            if temp.is_file() and temp.stat().st_size > 0:
                _replace_file(temp, target)
        except RuntimeError as exc:
            if temp.exists():
                temp.unlink()
            if mode == "auto" and _com_is_available():
                fallback.append({"backend": "dsoffice", "error": str(exc)})
                backend = "com"
            else:
                raise

    if backend == "com":
        if not _com_is_available():
            raise _backend_unavailable("PDF 转换", dsoffice_error)
        if not dsoffice_path and mode == "auto":
            fallback.append({
                "backend": "dsoffice",
                "error": dsoffice_error or DSOFFICE_INSTALL_HINT,
            })
        info = _convert_with_com(src, target, sheet=sheet)
        result = {
            "success": True,
            "file_path": str(target),
            "backend": "com",
            "backend_kind": "com",
            "com_engine": info.get("engine", info.get("stage", "com")),
            "com_prog_id": info.get("prog_id"),
            "com_application": info.get("name") or info.get("path"),
            "missing_fonts": [],
            "font_substitutions": [],
            "message": f"转换成功（COM: {info.get('prog_id')}）",
        }
        if fallback:
            result["fallback_from"] = fallback
        return result

    missing_fonts = payload.get("missingFonts") or []
    if not isinstance(missing_fonts, list):
        missing_fonts = [str(missing_fonts)]
    substitutions = plan_font_substitutions(missing_fonts)
    message = "转换成功"
    if missing_fonts:
        message += f"；缺失字体（已按字形度量替换）: {', '.join(missing_fonts)}"
        if substitutions:
            message += "；可用回退: " + ", ".join(
                f"{item['missing']}→{item['substitute']}" for item in substitutions
            )
    result = {
        "success": True,
        "file_path": str(target),
        "backend": payload.get("backend"),
        "backend_kind": "dsoffice",
        "missing_fonts": missing_fonts,
        "font_substitutions": substitutions,
        "message": message,
    }
    if fallback:
        result["fallback_from"] = fallback
    return result


def render_document_pages(
    input_path: str,
    output_dir: str | None = None,
    pages: str = "all",
    dpi: int = DEFAULT_DPI,
    sheet: str | None = None,
    cell_range: str | None = None,
    if_exists: str = "overwrite",
) -> dict:
    """将 Office/PDF 文档逐页渲染为 PNG 图片。

    支持 convert_to_pdf 的全部输入格式及 .pdf。pages 为 'all' 或 '1,3,5'；
    Excel 可用 sheet + cell_range（如 A1:D20）只渲染指定区域。
    dpi 默认 144（72-600）。输出目录必须为空/不存在：overwrite 先清空重建，error 直接报错。

    dsoffice 缺失或执行失败时：PDF 直接由 PyMuPDF 渲染；Office 文档先用 COM
    转成临时 PDF，再逐页渲染。返回 page_count（实际 PNG 数）、images 与后端字段。
    """
    validate_input_file(input_path, RENDER_INPUT_SUFFIXES)
    src = Path(input_path)
    is_pdf = src.suffix.lower() == ".pdf"
    mode = backend_mode()
    if mode == "none" and not is_pdf:
        raise _backend_unavailable("页面渲染")

    requested = Path(output_dir) if output_dir else src.parent / f"{src.stem}_pages"
    target = _resolve_target(requested, if_exists)
    target.parent.mkdir(parents=True, exist_ok=True)
    dsoffice_path, dsoffice_error = _resolve_dsoffice(mode)

    fallback: list[dict[str, str]] = []
    if not dsoffice_path and mode == "auto" and dsoffice_error:
        fallback.append({"backend": "dsoffice", "error": dsoffice_error})
    if dsoffice_path and mode in ("auto", "dsoffice"):
        try:
            payload = _render_with_dsoffice(
                dsoffice_path, src, target, pages, dpi, sheet, cell_range
            )
            return {
                "success": True,
                "output_dir": str(target),
                "page_count": payload.get("pageCount"),
                "images": payload.get("images") or [],
                "missing_fonts": payload.get("missingFonts") or [],
                "backend": "dsoffice",
                "backend_kind": "dsoffice",
                "message": f"渲染完成，共 {payload.get('pageCount')} 页",
            }
        except Exception as exc:  # noqa: BLE001 - auto may fall back to PyMuPDF
            if mode != "auto":
                raise
            fallback.append({"backend": "dsoffice", "error": str(exc)})

    if not is_pdf and not _com_is_available():
        raise _backend_unavailable("页面渲染", dsoffice_error)

    work_dir = Path(tempfile.mkdtemp(prefix=f".{target.name}-", dir=target.parent))
    pages_dir = work_dir / "pages"
    try:
        conversion_info: dict | None = None
        if is_pdf:
            pdf_path = src
            backend = "pymupdf"
            backend_kind = "pymupdf"
        else:
            pdf_path = work_dir / f"{src.stem}.pdf"
            conversion_info = _convert_with_com(
                src, pdf_path, sheet=sheet, cell_range=cell_range
            )
            backend = "com+pymupdf"
            backend_kind = "com"
        rendered = _render_pdf_with_pymupdf(pdf_path, pages_dir, pages, dpi)
        _replace_dir(pages_dir, target)
        rendered["images"] = _remap_image_paths(rendered["images"], target)
    finally:
        if work_dir.exists():
            shutil.rmtree(work_dir)

    result = {
        "success": True,
        "output_dir": str(target),
        "page_count": rendered["page_count"],
        "images": rendered["images"],
        "missing_fonts": [],
        "backend": backend,
        "backend_kind": backend_kind,
        "message": (
            f"渲染完成，共 {rendered['page_count']} 页"
            f"（{backend}；PyMuPDF 渲染）"
        ),
    }
    if rendered["page_count_reported"] != rendered["page_count"]:
        result["page_count_reported"] = rendered["page_count_reported"]
    if conversion_info:
        result["com_prog_id"] = conversion_info.get("prog_id")
        result["com_engine"] = conversion_info.get("engine")
    if fallback:
        result["fallback_from"] = fallback
    return result


def recalculate_excel_formulas(
    input_path: str,
    output_path: str | None = None,
    if_exists: str = "overwrite",
) -> dict:
    """重算 xlsx 公式并把结果写为缓存值。

    openpyxl 写入的公式没有缓存值，多数读取工具（含 read_excel_workbook 的 data_only 模式）
    看到的都是空；转换后的文件公式与数值并存。默认输出 <原名>_recalc.xlsx，也可指定 .ods。

    后端顺序同 convert：auto 先 dsoffice，缺失/失败回退 Excel COM；COM 回退用
    CalculateFullRebuild + Save，只支持 .xlsx。dsoffice 结果保持原有 backend 字段。
    """
    validate_input_file(input_path, RECALC_INPUT_SUFFIXES)
    mode = backend_mode()
    if mode == "none":
        raise _backend_unavailable("公式重算")
    dsoffice_path, dsoffice_error = _resolve_dsoffice(mode)
    if not dsoffice_path and not _com_is_available():
        raise _backend_unavailable("公式重算", dsoffice_error)

    src = Path(input_path)
    default_target = src.with_name(f"{src.stem}_recalc{src.suffix}")
    requested = Path(output_path) if output_path else default_target
    validate_output_path(str(requested), (".xlsx", ".ods"))
    target = _resolve_target(requested, if_exists)
    ensure_output_dir(str(target))

    fallback: list[dict[str, str]] = []
    backend = "dsoffice" if dsoffice_path else "com"
    payload: dict = {}
    if backend == "dsoffice":
        temp = _temp_file_near(target, ".dsoffice-recalc-")
        try:
            payload = _run_dsoffice(
                dsoffice_path,
                ["recalculate", "--input", str(src.resolve()), "--output", str(temp.resolve())],
                timeout=300,
            )
            if temp.is_file() and temp.stat().st_size > 0:
                _replace_file(temp, target)
        except RuntimeError as exc:
            if temp.exists():
                temp.unlink()
            if mode == "auto" and _com_is_available():
                fallback.append({"backend": "dsoffice", "error": str(exc)})
                backend = "com"
            else:
                raise

    if backend == "com":
        if not _com_is_available():
            raise _backend_unavailable("公式重算", dsoffice_error)
        if not dsoffice_path and mode == "auto":
            fallback.append({
                "backend": "dsoffice",
                "error": dsoffice_error or DSOFFICE_INSTALL_HINT,
            })
        info = _recalculate_with_com(src, target)
        result = {
            "success": True,
            "file_path": str(target),
            "backend": "com",
            "backend_kind": "com",
            "com_engine": info.get("engine", info.get("stage", "com")),
            "com_prog_id": info.get("prog_id"),
            "com_application": info.get("name") or info.get("path"),
            "calculation": info.get("calculation"),
            "missing_fonts": [],
            "message": (
                f"公式重算完成（COM: {info.get('prog_id')}，"
                f"{info.get('calculation')}），缓存值已写入"
            ),
        }
        if fallback:
            result["fallback_from"] = fallback
        return result

    result = {
        "success": True,
        "file_path": str(target),
        "backend": payload.get("backend"),
        "backend_kind": "dsoffice",
        "missing_fonts": payload.get("missingFonts") or [],
        "message": "公式重算完成，缓存值已写入",
    }
    if fallback:
        result["fallback_from"] = fallback
    return result
