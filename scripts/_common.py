"""Shared CLI contract for the office-writer skill entry scripts.

Every entry script is a thin, stateless wrapper around the vendored
``office_engine`` package.  Paths are resolved from ``__file__`` only; the
caller supplies all document paths as arguments.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from pydantic import BaseModel, ValidationError

from office_engine.guards import with_timeout

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
ENGINE_DIR = SCRIPT_DIR / "office_engine"
CONFIG_PATH = SKILL_DIR / "config" / "defaults.json"
TEMPLATE_DIR = SKILL_DIR / "templates"
ASSETS_DIR = SKILL_DIR / "assets"
REFERENCES_DIR = SKILL_DIR / "references"

_CONFIG_CACHE: dict[str, Any] | None = None


def load_config() -> dict[str, Any]:
    """Read config/defaults.json once; a missing or invalid file yields {}."""
    global _CONFIG_CACHE
    if _CONFIG_CACHE is None:
        try:
            value = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
            _CONFIG_CACHE = value if isinstance(value, dict) else {}
        except Exception:
            _CONFIG_CACHE = {}
    return _CONFIG_CACHE


def config_get(path: str, default: Any) -> Any:
    """Read a dotted key from config/defaults.json."""
    current: Any = load_config()
    for part in path.split("."):
        if not isinstance(current, dict) or part not in current:
            return default
        current = current[part]
    return current


COM_OPERATION_TIMEOUT = int(config_get("timeouts.com_operation_seconds", 60))
MARKDOWN_CONVERT_TIMEOUT = int(config_get("timeouts.markdown_convert_seconds", 120))
DSOFFICE_OPERATION_TIMEOUT = int(config_get("timeouts.dsoffice_operation_seconds", 300))
DEFAULT_RENDER_DPI = int(config_get("render.default_dpi", 144))

class JsonArgumentParser(argparse.ArgumentParser):
    """Argparse that keeps the stdout JSON contract on usage errors."""

    tool = "office-writer"

    def error(self, message: str):  # type: ignore[override]
        self.print_usage(sys.stderr)
        print(json.dumps({"ok": False, "tool": self.tool, "error": message}, ensure_ascii=False, indent=2))
        raise SystemExit(2)


def json_parser(tool: str, description: str, **kwargs) -> JsonArgumentParser:
    """Build the shared parser shape: --params / --params-file plus tool flags."""
    parser = JsonArgumentParser(
        prog=tool,
        description=description,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        **kwargs,
    )
    parser.tool = tool
    parser.add_argument("--params", help="完整工具参数 JSON 对象；同名显式 flag 优先")
    parser.add_argument("--params-file", help="包含完整工具参数 JSON 对象的文件；- 表示 stdin")
    return parser


class UsageError(ValueError):
    """Bad arguments or invalid input (exit code 2)."""


class ToolRuntimeError(RuntimeError):
    """Missing asset, unreadable document, or failed external call (exit code 3)."""


def configure_stdio() -> None:
    """Force UTF-8 on Windows consoles and pipes."""
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass


def emit(tool: str, data: dict[str, Any], exit_code: int = 0) -> int:
    """Print one JSON object to stdout and return the process exit code."""
    print(json.dumps({"ok": True, "tool": tool, "data": data}, ensure_ascii=False, indent=2))
    return exit_code


def emit_error(tool: str, message: str, exit_code: int) -> int:
    """Print one JSON error object to stdout and return the process exit code."""
    print(json.dumps({"ok": False, "tool": tool, "error": message}, ensure_ascii=False, indent=2))
    return exit_code


def read_text(text: str | None, input_file: str | None, label: str = "text") -> str:
    """Resolve text from --text, --input-file PATH, or --input-file - (stdin)."""
    if text is not None:
        return text
    if input_file == "-":
        raw = sys.stdin.buffer.read()
        for encoding in ("utf-8-sig", "utf-8", "gbk"):
            try:
                return raw.decode(encoding)
            except UnicodeDecodeError:
                continue
        raise UsageError("stdin 内容既不是 UTF-8 也不是 GBK，请改用 --input-file")
    if not input_file:
        raise UsageError(f"{label} 不能为空：请传 --text 或 --input-file")
    path = Path(input_file)
    if not path.is_file():
        raise UsageError(f"输入文件不存在：{path}")
    try:
        return path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as exc:
        raise UsageError(f"输入文件不是 UTF-8：{path}（{exc}）") from exc


def parse_params(args: argparse.Namespace) -> dict[str, Any]:
    """Load --params-file then --params; inline values override file values."""
    params: dict[str, Any] = {}
    params_file = getattr(args, "params_file", None)
    if params_file:
        params.update(_parse_json_object(read_text(None, params_file, "params-file"), "--params-file"))
    inline = getattr(args, "params", None)
    if inline:
        params.update(_parse_json_object(inline, "--params"))
    return params


def _parse_json_object(raw: str, label: str) -> dict[str, Any]:
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise UsageError(f"{label} 不是合法 JSON：{exc}") from exc
    if not isinstance(value, dict):
        raise UsageError(f"{label} 必须是 JSON 对象（以 {{ 开头）")
    return value


def arg(
    args: argparse.Namespace,
    params: dict[str, Any],
    name: str,
    default: Any = None,
    required: bool = False,
) -> Any:
    """Resolve one argument: explicit flag > --params key > default."""
    if hasattr(args, name):
        value = getattr(args, name)
        if value is not None:
            return value
    if name in params and params[name] is not None:
        return params[name]
    if required:
        flag = "--" + name.replace("_", "-")
        raise UsageError(f"缺少参数 {flag}（或 --params 中的 {name}）")
    return default


def int_arg(
    args: argparse.Namespace,
    params: dict[str, Any],
    name: str,
    default: int | None = None,
    *,
    required: bool = False,
) -> int:
    value = arg(args, params, name, default=default, required=required)
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise UsageError(f"{name} 必须是整数，收到 {value!r}") from exc


def bool_arg(
    args: argparse.Namespace,
    params: dict[str, Any],
    name: str,
    default: bool = False,
) -> bool:
    value = arg(args, params, name, default=default)
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and value in (0, 1):
        return bool(value)
    text = str(value).strip().lower()
    if text in {"true", "1", "yes", "on"}:
        return True
    if text in {"false", "0", "no", "off", ""}:
        return False
    raise UsageError(f"{name} 必须是布尔值，收到 {value!r}")


def model_arg(
    params: dict[str, Any],
    name: str,
    model: type[BaseModel],
    *,
    required: bool = True,
) -> Any:
    """Validate a complex argument with the same Pydantic model FastMCP used."""
    value = params.get(name)
    if value is None:
        if required:
            raise UsageError(f"缺少参数 {name}（通过 --params 或 --params-file 传入 JSON 对象）")
        return None
    if isinstance(value, model):
        return value
    try:
        return model.model_validate(value)
    except ValidationError as exc:
        raise UsageError(f"{name} 参数校验失败：{exc.errors(include_url=False)}") from exc


def model_list_arg(
    params: dict[str, Any],
    name: str,
    model: type[BaseModel],
    *,
    required: bool = True,
) -> list[Any]:
    """Validate a list of complex arguments item by item, like FastMCP did."""
    value = params.get(name)
    if value is None:
        if required:
            raise UsageError(f"缺少参数 {name}（通过 --params 或 --params-file 传入 JSON 数组）")
        return []
    if not isinstance(value, list):
        raise UsageError(f"{name} 必须是 JSON 数组")
    parsed: list[Any] = []
    for index, item in enumerate(value):
        if isinstance(item, model):
            parsed.append(item)
            continue
        try:
            parsed.append(model.model_validate(item))
        except ValidationError as exc:
            raise UsageError(f"{name}[{index}] 参数校验失败：{exc.errors(include_url=False)}") from exc
    return parsed


def choice(value: Any, allowed: set[str] | tuple[str, ...], name: str) -> str:
    text = str(value)
    if text not in allowed:
        raise UsageError(f"{name} 必须是 {', '.join(sorted(allowed))} 之一，收到 {text!r}")
    return text


def result_data(result: Any) -> dict[str, Any]:
    """Convert a Pydantic result or dict into a JSON-safe data object."""
    if isinstance(result, BaseModel):
        return result.model_dump(mode="json")
    if isinstance(result, dict):
        return result
    raise ToolRuntimeError(f"工具返回了不支持的数据类型：{type(result).__name__}")


def ensure_success(result: Any) -> Any:
    """Raise a runtime error when an upstream tool reports success=false."""
    success = result.get("success") if isinstance(result, dict) else getattr(result, "success", None)
    if success is not False:
        return result
    raise ToolRuntimeError(_failure_message(result))


def _failure_message(result: Any) -> str:
    if isinstance(result, dict):
        error = result.get("error")
        if isinstance(error, dict):
            return str(error.get("message") or error.get("code") or error)
        if error:
            return str(error)
        return str(result.get("message") or "工具返回 success=false")
    return str(getattr(result, "message", "") or "工具返回 success=false")


def _read_result_to_markdown(result: dict) -> str:
    """把 read_* 的扁平 dict 转成干净的 markdown 文本（上游 0.6.0 原样）。

    直接返回 dict 会被 FastMCP 整体序列化为嵌套 JSON，正文大字符串经历双重转义；
    这里把正文作为顶层文本返回，统计信息附在引用行，截断标记保留。
    """
    if not result.get("success"):
        return f"读取失败: {result.get('message') or '未知错误'}"
    parts = [str(result.get("content") or "")]
    stats = {
        k: v for k, v in result.items()
        if k not in {"success", "file_path", "content", "message", "truncated"} and v is not None
    }
    if stats:
        parts.append("> " + "；".join(f"{k}: {v}" for k, v in stats.items()))
    if result.get("truncated"):
        parts.append("> 内容过长已截断")
    return "\n\n".join(p for p in parts if p)


def read_payload(result: dict) -> dict[str, Any]:
    """Return both the upstream Markdown view and the raw structured result."""
    ensure_success(result)
    return {"markdown": _read_result_to_markdown(result), "result": result}


def cli_main(tool: str, parser: JsonArgumentParser, run, argv: list[str] | None = None) -> int:
    """Parse, execute, and map UsageError/ToolRuntimeError/unknown to JSON exits."""
    configure_stdio()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return int(exc.code or 0)
    try:
        data, exit_code = run(args)
        return emit(tool, data, exit_code)
    except UsageError as exc:
        return emit_error(tool, str(exc), 2)
    except ToolRuntimeError as exc:
        return emit_error(tool, str(exc), 3)
    except Exception as exc:  # noqa: BLE001 - CLI boundary must always emit JSON
        return emit_error(tool, f"{type(exc).__name__}: {exc}", 3)
