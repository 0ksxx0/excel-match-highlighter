#!/usr/bin/env python3
"""Compare configurable XLSX columns and highlight matches in the target copy.

The recommended entry points are the interactive prompt or
``match_config.json`` next to this script::

    python 多excal文件相同项对比脚本.py --interactive
    python 多excal文件相同项对比脚本.py --config match_config.json --write

Column lists are paired by position. A target row is a match only when every
configured key column has the same normalized value in one source row. The
target workbook is copied and only selected target cells receive yellow style
clones; neither input workbook is opened for writing.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import posixpath
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path


# Files are resolved relative to this script, so it can be run from PowerShell
# or by double-clicking without relying on the current working directory.
BASE_DIR = Path(__file__).resolve().parent
# These example names are fictional; real paths belong in a local configuration.
DEFAULT_SOURCE_XLSX = BASE_DIR / "source.xlsx"
DEFAULT_TARGET_XLSX = BASE_DIR / "target.xlsx"
DEFAULT_OUTPUT_XLSX = BASE_DIR / "result_highlighted.xlsx"
DEFAULT_CONFIG_FILE = BASE_DIR / "match_config.json"
DEFAULT_TARGET_SHEET_NAME = "参数表"
DEFAULT_SOURCE_SHEET_NAMES = ("来源表A", "来源表B")
DEFAULT_SOURCE_COLUMNS = (
    "点位编码",
    "点位对应端口号",
    "点位寄存器地址",
)
DEFAULT_TARGET_COLUMNS = (
    "参数Code",
    "监听端口",
    "寄存器地址",
)
DEFAULT_NORMALIZERS = ("code", "number", "number")

# A header row can be set to null in the JSON file to auto-detect it by the
# configured header names. The default files both use row 1.
DEFAULT_HEADER_ROW = 1
HEADER_SCAN_ROWS = 10

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", SPREADSHEET_NS)
ET.register_namespace("r", DOCUMENT_REL_NS)

CELL_REFERENCE = re.compile(r"^([A-Z]+)([1-9][0-9]*)$")
POINT_KEY = tuple[str, ...]

# The write path intentionally edits the original XML bytes rather than
# serializing an XML tree. This preserves WPS/Excel extension namespaces,
# relationship prefixes, and other content that ElementTree does not retain.
CELL_XFS_BLOCK_RE = re.compile(
    rb"(?P<open><cellXfs\b[^>]*>)(?P<body>.*?)(?P<close></cellXfs\s*>)",
    re.DOTALL,
)
FILLS_BLOCK_RE = re.compile(
    rb"(?P<open><fills\b[^>]*>)(?P<body>.*?)(?P<close></fills\s*>)",
    re.DOTALL,
)
# Match self-closing xfs first. With a generic greedy start tag, a self-closing
# <xf .../> can otherwise consume following styles until a later </xf>.
XF_RAW_RE = re.compile(
    rb"<xf\b(?:[^>]*?/>|[^>]*>.*?</xf\s*>)",
    re.DOTALL,
)
FILL_RAW_RE = re.compile(
    rb"<fill\b(?:[^>]*?/>|[^>]*>.*?</fill\s*>)",
    re.DOTALL,
)
CELL_START_TAG_RE = re.compile(rb"<c\b[^>]*>", re.DOTALL)
CELL_REFERENCE_ATTR_RE = re.compile(
    rb"\br=(?P<quote>[\"'])(?P<reference>[A-Z]+[1-9][0-9]*)(?P=quote)"
)
ROW_NUMBER_ATTR_RE = re.compile(
    rb"\br=(?P<quote>[\"'])(?P<number>[1-9][0-9]*)(?P=quote)"
)
DIMENSION_TAG_RE = re.compile(rb"<dimension\b[^>]*>", re.DOTALL)
ROW_OPEN_TAG_RE = re.compile(rb"<row\b[^>]*>", re.DOTALL)
ROW_BLOCK_RE = re.compile(
    # The opening tag must not be self-closing.  Without the negative
    # lookbehind, ``<row/>`` can consume a later row's closing tag and make
    # a style-only cell appear in the wrong row.
    rb"<row\b[^>]*(?<!/)>.*?</row\s*>",
    re.DOTALL,
)
ROW_SELF_CLOSING_RE = re.compile(rb"<row\b[^>]*/>", re.DOTALL)


class SafetyStop(RuntimeError):
    """Raised before any output is accepted when an assumption is violated."""


@dataclass(frozen=True)
class SideConfig:
    """One workbook side of a comparison."""

    path: Path
    sheets: tuple[str, ...]
    columns: tuple[str, ...]
    header_row: int | None


@dataclass(frozen=True)
class MatchConfig:
    target: SideConfig
    source: SideConfig
    mark_columns: tuple[str, ...]
    normalizers: tuple[str, ...]
    output_path: Path
    dry_run: bool
    require_all_source_keys_to_match: bool
    max_unmatched_keys_to_print: int


@dataclass(frozen=True)
class HeaderColumns:
    header_row: int
    columns: tuple[str, ...]


@dataclass(frozen=True)
class SourceValues:
    point_keys: frozenset[POINT_KEY]
    key_locations: dict[POINT_KEY, tuple[str, int]]
    sheet_names: tuple[str, ...]
    source_key_rows: int
    source_rows_missing_key_fields: int


@dataclass
class TargetAnalysis:
    sheet_part: str
    sheet_root: ET.Element
    matching_cells: list[ET.Element]
    matching_row_numbers: tuple[int, ...]
    target_keys: frozenset[POINT_KEY]
    target_rows_missing_key_fields: int
    headers: HeaderColumns
    mark_columns: tuple[str, ...]
    missing_mark_references: tuple[str, ...]


def x(tag: str) -> str:
    return f"{{{SPREADSHEET_NS}}}{tag}"


def relationship(tag: str) -> str:
    return f"{{{DOCUMENT_REL_NS}}}{tag}"


def package_relationship(tag: str) -> str:
    return f"{{{PACKAGE_REL_NS}}}{tag}"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalized_header(value: object) -> str:
    """Normalize a header only; spaces are irrelevant in field labels."""
    text = "" if value is None else str(value)
    return "".join(unicodedata.normalize("NFKC", text).split()).upper()


def normalized_code(value: object) -> str:
    """Codes compare case-insensitively and ignore accidental whitespace."""
    text = "" if value is None else str(value)
    return "".join(unicodedata.normalize("NFKC", text).split()).upper()


def normalized_number(value: object) -> str:
    """Canonicalize numeric ports/registers (1502 and 1502.0 compare equal)."""
    raw_text = "" if value is None else str(value)
    text = "".join(unicodedata.normalize("NFKC", raw_text).split())
    if not text:
        return ""
    try:
        number = Decimal(text)
    except InvalidOperation:
        # A non-numeric value is still compared exactly after normalization;
        # this makes malformed source data visible instead of silently matching.
        return text.casefold()
    if not number.is_finite():
        return text.casefold()
    canonical = format(number, "f")
    if "." in canonical:
        canonical = canonical.rstrip("0").rstrip(".")
    return canonical or "0"


def normalized_text(value: object) -> str:
    """Normalize a general text field without changing its case."""
    text = "" if value is None else str(value)
    return unicodedata.normalize("NFKC", text).strip()


NORMALIZER_NAMES = frozenset({"code", "number", "text", "raw"})


def normalize_value(value: object, normalizer: str) -> str:
    """Apply one of the deliberately small, auditable comparison rules."""
    mode = normalizer.casefold()
    if mode == "code":
        return normalized_code(value)
    if mode == "number":
        return normalized_number(value)
    if mode == "text":
        return normalized_text(value)
    if mode == "raw":
        return "" if value is None else str(value)
    raise SafetyStop(
        f"Unsupported normalizer {normalizer!r}; choose one of "
        f"{', '.join(sorted(NORMALIZER_NAMES))}."
    )


def _config_mapping(value: object, label: str) -> dict[str, object]:
    if not isinstance(value, dict):
        raise SafetyStop(f"Configuration field {label!r} must be an object.")
    return value


def _config_list(value: object, label: str) -> tuple[str, ...]:
    """Read a JSON list, also accepting comma/semicolon-separated CLI text."""
    if isinstance(value, str):
        values = re.split(r"[,;\uFF0C\uFF1B\r\n]+", value)
    elif isinstance(value, (list, tuple)):
        values = list(value)
    elif isinstance(value, dict):
        # Accept {code: ..., port: ..., register: ...} as a convenience while
        # retaining the canonical ordered-list format in the example file.
        preferred = ("code", "port", "register")
        if all(name in value for name in preferred):
            values = [value[name] for name in preferred]
        else:
            values = list(value.values())
    else:
        raise SafetyStop(f"Configuration field {label!r} must be a list or text.")

    result: list[str] = []
    for item in values:
        if not isinstance(item, str) or not item.strip():
            raise SafetyStop(f"Configuration field {label!r} contains an empty/non-text item.")
        result.append(item.strip())
    if not result:
        raise SafetyStop(f"Configuration field {label!r} must not be empty.")
    return tuple(result)


def _config_bool(value: object, label: str, default: bool) -> bool:
    if value is None:
        return default
    if not isinstance(value, bool):
        raise SafetyStop(f"Configuration field {label!r} must be true or false.")
    return value


def _config_int(value: object, label: str, default: int | None) -> int | None:
    if value is None:
        return default
    if isinstance(value, bool) or not isinstance(value, int):
        raise SafetyStop(f"Configuration field {label!r} must be an integer or null.")
    return value


def _config_path(value: object, base_dir: Path, label: str) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise SafetyStop(f"Configuration field {label!r} must be a non-empty path.")
    path = Path(value.strip())
    if not path.is_absolute():
        path = base_dir / path
    return path.resolve()


def display_path(path: Path) -> str:
    """Avoid exposing the script's parent directories in console messages."""
    try:
        return str(path.resolve().relative_to(BASE_DIR))
    except ValueError:
        return str(path)


def default_config_data() -> dict[str, object]:
    """Return fictional example settings in read-only mode."""
    return {
        "target": {
            "file": DEFAULT_TARGET_XLSX.name,
            "sheet": DEFAULT_TARGET_SHEET_NAME,
            "header_row": DEFAULT_HEADER_ROW,
            "key_columns": list(DEFAULT_TARGET_COLUMNS),
            "mark_columns": list(DEFAULT_TARGET_COLUMNS),
        },
        "source": {
            "file": DEFAULT_SOURCE_XLSX.name,
            "sheets": list(DEFAULT_SOURCE_SHEET_NAMES),
            "header_row": DEFAULT_HEADER_ROW,
            "key_columns": list(DEFAULT_SOURCE_COLUMNS),
        },
        "output": DEFAULT_OUTPUT_XLSX.name,
        "normalizers": list(DEFAULT_NORMALIZERS),
        "options": {
            "dry_run": True,
            "require_all_source_keys_to_match": False,
            "max_unmatched_keys": 20,
        },
    }


def load_raw_config(config_path: Path | None) -> tuple[dict[str, object], Path]:
    """Load JSON configuration, or use embedded defaults when no file exists."""
    selected = config_path
    if selected is None and DEFAULT_CONFIG_FILE.is_file():
        selected = DEFAULT_CONFIG_FILE
    if selected is None:
        return default_config_data(), BASE_DIR

    selected = selected.resolve()
    if not selected.is_file():
        raise SafetyStop(f"Configuration file does not exist: {display_path(selected)}")
    try:
        raw = json.loads(selected.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise SafetyStop(f"Could not read JSON configuration {selected}: {error}") from error
    if not isinstance(raw, dict):
        raise SafetyStop("The top level of the JSON configuration must be an object.")
    return raw, selected.parent


def _interactive_mapping(raw: dict[str, object], key: str) -> dict[str, object]:
    value = raw.get(key)
    return value if isinstance(value, dict) else {}


def _interactive_default_list(
    mapping: dict[str, object], keys: tuple[str, ...], fallback: tuple[str, ...]
) -> tuple[str, ...]:
    value: object = None
    for key in keys:
        if key in mapping:
            value = mapping[key]
            break
    if value is None:
        return fallback
    try:
        return _config_list(value, keys[0])
    except SafetyStop:
        return fallback


def _interactive_default_path(
    mapping: dict[str, object], keys: tuple[str, ...], base_dir: Path, label: str
) -> str | None:
    value: object = None
    for key in keys:
        if key in mapping:
            value = mapping[key]
            break
    if isinstance(value, dict):
        value = value.get("file", value.get("path"))
    if not isinstance(value, str) or not value.strip():
        return None
    try:
        return str(_config_path(value, base_dir, label))
    except SafetyStop:
        return value.strip()


def _interactive_input(prompt: str) -> str:
    try:
        return input(prompt)
    except (EOFError, KeyboardInterrupt) as error:
        raise SafetyStop("Interactive input was cancelled.") from error


def _strip_outer_quotes(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
        return value[1:-1].strip()
    return value


def _prompt_text(label: str, default: object = None, *, is_path: bool = False) -> str:
    default_text = "" if default is None else str(default).strip()
    display_text = display_path(Path(default_text)) if is_path and default_text else default_text
    suffix = f" [{display_text}]" if display_text else ""
    while True:
        value = _strip_outer_quotes(_interactive_input(f"{label}{suffix}: "))
        if value:
            return value
        if default_text:
            return default_text
        print("此项不能为空，请重新输入。")


def _prompt_int(label: str, default: int | None, minimum: int = 1) -> int:
    default_text = "" if default is None else str(default)
    suffix = f" [{default_text}]" if default_text else ""
    while True:
        value = _strip_outer_quotes(_interactive_input(f"{label}{suffix}: "))
        if not value and default is not None:
            return default
        try:
            number = int(value)
        except ValueError:
            number = 0
        if number >= minimum:
            return number
        print(f"请输入不小于 {minimum} 的整数。")


def _prompt_header_row(label: str, default: int | None) -> int | None:
    default_text = "自动识别" if default is None else str(default)
    while True:
        value = _strip_outer_quotes(_interactive_input(f"{label}（输入 auto 可自动识别） [{default_text}]: "))
        if not value:
            return default
        if value.casefold() in {"auto", "null", "none", "自动", "自动识别"}:
            return None
        try:
            number = int(value)
        except ValueError:
            number = 0
        if number >= 1:
            return number
        print("表头行必须是大于等于 1 的整数，或输入 auto。")


def _prompt_yes_no(label: str, default: bool) -> bool:
    suffix = " [Y/n]" if default else " [y/N]"
    yes = {"y", "yes", "1", "是", "确定"}
    no = {"n", "no", "0", "否", "取消"}
    while True:
        value = _strip_outer_quotes(_interactive_input(f"{label}{suffix}: ")).casefold()
        if not value:
            return default
        if value in yes:
            return True
        if value in no:
            return False
        print("请输入 y/yes 或 n/no。")


def _prompt_normalizer(label: str, default: str) -> str:
    default = default.casefold() if default.casefold() in NORMALIZER_NAMES else "text"
    while True:
        value = _prompt_text(
            f"{label}（code=忽略大小写和空格，number=数字，text=文本，raw=原始）",
            default,
        ).casefold()
        if value in NORMALIZER_NAMES:
            return value
        print(f"比较规则必须是：{', '.join(sorted(NORMALIZER_NAMES))}。")


def interactive_config(
    config_path: Path | None,
    force_write: bool,
    force_dry_run: bool,
) -> tuple[dict[str, object], Path]:
    """Collect a comparison configuration through ordinary console prompts."""
    defaults, defaults_base = load_raw_config(config_path)
    target_defaults = _interactive_mapping(defaults, "target")
    source_defaults = _interactive_mapping(defaults, "source")
    options_defaults = _interactive_mapping(defaults, "options")

    target_file = _prompt_text(
        "表1（需要标黄）的 Excel 文件路径",
        _interactive_default_path(
            target_defaults, ("file", "path"), defaults_base, "target.file"
        ),
        is_path=True,
    )
    target_sheet_default = target_defaults.get(
        "sheet", target_defaults.get("worksheet")
    )
    if not isinstance(target_sheet_default, str) or not target_sheet_default.strip():
        target_sheet_default = None
    target_sheet = _prompt_text("表1中需要标黄的工作表名称", target_sheet_default)

    source_file = _prompt_text(
        "表2（对比来源）的 Excel 文件路径",
        _interactive_default_path(
            source_defaults, ("file", "path"), defaults_base, "source.file"
        ),
        is_path=True,
    )
    source_sheet_defaults = _interactive_default_list(
        source_defaults,
        ("sheets", "worksheets"),
        DEFAULT_SOURCE_SHEET_NAMES,
    )
    source_sheet_count = _prompt_int(
        "表2需要参与对比的工作表数量", len(source_sheet_defaults), 1
    )
    source_sheets: list[str] = []
    for index in range(source_sheet_count):
        default_sheet = (
            source_sheet_defaults[index]
            if index < len(source_sheet_defaults)
            else None
        )
        source_sheets.append(
            _prompt_text(f"表2第 {index + 1} 个工作表名称", default_sheet)
        )

    target_column_defaults = _interactive_default_list(
        target_defaults,
        ("key_columns", "columns"),
        DEFAULT_TARGET_COLUMNS,
    )
    source_column_defaults = _interactive_default_list(
        source_defaults,
        ("key_columns", "columns"),
        DEFAULT_SOURCE_COLUMNS,
    )
    pair_count = _prompt_int(
        "需要比较的字段组数", len(target_column_defaults), 1
    )
    target_columns: list[str] = []
    source_columns: list[str] = []
    for index in range(pair_count):
        target_default = (
            target_column_defaults[index]
            if index < len(target_column_defaults)
            else None
        )
        source_default = (
            source_column_defaults[index]
            if index < len(source_column_defaults)
            else None
        )
        target_columns.append(
            _prompt_text(
                f"对比项 {index + 1}：表1（目标表）中的列名或列字母",
                target_default,
            )
        )
        source_columns.append(
            _prompt_text(
                f"对比项 {index + 1}：表2（来源表）中的列名或列字母",
                source_default,
            )
        )

    configured_normalizer_defaults = _interactive_default_list(
        {"normalizers": defaults.get("normalizers")},
        ("normalizers",),
        ("text",) * pair_count,
    )
    # Configured code/number rules are useful when retaining the
    # example columns, but arbitrary new column pairs should start
    # from conservative text comparison.
    same_as_configured_columns = (
        tuple(target_columns) == target_column_defaults
        and tuple(source_columns) == source_column_defaults
    )
    normalizer_defaults = (
        configured_normalizer_defaults
        if same_as_configured_columns
        else ("text",) * pair_count
    )
    normalizers = [
        _prompt_normalizer(
            f"对比项 {index + 1} 的比较规则",
            normalizer_defaults[index]
            if index < len(normalizer_defaults)
            else "text",
        )
        for index in range(pair_count)
    ]

    # A new interactive column selection should also become the default yellow
    # selection.  Reusing a previous config's mark_columns here could color
    # unrelated fields after the user changes the comparison columns.
    mark_default_text = ",".join(target_columns)
    mark_columns = _config_list(
        _prompt_text(
            "命中后要标黄的表1列（逗号分隔，直接回车使用目标对比列）",
            mark_default_text,
        ),
        "target.mark_columns",
    )

    target_header_default = target_defaults.get("header_row", 1)
    source_header_default = source_defaults.get("header_row", 1)
    if not isinstance(target_header_default, int) or isinstance(target_header_default, bool):
        target_header_default = None
    if not isinstance(source_header_default, int) or isinstance(source_header_default, bool):
        source_header_default = None
    target_header_row = _prompt_header_row("表1表头所在行", target_header_default)
    source_header_row = _prompt_header_row("表2表头所在行", source_header_default)

    target_path = _config_path(target_file, Path.cwd(), "target.file")
    output_default = _interactive_default_path(
        defaults, ("output", "output_file"), defaults_base, "output"
    )
    if output_default is None:
        output_default = str(target_path.with_name(f"{target_path.stem}_匹配标黄.xlsx"))
    output_default_path = Path(output_default)
    if output_default_path.exists():
        suffix = output_default_path.suffix or ".xlsx"
        index = 1
        candidate = output_default_path.with_name(
            f"{output_default_path.stem}_{index}{suffix}"
        )
        while candidate.exists():
            index += 1
            candidate = output_default_path.with_name(
                f"{output_default_path.stem}_{index}{suffix}"
            )
        output_default = str(candidate)
    output = _prompt_text("输出文件路径（不会覆盖已有文件）", output_default, is_path=True)

    if force_write:
        dry_run = False
    elif force_dry_run:
        dry_run = True
    else:
        dry_run = not _prompt_yes_no("是否立即生成标黄文件", False)

    require_all = options_defaults.get(
        "require_all_source_keys_to_match",
        options_defaults.get("require_all_source_keys", False),
    )
    if not isinstance(require_all, bool):
        require_all = False
    max_unmatched = options_defaults.get(
        "max_unmatched_keys", options_defaults.get("max_unmatched_keys_to_print", 20)
    )
    if isinstance(max_unmatched, bool) or not isinstance(max_unmatched, int) or max_unmatched < 0:
        max_unmatched = 20

    return (
        {
            "target": {
                "file": target_file,
                "sheet": target_sheet,
                "header_row": target_header_row,
                "key_columns": target_columns,
                "mark_columns": list(mark_columns),
            },
            "source": {
                "file": source_file,
                "sheets": source_sheets,
                "header_row": source_header_row,
                "key_columns": source_columns,
            },
            "normalizers": normalizers,
            "output": output,
            "options": {
                "dry_run": dry_run,
                "require_all_source_keys_to_match": require_all,
                "max_unmatched_keys": max_unmatched,
            },
        },
        Path.cwd().resolve(),
    )


def build_match_config(raw: dict[str, object], base_dir: Path) -> MatchConfig:
    target_raw = _config_mapping(raw.get("target"), "target")
    source_raw = _config_mapping(raw.get("source"), "source")

    target_file = target_raw.get("file", target_raw.get("path"))
    source_file = source_raw.get("file", source_raw.get("path"))
    target_sheet = target_raw.get("sheet", target_raw.get("worksheet"))
    if not isinstance(target_sheet, str) or not target_sheet.strip():
        raise SafetyStop("Configuration field 'target.sheet' must be a non-empty name.")

    source_sheets = _config_list(
        source_raw.get("sheets", source_raw.get("worksheets")),
        "source.sheets",
    )
    if len(set(source_sheets)) != len(source_sheets):
        raise SafetyStop("Configuration field 'source.sheets' contains duplicates.")

    target_columns_value = target_raw.get(
        "key_columns", target_raw.get("columns")
    )
    source_columns_value = source_raw.get(
        "key_columns", source_raw.get("columns")
    )
    target_columns = _config_list(target_columns_value, "target.key_columns")
    source_columns = _config_list(source_columns_value, "source.key_columns")
    if len(target_columns) != len(source_columns):
        raise SafetyStop(
            "target.key_columns and source.key_columns must have the same length "
            "and are paired by position."
        )
    if len(target_columns) < 1:
        raise SafetyStop("At least one key column is required on each side.")

    mark_columns = _config_list(
        target_raw.get("mark_columns", raw.get("mark_columns", target_columns)),
        "target.mark_columns",
    )

    options = _config_mapping(raw.get("options", {}), "options")
    normalizers_value = raw.get("normalizers", options.get("normalizers"))
    if normalizers_value is None:
        # Generic configurations default to conservative text comparison.
        # The bundled point-table example opts into code/number normalization
        # explicitly in match_config.json.
        normalizers = ("text",) * len(target_columns)
    else:
        normalizers = tuple(
            item.casefold()
            for item in _config_list(normalizers_value, "normalizers")
        )
    if len(normalizers) != len(target_columns):
        raise SafetyStop("normalizers must have one entry for each key column.")
    invalid_normalizers = sorted(set(normalizers) - NORMALIZER_NAMES)
    if invalid_normalizers:
        raise SafetyStop(
            f"Unsupported normalizers: {invalid_normalizers}; "
            f"choose from {sorted(NORMALIZER_NAMES)}."
        )

    # ``null`` is meaningful for header rows: it asks the resolver to search
    # the first few rows.  Preserve an explicit null instead of letting the
    # integer helper replace it with the ordinary row-1 default.
    target_header_value = target_raw.get("header_row", DEFAULT_HEADER_ROW)
    source_header_value = source_raw.get("header_row", DEFAULT_HEADER_ROW)
    target_header_row = (
        None
        if target_header_value is None
        else _config_int(
            target_header_value,
            "target.header_row",
            DEFAULT_HEADER_ROW,
        )
    )
    source_header_row = (
        None
        if source_header_value is None
        else _config_int(
            source_header_value,
            "source.header_row",
            DEFAULT_HEADER_ROW,
        )
    )
    for label, value in (
        ("target.header_row", target_header_row),
        ("source.header_row", source_header_row),
    ):
        if value is not None and value < 1:
            raise SafetyStop(f"Configuration field {label!r} must be >= 1 or null.")

    output_value = raw.get("output", raw.get("output_file"))
    if isinstance(output_value, dict):
        output_object = _config_mapping(output_value, "output")
        overwrite = output_object.get("overwrite", False)
        if not isinstance(overwrite, bool):
            raise SafetyStop("output.overwrite must be false; overwriting is disabled.")
        if overwrite:
            raise SafetyStop("Output overwrite is intentionally disabled for safety.")
        output_value = output_object.get("file", output_object.get("path"))

    target_path = _config_path(target_file, base_dir, "target.file")
    source_path = _config_path(source_file, base_dir, "source.file")
    if output_value is None:
        output_path = target_path.with_name(f"{target_path.stem}_匹配标黄.xlsx")
    else:
        output_path = _config_path(output_value, base_dir, "output")

    dry_run = _config_bool(options.get("dry_run"), "options.dry_run", True)
    require_all = _config_bool(
        options.get(
            "require_all_source_keys_to_match",
            options.get("require_all_source_keys"),
        ),
        "options.require_all_source_keys_to_match",
        False,
    )
    max_unmatched = _config_int(
        options.get(
            "max_unmatched_keys",
            options.get("max_unmatched_keys_to_print"),
        ),
        "options.max_unmatched_keys",
        20,
    )
    if max_unmatched is None or max_unmatched < 0:
        raise SafetyStop("options.max_unmatched_keys must be a non-negative integer.")

    return MatchConfig(
        target=SideConfig(target_path, (target_sheet.strip(),), target_columns, target_header_row),
        source=SideConfig(source_path, source_sheets, source_columns, source_header_row),
        mark_columns=mark_columns,
        normalizers=normalizers,
        output_path=output_path,
        dry_run=dry_run,
        require_all_source_keys_to_match=require_all,
        max_unmatched_keys_to_print=max_unmatched,
    )


def argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Compare configured XLSX key columns and highlight target matches."
    )
    parser.add_argument(
        "--config",
        type=Path,
        help="JSON configuration file (default: match_config.json next to the script).",
    )
    parser.add_argument(
        "-i",
        "--interactive",
        "--prompt",
        action="store_true",
        help="Ask for files, sheets, comparison columns, and output settings interactively (default with no arguments).",
    )
    parser.add_argument("--target-file", "--target", dest="target_file")
    parser.add_argument("--target-sheet", dest="target_sheet")
    parser.add_argument(
        "--target-columns", "--target-key-columns", dest="target_columns",
        help="Comma-separated target key headers or Excel column letters.",
    )
    parser.add_argument(
        "--mark-columns", dest="mark_columns",
        help="Comma-separated target columns to mark; defaults to target key columns.",
    )
    parser.add_argument("--source-file", "--source", dest="source_file")
    parser.add_argument(
        "--source-sheets", dest="source_sheets",
        help="Comma-separated source worksheet names.",
    )
    parser.add_argument(
        "--source-sheet", dest="source_sheet", action="append",
        help="Source worksheet name; may be repeated.",
    )
    parser.add_argument(
        "--source-columns", "--source-key-columns", dest="source_columns",
        help="Comma-separated source key headers or Excel column letters.",
    )
    parser.add_argument("--target-header-row", type=int, dest="target_header_row")
    parser.add_argument("--source-header-row", type=int, dest="source_header_row")
    parser.add_argument(
        "--normalizers",
        help="Comma-separated normalizers: code, number, text, or raw.",
    )
    parser.add_argument("--output", "--output-file", dest="output")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true", help="Create the output workbook.")
    mode.add_argument("--dry-run", action="store_true", help="Validate only; do not create output.")
    source_gate = parser.add_mutually_exclusive_group()
    source_gate.add_argument(
        "--require-all-source-keys", dest="require_all_source_keys",
        action="store_true", help="Stop unless every source key exists in target.",
    )
    source_gate.add_argument(
        "--allow-unmatched-source-keys", dest="require_all_source_keys",
        action="store_false", help="Highlight confirmed matches and report unmatched keys.",
    )
    parser.set_defaults(require_all_source_keys=None)
    parser.add_argument("--max-unmatched-keys", type=int, dest="max_unmatched_keys")
    return parser


def apply_cli_overrides(raw: dict[str, object], args: argparse.Namespace) -> None:
    """Apply only explicitly supplied CLI values to the loaded JSON object."""
    target = raw.setdefault("target", {})
    source = raw.setdefault("source", {})
    options = raw.setdefault("options", {})
    if not isinstance(target, dict) or not isinstance(source, dict) or not isinstance(options, dict):
        raise SafetyStop("target, source, and options must be objects before CLI overrides.")

    if args.target_file is not None:
        target["file"] = str(Path(args.target_file).resolve())
    if args.target_sheet is not None:
        target["sheet"] = args.target_sheet
    if args.target_columns is not None:
        target["key_columns"] = list(_config_list(args.target_columns, "--target-columns"))
    if args.mark_columns is not None:
        target["mark_columns"] = list(_config_list(args.mark_columns, "--mark-columns"))
    if args.source_file is not None:
        source["file"] = str(Path(args.source_file).resolve())
    if args.source_sheets is not None and args.source_sheet:
        raise SafetyStop("Use either --source-sheets or repeated --source-sheet, not both.")
    if args.source_sheets is not None:
        source["sheets"] = list(_config_list(args.source_sheets, "--source-sheets"))
    elif args.source_sheet:
        source["sheets"] = list(_config_list(args.source_sheet, "--source-sheet"))
    if args.source_columns is not None:
        source["key_columns"] = list(_config_list(args.source_columns, "--source-columns"))
    if args.target_header_row is not None:
        target["header_row"] = args.target_header_row
    if args.source_header_row is not None:
        source["header_row"] = args.source_header_row
    if args.normalizers is not None:
        raw["normalizers"] = list(_config_list(args.normalizers, "--normalizers"))
    if args.output is not None:
        raw["output"] = str(Path(args.output).resolve())
    if args.write:
        options["dry_run"] = False
    elif args.dry_run:
        options["dry_run"] = True
    if args.require_all_source_keys is not None:
        options["require_all_source_keys_to_match"] = args.require_all_source_keys
    if args.max_unmatched_keys is not None:
        options["max_unmatched_keys"] = args.max_unmatched_keys


def cell_reference_parts(reference: str | None) -> tuple[str, int]:
    match = CELL_REFERENCE.fullmatch(reference or "")
    if match is None:
        raise SafetyStop(f"Invalid cell reference in workbook XML: {reference!r}")
    return match.group(1), int(match.group(2))


def read_xml(archive: zipfile.ZipFile, part: str) -> ET.Element:
    try:
        return ET.fromstring(archive.read(part))
    except KeyError as error:
        raise SafetyStop(f"Required XLSX package part is missing: {part}") from error
    except ET.ParseError as error:
        raise SafetyStop(f"Invalid XML in XLSX package part: {part}") from error


def load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    part = "xl/sharedStrings.xml"
    if part not in archive.namelist():
        return []

    root = read_xml(archive, part)
    values: list[str] = []
    for item in root.findall(x("si")):
        values.append("".join(node.text or "" for node in item.iter(x("t"))))
    return values


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(x("t")))

    value = cell.find(x("v"))
    if value is None or value.text is None:
        return ""

    if cell_type == "s":
        try:
            index = int(value.text)
            if index < 0:
                raise ValueError("negative shared-string index")
            return shared_strings[index]
        except (IndexError, ValueError) as error:
            reference = cell.get("r", "<unknown>")
            raise SafetyStop(
                f"Invalid shared-string reference at {reference}: {value.text!r}"
            ) from error
    return value.text


def has_formula(cell: ET.Element) -> bool:
    return cell.find(x("f")) is not None


def row_cells(row: ET.Element) -> dict[str, ET.Element]:
    """Index a worksheet row by column and reject duplicate cell references."""
    indexed: dict[str, ET.Element] = {}
    for cell in row.findall(x("c")):
        column, _ = cell_reference_parts(cell.get("r"))
        if column in indexed:
            raise SafetyStop(f"Duplicate cell column in row {row.get('r')}: {column}")
        indexed[column] = cell
    return indexed


def key_value(
    cell: ET.Element | None,
    shared_strings: list[str],
    normalizer: str,
    location: str,
) -> str:
    if cell is None:
        return ""
    if has_formula(cell):
        raise SafetyStop(f"Match field contains a formula: {location}")
    value = cell_text(cell, shared_strings)
    return normalize_value(value, normalizer)


def row_key(
    row: ET.Element,
    headers: HeaderColumns,
    shared_strings: list[str],
    sheet_name: str,
    normalizers: tuple[str, ...],
) -> POINT_KEY | None:
    cells = row_cells(row)
    row_number = row.get("r", "?")
    if len(headers.columns) != len(normalizers):
        raise SafetyStop("Resolved key columns and normalizers have different lengths.")
    values = tuple(
        key_value(
            cells.get(column),
            shared_strings,
            normalizer,
            f"{sheet_name}!{column}{row_number}",
        )
        for column, normalizer in zip(headers.columns, normalizers)
    )
    if any(not value for value in values):
        return None
    return values


def sheet_rows(sheet_root: ET.Element) -> list[ET.Element]:
    data = sheet_root.find(x("sheetData"))
    if data is None:
        raise SafetyStop("Worksheet has no sheetData element.")
    return data.findall(x("row"))


def resolve_part(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def workbook_sheet_parts(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = read_xml(archive, "xl/workbook.xml")
    rels = read_xml(archive, "xl/_rels/workbook.xml.rels")
    relationship_targets = {
        node.get("Id"): node.get("Target")
        for node in rels.findall(package_relationship("Relationship"))
        if node.get("Id") and node.get("Target")
    }

    sheets = workbook.find(x("sheets"))
    if sheets is None:
        raise SafetyStop("workbook.xml has no sheets element.")

    result: list[tuple[str, str]] = []
    seen_names: set[str] = set()
    for sheet in sheets.findall(x("sheet")):
        name = sheet.get("name")
        rel_id = sheet.get(relationship("id"))
        if not name or not rel_id or rel_id not in relationship_targets:
            raise SafetyStop("A workbook sheet has no resolvable worksheet relationship.")
        if name in seen_names:
            raise SafetyStop(f"Workbook contains duplicate sheet name: {name!r}")
        seen_names.add(name)
        result.append((name, resolve_part("xl/workbook.xml", relationship_targets[rel_id])))
    return result


def column_letter_from_spec(spec: str) -> str | None:
    """Return an Excel column letter when *spec* is written as a letter."""
    def checked_letter(text: str) -> str:
        candidate = text.upper()
        number = 0
        for character in candidate:
            number = number * 26 + ord(character) - ord("A") + 1
        if number > 16384:
            raise SafetyStop(f"Excel column is outside XFD: {spec!r}")
        return candidate

    text = spec.strip()
    lowered = text.casefold()
    if lowered.startswith("header:"):
        return None
    if lowered.startswith("column:") or lowered.startswith("col:"):
        text = text.split(":", 1)[1].strip()
        if not re.fullmatch(r"[A-Za-z]+", text):
            raise SafetyStop(f"Invalid Excel column specification: {spec!r}")
        return checked_letter(text)
    # Excel columns end at XFD. Limiting the implicit form to three letters
    # avoids treating common headers such as CODE or PORT as column letters.
    if re.fullmatch(r"[A-Za-z]{1,3}", text):
        candidate = checked_letter(text)
        return candidate
    return None


def header_name_from_spec(spec: str) -> str:
    text = spec.strip()
    if text.casefold().startswith("header:"):
        text = text.split(":", 1)[1].strip()
    return normalized_header(text)


def _row_number(row: ET.Element, sheet_name: str) -> int:
    try:
        return int(row.get("r", "0"))
    except ValueError as error:
        raise SafetyStop(
            f"Invalid row number in worksheet {sheet_name!r}: {row.get('r')!r}"
        ) from error


def _resolve_columns_in_header(
    row: ET.Element,
    shared_strings: list[str],
    specs: tuple[str, ...],
    sheet_name: str,
    allow_missing: bool,
) -> tuple[str, ...] | None:
    labels: dict[str, list[str]] = {}
    for cell in row.findall(x("c")):
        label = normalized_header(cell_text(cell, shared_strings))
        if not label:
            continue
        column, _ = cell_reference_parts(cell.get("r"))
        labels.setdefault(label, []).append(column)

    columns: list[str] = []
    for spec in specs:
        text = spec.strip()
        lowered = text.casefold()
        forced_column = lowered.startswith("column:") or lowered.startswith("col:")
        forced_header = lowered.startswith("header:")
        candidates = labels.get(header_name_from_spec(spec), [])
        if len(candidates) > 1 and not forced_column:
            raise SafetyStop(
                f"Header {spec!r} occurs more than once in "
                f"{sheet_name}!{row.get('r', '?')}."
            )

        # Bare short ASCII names are ambiguous (for example, ``ID`` can be a
        # header or an Excel column). Prefer an actual header match and use a
        # column letter only as the fallback. ``column:ID`` and ``header:ID``
        # remain available when the intent must be explicit.
        if candidates and not forced_column:
            column = candidates[0]
        elif forced_header:
            if allow_missing:
                return None
            raise SafetyStop(
                f"Configured column {spec!r} was not found in "
                f"{sheet_name}!{row.get('r', '?')}."
            )
        else:
            explicit_letter = column_letter_from_spec(spec)
            if explicit_letter is None:
                if allow_missing:
                    return None
                raise SafetyStop(
                    f"Configured column {spec!r} was not found in "
                    f"{sheet_name}!{row.get('r', '?')}."
                )
            column = explicit_letter
        if column in columns:
            raise SafetyStop(
                f"Configured columns resolve to duplicate column {column} in {sheet_name}."
            )
        columns.append(column)
    return tuple(columns)


def resolve_headers(
    sheet_root: ET.Element,
    shared_strings: list[str],
    specs: tuple[str, ...],
    header_row: int | None,
    sheet_name: str,
) -> HeaderColumns:
    """Resolve configured names/letters to worksheet columns."""
    rows = sheet_rows(sheet_root)
    if header_row is not None:
        matching_rows = [row for row in rows if _row_number(row, sheet_name) == header_row]
        if len(matching_rows) != 1:
            raise SafetyStop(
                f"Expected exactly one header row {header_row} in {sheet_name}, "
                f"found {len(matching_rows)}."
            )
        columns = _resolve_columns_in_header(
            matching_rows[0], shared_strings, specs, sheet_name, allow_missing=False
        )
        if columns is None:
            raise SafetyStop(
                f"Could not resolve configured columns {list(specs)} in "
                f"{sheet_name}!{header_row}."
            )
        return HeaderColumns(header_row, columns)

    # With header_row=null, names are searched only in the first few rows. A
    # pure column-letter configuration uses the first worksheet row as the
    # data boundary, which is the least surprising Excel convention.
    for row in rows:
        row_number = _row_number(row, sheet_name)
        if row_number > HEADER_SCAN_ROWS:
            continue
        columns = _resolve_columns_in_header(
            row, shared_strings, specs, sheet_name, allow_missing=True
        )
        if columns is not None:
            return HeaderColumns(row_number, columns)
    raise SafetyStop(
        f"Could not resolve configured columns {list(specs)} in the first "
        f"{HEADER_SCAN_ROWS} rows of {sheet_name}."
    )


def reject_signed_or_macro_package(archive: zipfile.ZipFile, path: Path) -> None:
    protected_parts = [
        name
        for name in archive.namelist()
        if name.lower() == "xl/vbaproject.bin"
        or name.lower().startswith("_xmlsignatures/")
        or name.lower().startswith("xl/signatures/")
    ]
    if protected_parts:
        raise SafetyStop(
            f"{path.name} contains macro/signature parts and is refused: {protected_parts}"
        )


def reject_duplicate_zip_members(archive: zipfile.ZipFile, path: Path) -> None:
    names = archive.namelist()
    if len(names) != len(set(names)):
        raise SafetyStop(f"{path.name} contains duplicate ZIP member names and is refused.")


def extract_source_values(
    source_config: SideConfig, normalizers: tuple[str, ...]
) -> SourceValues:
    point_keys: set[POINT_KEY] = set()
    key_locations: dict[POINT_KEY, tuple[str, int]] = {}
    selected_sheets: list[str] = []
    source_key_rows = 0
    source_rows_missing_key_fields = 0

    with zipfile.ZipFile(source_config.path, "r") as archive:
        reject_signed_or_macro_package(archive, source_config.path)
        reject_duplicate_zip_members(archive, source_config.path)
        shared_strings = load_shared_strings(archive)
        sheet_parts = dict(workbook_sheet_parts(archive))
        missing_sheets = [
            name for name in source_config.sheets if name not in sheet_parts
        ]
        if missing_sheets:
            raise SafetyStop(
                f"Source worksheet(s) not found in {source_config.path.name}: "
                f"{missing_sheets}; available sheets: {list(sheet_parts)}"
            )
        for sheet_name in source_config.sheets:
            sheet_part = sheet_parts[sheet_name]
            root = read_xml(archive, sheet_part)
            headers = resolve_headers(
                root,
                shared_strings,
                source_config.columns,
                source_config.header_row,
                sheet_name,
            )
            selected_sheets.append(sheet_name)
            for row in sheet_rows(root):
                row_number = _row_number(row, sheet_name)
                if row_number <= headers.header_row:
                    continue
                key = row_key(row, headers, shared_strings, sheet_name, normalizers)
                if key is None:
                    source_rows_missing_key_fields += 1
                    continue
                if key in key_locations:
                    previous_sheet, previous_row = key_locations[key]
                    raise SafetyStop(
                        "Duplicate source key: "
                        f"{key!r} at {previous_sheet}!{previous_row} and "
                        f"{sheet_name}!{row_number}"
                    )
                point_keys.add(key)
                key_locations[key] = (sheet_name, row_number)
                source_key_rows += 1

    if not point_keys:
        raise SafetyStop("Source has no complete rows for the configured key columns.")

    return SourceValues(
        frozenset(point_keys),
        key_locations,
        tuple(selected_sheets),
        source_key_rows,
        source_rows_missing_key_fields,
    )


def analyze_target(
    target_config: SideConfig,
    source_values: SourceValues,
    mark_columns: tuple[str, ...],
    normalizers: tuple[str, ...],
) -> TargetAnalysis:
    target_sheet_name = target_config.sheets[0]
    with zipfile.ZipFile(target_config.path, "r") as archive:
        reject_signed_or_macro_package(archive, target_config.path)
        reject_duplicate_zip_members(archive, target_config.path)
        shared_strings = load_shared_strings(archive)
        sheet_parts = dict(workbook_sheet_parts(archive))
        if target_sheet_name not in sheet_parts:
            raise SafetyStop(
                f"Target sheet {target_sheet_name!r} was not found. "
                f"Available sheets: {list(sheet_parts)}"
            )

        sheet_part = sheet_parts[target_sheet_name]
        root = read_xml(archive, sheet_part)
        headers = resolve_headers(
            root,
            shared_strings,
            target_config.columns,
            target_config.header_row,
            target_sheet_name,
        )
        resolved_mark_columns = resolve_headers(
            root,
            shared_strings,
            mark_columns,
            headers.header_row,
            target_sheet_name,
        ).columns

        matches: list[ET.Element] = []
        matching_rows: list[int] = []
        target_keys: set[POINT_KEY] = set()
        target_rows_missing_key_fields = 0
        missing_mark_references: list[str] = []
        seen_references: set[str] = set()
        for row in sheet_rows(root):
            row_number = _row_number(row, target_sheet_name)
            if row_number <= headers.header_row:
                continue
            key = row_key(row, headers, shared_strings, target_sheet_name, normalizers)
            if key is None:
                target_rows_missing_key_fields += 1
                continue
            target_keys.add(key)
            if key not in source_values.point_keys:
                continue

            matching_rows.append(row_number)
            cells = row_cells(row)
            # Only the explicitly selected target columns are highlighted;
            # all other parameter columns and their formatting stay unchanged.
            for column in resolved_mark_columns:
                cell = cells.get(column)
                if cell is None:
                    # XLSX commonly omits a blank cell entirely. Keep a
                    # placeholder so the write path can materialize that
                    # cell with a yellow style without inventing a value.
                    reference = f"{column}{row_number}"
                    cell = ET.Element(x("c"), {"r": reference})
                    missing_mark_references.append(reference)
                reference = cell.get("r", "")
                if reference not in seen_references:
                    matches.append(cell)
                    seen_references.add(reference)

    return TargetAnalysis(
        sheet_part,
        root,
        matches,
        tuple(matching_rows),
        frozenset(target_keys),
        target_rows_missing_key_fields,
        headers,
        resolved_mark_columns,
        tuple(missing_mark_references),
    )


def is_solid_yellow(fill: ET.Element) -> bool:
    pattern = fill.find(x("patternFill"))
    if pattern is None or pattern.get("patternType") != "solid":
        return False
    foreground = pattern.find(x("fgColor"))
    if foreground is None:
        return False
    rgb = foreground.get("rgb", "").upper().lstrip("#")
    return rgb == "FFFF00" or (len(rgb) == 8 and rgb.endswith("FFFF00"))


def find_yellow_fill_id(styles_root: ET.Element) -> int:
    fills = styles_root.find(x("fills"))
    if fills is None:
        raise SafetyStop("styles.xml has no fills element.")

    fill_nodes = fills.findall(x("fill"))
    for index, fill in enumerate(fill_nodes):
        if is_solid_yellow(fill):
            return index
    return -1


def ensure_yellow_fill(
    styles_raw: bytes, styles_root: ET.Element
) -> tuple[bytes, ET.Element, int]:
    """Reuse a yellow fill or append one while preserving existing XML bytes."""
    existing_id = find_yellow_fill_id(styles_root)
    if existing_id >= 0:
        return styles_raw, styles_root, existing_id

    fills = styles_root.find(x("fills"))
    if fills is None:
        raise SafetyStop("styles.xml has no fills element.")
    fill_nodes = fills.findall(x("fill"))
    block = FILLS_BLOCK_RE.search(styles_raw)
    if block is None or FILLS_BLOCK_RE.search(styles_raw, block.end()) is not None:
        raise SafetyStop("styles.xml does not contain exactly one raw fills block.")
    raw_count = raw_xml_attribute(block.group("open"), b"count")
    try:
        declared_count = int((raw_count or b"").decode("ascii"))
    except ValueError as error:
        raise SafetyStop("fills has no valid count attribute.") from error
    raw_fills = FILL_RAW_RE.findall(block.group("body"))
    if declared_count != len(fill_nodes) or len(raw_fills) != len(fill_nodes):
        raise SafetyStop(
            "The raw and parsed fills counts disagree; refusing to edit styles.xml."
        )

    # Eight-digit ARGB is accepted by Excel/WPS; indexed=64 keeps the
    # background definition conventional for a solid fill.
    yellow_fill = (
        b"<fill><patternFill patternType=\"solid\">"
        b"<fgColor rgb=\"FFFFFF00\"/><bgColor indexed=\"64\"/>"
        b"</patternFill></fill>"
    )
    patched_open = set_raw_xml_attribute(
        block.group("open"), b"count", str(declared_count + 1).encode("ascii")
    )
    patched_block = (
        patched_open
        + block.group("body")
        + yellow_fill
        + block.group("close")
    )
    patched_styles = styles_raw[: block.start()] + patched_block + styles_raw[block.end() :]
    try:
        patched_root = ET.fromstring(patched_styles)
    except ET.ParseError as error:
        raise SafetyStop("Appended yellow fill produced invalid styles.xml.") from error
    new_id = find_yellow_fill_id(patched_root)
    if new_id < 0:
        raise SafetyStop("Could not find the appended yellow fill in styles.xml.")
    return patched_styles, patched_root, new_id


def style_index(cell: ET.Element, style_count: int) -> int:
    try:
        index = int(cell.get("s", "0"))
    except ValueError as error:
        raise SafetyStop(f"Invalid style index at {cell.get('r')}: {cell.get('s')!r}") from error
    if not 0 <= index < style_count:
        raise SafetyStop(
            f"Style index at {cell.get('r')} is outside cellXfs: {index}"
        )
    return index


def raw_attribute_pattern(name: bytes) -> re.Pattern[bytes]:
    return re.compile(
        rb"(?P<space>\s)"
        + re.escape(name)
        + rb"=(?P<quote>[\"'])(?P<value>.*?)(?P=quote)"
    )


def raw_xml_attribute(tag: bytes, name: bytes) -> bytes | None:
    matches = list(raw_attribute_pattern(name).finditer(tag))
    if len(matches) > 1:
        raise SafetyStop(f"XML tag has duplicate {name.decode('ascii')} attributes.")
    return matches[0].group("value") if matches else None


def set_raw_xml_attribute(tag: bytes, name: bytes, value: bytes) -> bytes:
    """Set one attribute without normalizing any unrelated XML bytes."""
    if not tag.startswith(b"<") or not tag.endswith(b">"):
        raise SafetyStop("Expected an XML start tag while applying a style.")

    pattern = raw_attribute_pattern(name)
    matches = list(pattern.finditer(tag))
    if len(matches) > 1:
        raise SafetyStop(f"XML tag has duplicate {name.decode('ascii')} attributes.")
    if matches:
        return pattern.sub(
            lambda match: (
                match.group("space")
                + name
                + b"="
                + match.group("quote")
                + value
                + match.group("quote")
            ),
            tag,
            count=1,
        )

    suffix = b"/>" if tag.endswith(b"/>") else b">"
    return tag[: -len(suffix)] + b" " + name + b'=\"' + value + b'\"' + suffix


def clone_xf_with_yellow_fill(raw_xf: bytes, yellow_fill_id: int) -> bytes:
    start_end = raw_xf.find(b">") + 1
    if start_end == 0:
        raise SafetyStop("Could not locate the end of a cell style XML start tag.")
    start_tag = raw_xf[:start_end]
    patched_start_tag = set_raw_xml_attribute(
        set_raw_xml_attribute(start_tag, b"fillId", str(yellow_fill_id).encode("ascii")),
        b"applyFill",
        b"1",
    )
    return patched_start_tag + raw_xf[start_end:]


def patch_styles_xml(
    styles_raw: bytes,
    matching_cells: list[ET.Element],
    force_references: tuple[str, ...] = (),
) -> tuple[bytes, dict[str, int], int]:
    """Append yellow style clones while preserving every existing XML byte."""
    try:
        styles_root = ET.fromstring(styles_raw)
    except ET.ParseError as error:
        raise SafetyStop("styles.xml is not well-formed XML.") from error

    cell_xfs = styles_root.find(x("cellXfs"))
    if cell_xfs is None:
        raise SafetyStop("styles.xml has no cellXfs element.")
    original_xfs = cell_xfs.findall(x("xf"))
    if not original_xfs:
        raise SafetyStop("styles.xml has an empty cellXfs element.")

    original_count = len(original_xfs)
    references_by_old_style: dict[int, list[str]] = {}
    already_yellow_cells = 0
    forced_references = set(force_references)
    for cell in matching_cells:
        reference = cell.get("r", "")
        if not reference:
            raise SafetyStop("A matched target cell has no XML reference.")
        old_style = style_index(cell, original_count)
        if style_is_yellow(styles_root, str(old_style)) and reference not in forced_references:
            already_yellow_cells += 1
            continue
        references_by_old_style.setdefault(old_style, []).append(reference)

    if not references_by_old_style:
        return styles_raw, {}, already_yellow_cells

    styles_raw, styles_root, yellow_fill_id = ensure_yellow_fill(
        styles_raw, styles_root
    )

    block = CELL_XFS_BLOCK_RE.search(styles_raw)
    if block is None or CELL_XFS_BLOCK_RE.search(styles_raw, block.end()) is not None:
        raise SafetyStop("styles.xml does not contain exactly one raw cellXfs block.")
    raw_count = raw_xml_attribute(block.group("open"), b"count")
    try:
        declared_count = int((raw_count or b"").decode("ascii"))
    except ValueError as error:
        raise SafetyStop("cellXfs has no valid count attribute.") from error
    raw_xfs = XF_RAW_RE.findall(block.group("body"))
    if declared_count != original_count or len(raw_xfs) != original_count:
        raise SafetyStop(
            "The raw and parsed cellXfs counts disagree; refusing to edit styles.xml."
        )

    replacement_style_ids: dict[int, int] = {}
    clones: list[bytes] = []
    for old_style in sorted(references_by_old_style):
        replacement_style_ids[old_style] = original_count + len(clones)
        clones.append(clone_xf_with_yellow_fill(raw_xfs[old_style], yellow_fill_id))

    patched_open = set_raw_xml_attribute(
        block.group("open"),
        b"count",
        str(original_count + len(clones)).encode("ascii"),
    )
    patched_block = patched_open + block.group("body") + b"".join(clones) + block.group("close")
    patched_styles = (
        styles_raw[: block.start()] + patched_block + styles_raw[block.end() :]
    )
    style_assignments = {
        reference: replacement_style_ids[old_style]
        for old_style, references in references_by_old_style.items()
        for reference in references
    }
    return patched_styles, style_assignments, already_yellow_cells


def excel_column_index(column: str) -> int:
    """Return the one-based numeric index of an Excel column letter."""
    number = 0
    for character in column:
        if not "A" <= character <= "Z":
            raise SafetyStop(f"Invalid Excel column in cell reference: {column!r}")
        number = number * 26 + ord(character) - ord("A") + 1
    if not 1 <= number <= 16384:
        raise SafetyStop(f"Excel column is outside XFD: {column!r}")
    return number


def raw_row_number(row_tag: bytes) -> int:
    match = ROW_NUMBER_ATTR_RE.search(row_tag)
    if match is None:
        raise SafetyStop("A target worksheet row has no valid row number.")
    try:
        return int(match.group("number"))
    except ValueError as error:
        raise SafetyStop("A target worksheet row has an invalid row number.") from error


def insert_cells_into_row(
    row_block: bytes, additions: list[tuple[str, int]]
) -> tuple[bytes, set[str]]:
    """Insert style-only cells into one raw ``<row>`` block in column order."""
    open_end = row_block.find(b">")
    close_start = row_block.rfind(b"</row")
    if open_end < 0 or close_start <= open_end:
        raise SafetyStop("Could not parse a target worksheet row block.")
    body = row_block[open_end + 1 : close_start]
    extension_match = re.search(rb"<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?extLst\b", body)
    extension_start = extension_match.start() if extension_match else len(body)

    existing_refs: set[str] = set()
    existing_positions: list[tuple[int, int]] = []
    for match in CELL_START_TAG_RE.finditer(body):
        reference_match = CELL_REFERENCE_ATTR_RE.search(match.group(0))
        if reference_match is None:
            continue
        reference = reference_match.group("reference").decode("ascii")
        if reference in existing_refs:
            raise SafetyStop(f"Target sheet contains duplicate cell reference: {reference}")
        existing_refs.add(reference)
        column, _ = cell_reference_parts(reference)
        existing_positions.append((match.start(), excel_column_index(column)))

    for reference, _style in additions:
        if reference in existing_refs:
            raise SafetyStop(f"Target sheet already contains cell reference: {reference}")

    inserts: dict[int, list[tuple[int, bytes]]] = {}
    added: set[str] = set()
    for reference, style in sorted(
        additions,
        key=lambda item: (cell_reference_parts(item[0])[1], excel_column_index(cell_reference_parts(item[0])[0])),
    ):
        column, _ = cell_reference_parts(reference)
        column_number = excel_column_index(column)
        insert_at = extension_start
        for position, existing_column_number in existing_positions:
            if position >= extension_start:
                raise SafetyStop(
                    "Target worksheet row contains a cell after its extension list."
                )
            if existing_column_number > column_number:
                insert_at = position
                break
        tag = (
            f'<c r="{reference}" s="{style}"/>'.encode("ascii")
        )
        inserts.setdefault(insert_at, []).append((column_number, tag))
        added.add(reference)

    rebuilt = bytearray()
    cursor = 0
    for position in sorted(inserts):
        rebuilt.extend(body[cursor:position])
        for _column_number, tag in sorted(inserts[position], key=lambda item: item[0]):
            rebuilt.extend(tag)
        cursor = position
    rebuilt.extend(body[cursor:])
    return row_block[: open_end + 1] + bytes(rebuilt) + row_block[close_start:], added


def materialize_cells_in_self_closing_row(
    row_tag: bytes, additions: list[tuple[str, int]]
) -> tuple[bytes, set[str]]:
    """Turn an empty raw row into a row containing style-only cells."""
    if not row_tag.endswith(b"/>"):
        raise SafetyStop("Expected a self-closing worksheet row.")
    open_tag = row_tag[:-2] + b">"
    cells = b"".join(
        f'<c r="{reference}" s="{style}"/>'.encode("ascii")
        for reference, style in sorted(
            additions,
            key=lambda item: (cell_reference_parts(item[0])[1], excel_column_index(cell_reference_parts(item[0])[0])),
        )
    )
    return open_tag + cells + b"</row>", {reference for reference, _ in additions}


def column_letters_from_index(number: int) -> str:
    """Convert a one-based Excel column number to its letter form."""
    if not 1 <= number <= 16384:
        raise SafetyStop(f"Excel column number is outside XFD: {number}")
    letters: list[str] = []
    while number:
        number, remainder = divmod(number - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def cell_range_bounds(reference: str) -> tuple[int, int, int, int]:
    """Return min-column, min-row, max-column, max-row for a cell/range."""
    parts = reference.strip().upper().split(":")
    if len(parts) not in (1, 2):
        raise SafetyStop(f"Invalid worksheet range: {reference!r}")
    start_column, start_row = cell_reference_parts(parts[0])
    end_column, end_row = cell_reference_parts(parts[-1])
    start_column_number = excel_column_index(start_column)
    end_column_number = excel_column_index(end_column)
    if start_column_number > end_column_number or start_row > end_row:
        raise SafetyStop(f"Worksheet range is reversed: {reference!r}")
    return start_column_number, start_row, end_column_number, end_row


def patch_sheet_dimension_and_spans(
    sheet_raw: bytes, added_references: set[str]
) -> bytes:
    """Keep worksheet dimension metadata consistent with materialized cells.

    ``dimension`` and row ``spans`` are technically optimization hints, but
    Excel/WPS use them while loading large sheets.  When a configured mark
    column is outside the original used range, stale hints can hide the newly
    inserted cells until a repair/recalculation pass.  Update only metadata
    affected by the approved blank-cell additions.
    """
    if not added_references:
        return sheet_raw

    added_bounds = [cell_range_bounds(reference) for reference in added_references]
    dimension_match = DIMENSION_TAG_RE.search(sheet_raw)
    if dimension_match is not None:
        dimension_tag = dimension_match.group(0)
        raw_dimension = raw_xml_attribute(dimension_tag, b"ref")
        if raw_dimension is None:
            raise SafetyStop("Worksheet dimension has no ref attribute.")
        try:
            existing_bounds = cell_range_bounds(raw_dimension.decode("ascii"))
        except (UnicodeDecodeError, SafetyStop) as error:
            raise SafetyStop("Worksheet dimension has an invalid ref range.") from error
        bounds = (
            min(existing_bounds[0], *(item[0] for item in added_bounds)),
            min(existing_bounds[1], *(item[1] for item in added_bounds)),
            max(existing_bounds[2], *(item[2] for item in added_bounds)),
            max(existing_bounds[3], *(item[3] for item in added_bounds)),
        )
        updated_ref = (
            f"{column_letters_from_index(bounds[0])}{bounds[1]}:"
            f"{column_letters_from_index(bounds[2])}{bounds[3]}"
        ).encode("ascii")
        updated_dimension_tag = set_raw_xml_attribute(
            dimension_tag, b"ref", updated_ref
        )
        sheet_raw = (
            sheet_raw[: dimension_match.start()]
            + updated_dimension_tag
            + sheet_raw[dimension_match.end() :]
        )

    additions_by_row: dict[int, list[int]] = {}
    for reference in added_references:
        column, row_number = cell_reference_parts(reference)
        additions_by_row.setdefault(row_number, []).append(excel_column_index(column))

    def replace_row_open(match: re.Match[bytes]) -> bytes:
        row_tag = match.group(0)
        row_number = raw_row_number(row_tag)
        added_columns = additions_by_row.get(row_number)
        if not added_columns:
            return row_tag

        # The caller has already materialized additions, so retaining an
        # existing spans range is sufficient when present; otherwise derive
        # the range from the newly added cells.
        spans_value = raw_xml_attribute(row_tag, b"spans")
        if spans_value is not None:
            try:
                span_parts = spans_value.decode("ascii").split(":")
                if len(span_parts) != 2:
                    raise ValueError
                span_start, span_end = (int(item) for item in span_parts)
                if not 1 <= span_start <= span_end <= 16384:
                    raise ValueError
            except (UnicodeDecodeError, ValueError) as error:
                raise SafetyStop(
                    f"Worksheet row {row_number} has an invalid spans attribute."
                ) from error
            span_start = min(span_start, *added_columns)
            span_end = max(span_end, *added_columns)
        else:
            span_start = min(added_columns)
            span_end = max(added_columns)
        return set_raw_xml_attribute(
            row_tag,
            b"spans",
            f"{span_start}:{span_end}".encode("ascii"),
        )

    # Row opening tags are matched after cell insertion.  This handles both
    # ordinary rows and rows that were converted from self-closing form.
    return ROW_OPEN_TAG_RE.sub(replace_row_open, sheet_raw)


def patch_sheet_cell_styles(
    sheet_raw: bytes,
    style_assignments: dict[str, int],
    missing_references: tuple[str, ...] = (),
) -> tuple[bytes, set[str]]:
    """Style approved cells and materialize approved blank target cells."""
    seen: set[str] = set()

    def replace_cell(match: re.Match[bytes]) -> bytes:
        tag = match.group(0)
        reference_match = CELL_REFERENCE_ATTR_RE.search(tag)
        if reference_match is None:
            return tag
        reference = reference_match.group("reference").decode("ascii")
        if reference not in style_assignments:
            return tag
        if reference in seen:
            raise SafetyStop(f"Target sheet contains duplicate cell reference: {reference}")
        seen.add(reference)
        return set_raw_xml_attribute(
            tag,
            b"s",
            str(style_assignments[reference]).encode("ascii"),
        )

    patched_sheet = CELL_START_TAG_RE.sub(replace_cell, sheet_raw)
    missing = set(style_assignments).difference(seen)
    expected_missing = set(missing_references)
    if missing != expected_missing:
        raise SafetyStop(
            "Approved target cells missing from raw worksheet XML do not match "
            f"the planned blank cells: actual={sorted(missing)}, "
            f"expected={sorted(expected_missing)}"
        )
    if not missing:
        return patched_sheet, seen

    additions_by_row: dict[int, list[tuple[str, int]]] = {}
    for reference in sorted(missing):
        _column, row_number = cell_reference_parts(reference)
        additions_by_row.setdefault(row_number, []).append(
            (reference, style_assignments[reference])
        )

    inserted: set[str] = set()

    def replace_row(match: re.Match[bytes]) -> bytes:
        row_block = match.group(0)
        row_number = raw_row_number(row_block[: row_block.find(b">") + 1])
        additions = additions_by_row.get(row_number)
        if not additions:
            return row_block
        patched_row, added = insert_cells_into_row(row_block, additions)
        inserted.update(added)
        return patched_row

    patched_sheet = ROW_BLOCK_RE.sub(replace_row, patched_sheet)

    def replace_empty_row(match: re.Match[bytes]) -> bytes:
        row_tag = match.group(0)
        row_number = raw_row_number(row_tag)
        additions = additions_by_row.get(row_number)
        if not additions:
            return row_tag
        patched_row, added = materialize_cells_in_self_closing_row(row_tag, additions)
        inserted.update(added)
        return patched_row

    patched_sheet = ROW_SELF_CLOSING_RE.sub(replace_empty_row, patched_sheet)
    if inserted != missing:
        raise SafetyStop(
            "Approved blank target cells were not found in raw worksheet rows: "
            f"{sorted(missing - inserted)}"
        )
    patched_sheet = patch_sheet_dimension_and_spans(patched_sheet, inserted)
    seen.update(inserted)
    return patched_sheet, seen


def write_package_copy(
    input_path: Path, temporary_path: Path, replacements: dict[str, bytes]
) -> None:
    with zipfile.ZipFile(input_path, "r") as source:
        names = source.namelist()
        if len(names) != len(set(names)):
            raise SafetyStop("Target XLSX has duplicate ZIP member names; refusing to copy.")
        missing = set(replacements).difference(names)
        if missing:
            raise SafetyStop(f"Replacement parts are absent from target XLSX: {sorted(missing)}")

        with zipfile.ZipFile(
            temporary_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            allowZip64=True,
        ) as output:
            for info in source.infolist():
                payload = replacements.get(info.filename, source.read(info.filename))
                output.writestr(copy.copy(info), payload)


def cell_payload_signature(sheet_root: ET.Element) -> dict[str, tuple[object, ...]]:
    """All cell content except the style attribute; used for output verification."""
    signature: dict[str, tuple[object, ...]] = {}
    for cell in sheet_root.iter(x("c")):
        reference = cell.get("r", "")
        attributes = tuple(
            sorted((key, value) for key, value in cell.attrib.items() if key != "s")
        )
        children = tuple(ET.tostring(child, encoding="utf-8") for child in list(cell))
        signature[reference] = (attributes, children)
    return signature


def cell_style_map(sheet_root: ET.Element) -> dict[str, str]:
    return {
        cell.get("r", ""): cell.get("s", "0")
        for cell in sheet_root.iter(x("c"))
    }


def style_is_yellow(styles_root: ET.Element, style: str) -> bool:
    cell_xfs = styles_root.find(x("cellXfs"))
    fills = styles_root.find(x("fills"))
    if cell_xfs is None or fills is None:
        return False
    xfs = cell_xfs.findall(x("xf"))
    fill_nodes = fills.findall(x("fill"))
    try:
        fill_id = int(xfs[int(style)].get("fillId", "0"))
    except (IndexError, ValueError):
        return False
    return 0 <= fill_id < len(fill_nodes) and is_solid_yellow(fill_nodes[fill_id])


def verify_patched_xml(
    before_sheet_raw: bytes,
    after_sheet_raw: bytes,
    after_styles_raw: bytes,
    expected_style_assignments: dict[str, int],
    expected_added_references: tuple[str, ...] = (),
) -> None:
    """Verify the planned cell-style patch without needing an output file."""
    try:
        before_sheet = ET.fromstring(before_sheet_raw)
        after_sheet = ET.fromstring(after_sheet_raw)
        after_styles = ET.fromstring(after_styles_raw)
    except ET.ParseError as error:
        raise SafetyStop("The planned XML patch is not well-formed.") from error

    before_payload = cell_payload_signature(before_sheet)
    after_payload = cell_payload_signature(after_sheet)
    expected_added = set(expected_added_references)
    added_references = set(after_payload).difference(before_payload)
    removed_references = set(before_payload).difference(after_payload)
    if removed_references or added_references != expected_added:
        raise SafetyStop(
            "Target cell additions/removals do not equal the approved blank cells: "
            f"added={sorted(added_references)}, removed={sorted(removed_references)}, "
            f"expected_added={sorted(expected_added)}"
        )
    for reference in expected_added:
        attributes, children = after_payload[reference]
        if attributes != (("r", reference),) or children:
            raise SafetyStop(
                f"Materialized blank cell {reference} contains a value, formula, "
                "or non-style attribute."
            )
    common_references = set(before_payload).intersection(after_payload)
    if any(before_payload[reference] != after_payload[reference] for reference in common_references):
        raise SafetyStop("A target cell value, formula, or non-style attribute changed.")

    before_styles = cell_style_map(before_sheet)
    after_styles_by_cell = cell_style_map(after_sheet)
    actually_restyled = {
        reference
        for reference in set(before_styles) | set(after_styles_by_cell)
        if before_styles.get(reference, "0") != after_styles_by_cell.get(reference, "0")
    }
    if actually_restyled != set(expected_style_assignments):
        raise SafetyStop(
            "Cells restyled by the XML patch do not equal the approved matched cells: "
            f"actual={sorted(actually_restyled)} "
            f"expected={sorted(expected_style_assignments)}"
        )

    non_yellow = [
        reference
        for reference in actually_restyled
        if (
            after_styles_by_cell.get(reference)
            != str(expected_style_assignments[reference])
            or not style_is_yellow(after_styles, after_styles_by_cell[reference])
        )
    ]
    if non_yellow:
        raise SafetyStop(f"Some restyled cells are not solid yellow: {non_yellow}")


def verify_output(
    input_path: Path,
    output_path: Path,
    replacements: dict[str, bytes],
    target_sheet_part: str,
    expected_style_assignments: dict[str, int],
    expected_added_references: tuple[str, ...] = (),
) -> list[str]:
    with zipfile.ZipFile(input_path, "r") as before, zipfile.ZipFile(
        output_path, "r"
    ) as after:
        invalid_member = after.testzip()
        if invalid_member is not None:
            raise SafetyStop(f"Output ZIP is corrupt at member: {invalid_member}")
        if before.namelist() != after.namelist():
            raise SafetyStop("Output ZIP member list differs from the target input.")

        expected_changed_parts = {
            name
            for name, replacement in replacements.items()
            if before.read(name) != replacement
        }
        changed_parts = [
            name
            for name in before.namelist()
            if before.read(name) != after.read(name)
        ]
        if set(changed_parts) != expected_changed_parts:
            raise SafetyStop(
                "Output changed unexpected package parts: "
                f"{changed_parts}; expected only {sorted(expected_changed_parts)}"
            )
        for name in before.namelist():
            expected_payload = replacements.get(name, before.read(name))
            if after.read(name) != expected_payload:
                raise SafetyStop(f"Output package member differs from its approved bytes: {name}")

        verify_patched_xml(
            before.read(target_sheet_part),
            after.read(target_sheet_part),
            after.read("xl/styles.xml"),
            expected_style_assignments,
            expected_added_references,
        )

    return changed_parts


def require_safe_paths(config: MatchConfig) -> None:
    for label, path in (
        ("Source input", config.source.path),
        ("Target input", config.target.path),
    ):
        if not path.is_file():
            raise SafetyStop(f"{label} does not exist: {display_path(path)}")
        if path.suffix.lower() != ".xlsx":
            raise SafetyStop(f"{label} must be an .xlsx file: {display_path(path)}")
    if config.output_path.suffix.lower() != ".xlsx":
        raise SafetyStop(f"Output must be an .xlsx file: {display_path(config.output_path)}")

    source_path = config.source.path.resolve()
    target_path = config.target.path.resolve()
    output_path = config.output_path.resolve()
    if source_path == target_path:
        raise SafetyStop("Source and target inputs must be different files.")
    if output_path in {source_path, target_path}:
        raise SafetyStop("Output path must not be either input path.")
    if not config.dry_run and config.output_path.exists():
        raise SafetyStop(
            f"Output already exists; refusing to overwrite it: {display_path(config.output_path)}"
        )


def main(argv: list[str] | None = None) -> None:
    args = argument_parser().parse_args(argv)
    override_names = (
        "target_file",
        "target_sheet",
        "target_columns",
        "mark_columns",
        "source_file",
        "source_sheets",
        "source_sheet",
        "source_columns",
        "target_header_row",
        "source_header_row",
        "normalizers",
        "output",
        "require_all_source_keys",
        "max_unmatched_keys",
    )
    has_overrides = any(
        (value := getattr(args, name)) is not None and value != []
        for name in override_names
    )
    # No arguments is the convenient reusable path. Explicit configuration
    # continues to support unattended runs.
    use_interactive = args.interactive or (
        args.config is None
        and not args.write
        and not args.dry_run
        and not has_overrides
    )
    if use_interactive:
        if args.interactive and has_overrides:
            raise SafetyStop(
                "--interactive cannot be combined with file/column overrides; "
                "answer those values at the prompts instead."
            )
        raw_config, config_base = interactive_config(
            args.config,
            args.write,
            args.dry_run,
        )
    else:
        raw_config, config_base = load_raw_config(args.config)
        apply_cli_overrides(raw_config, args)
    config = build_match_config(raw_config, config_base)
    require_safe_paths(config)

    source_hash_before = sha256(config.source.path)
    target_hash_before = sha256(config.target.path)
    source_values = extract_source_values(config.source, config.normalizers)
    target = analyze_target(
        config.target,
        source_values,
        config.mark_columns,
        config.normalizers,
    )

    print(f"Target file: {display_path(config.target.path)}")
    print(f"Target sheet: {config.target.sheets[0]}")
    print("Source file:", display_path(config.source.path))
    print("Source sheets selected:", ", ".join(source_values.sheet_names))
    print("Source key columns:", ", ".join(config.source.columns))
    print("Normalizers:", ", ".join(config.normalizers))
    print(
        "Source keys: "
        f"{len(source_values.point_keys)} unique complete combinations from "
        f"{source_values.source_key_rows} rows; "
        f"{source_values.source_rows_missing_key_fields} source rows skipped for "
        "missing key fields."
    )
    print(
        "Target key columns resolved: "
        f"{', '.join(target.headers.columns)}; header row={target.headers.header_row}."
    )
    print("Target mark columns resolved:", ", ".join(target.mark_columns))
    if target.missing_mark_references:
        print(
            "Blank mark cells to materialize: "
            f"{len(target.missing_mark_references)} "
            "(they will receive style only; values and formulas stay empty)."
        )
    print(
        f"Exact row matches: {len(target.matching_row_numbers)} rows; "
        f"marked target cells = {len(target.matching_cells)}; "
        f"{target.target_rows_missing_key_fields} target rows skipped for missing key fields."
    )

    unmatched_source_keys = sorted(source_values.point_keys - target.target_keys)
    if unmatched_source_keys:
        print(
            f"Source keys with no exact target row: {len(unmatched_source_keys)} "
            f"(showing at most {config.max_unmatched_keys_to_print})."
        )
        for key in unmatched_source_keys[: config.max_unmatched_keys_to_print]:
            sheet_name, row_number = source_values.key_locations[key]
            print(
                "UNMATCHED SOURCE KEY: "
                f"key={key!r}, source={sheet_name}!{row_number}"
            )
        if not config.require_all_source_keys_to_match:
            print(
                "Continuing with exact matches only: the unmatched source keys "
                "will not be highlighted."
            )

    if not target.matching_cells:
        raise SafetyStop("No exact matches were found; no output will be created.")
    if config.require_all_source_keys_to_match and unmatched_source_keys:
        raise SafetyStop(
            "Not every source key exists in the target. Resolve the printed "
            "differences before generating a partial yellow-marking workbook."
        )
    with zipfile.ZipFile(config.target.path, "r") as archive:
        try:
            styles_raw = archive.read("xl/styles.xml")
            target_sheet_raw = archive.read(target.sheet_part)
        except KeyError as error:
            raise SafetyStop(
                f"Target XLSX is missing required package part: {error.args[0]}"
            ) from error
    patched_styles, style_assignments, already_yellow_cells = patch_styles_xml(
        styles_raw,
        target.matching_cells,
        target.missing_mark_references,
    )
    patched_sheet, changed_references = patch_sheet_cell_styles(
        target_sheet_raw,
        style_assignments,
        target.missing_mark_references,
    )
    verify_patched_xml(
        target_sheet_raw,
        patched_sheet,
        patched_styles,
        style_assignments,
        target.missing_mark_references,
    )
    if sha256(config.source.path) != source_hash_before:
        raise SafetyStop("Source input changed while the script was running.")
    if sha256(config.target.path) != target_hash_before:
        raise SafetyStop("Target input changed while the script was running.")
    if config.dry_run:
        print(
            "DRY_RUN=True: in-memory style validation passed; "
            "no Excel file was created or modified."
        )
        print(
            f"Validated yellow cells: {len(style_assignments) + already_yellow_cells}; "
            f"new style assignments: {len(style_assignments)}."
        )
        return
    if not style_assignments:
        raise SafetyStop("Every matched cell is already yellow; no output is needed.")
    if config.output_path.exists():
        raise SafetyStop(
            f"Output already exists; refusing to overwrite it: {display_path(config.output_path)}"
        )

    replacements = {
        "xl/styles.xml": patched_styles,
        target.sheet_part: patched_sheet,
    }
    config.output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = config.output_path.with_name(
        f"{config.output_path.stem}.__temporary__{config.output_path.suffix}"
    )
    if temporary_path.exists():
        raise SafetyStop(
            f"Temporary output already exists; refusing to overwrite it: {display_path(temporary_path)}"
        )

    try:
        write_package_copy(config.target.path, temporary_path, replacements)
        changed_parts = verify_output(
            config.target.path,
            temporary_path,
            replacements,
            target.sheet_part,
            style_assignments,
            target.missing_mark_references,
        )
        if sha256(config.source.path) != source_hash_before:
            raise SafetyStop("Source input changed while the script was running.")
        if sha256(config.target.path) != target_hash_before:
            raise SafetyStop("Target input changed while the script was running.")

        os.replace(temporary_path, config.output_path)
        print(f"Created: {display_path(config.output_path)}")
        print(f"Changed package parts only: {', '.join(changed_parts)}")
        print(f"Restyled yellow cells: {len(changed_references)}")
        print(f"Matched cells already yellow before this run: {already_yellow_cells}")
    except Exception:
        if temporary_path.exists():
            temporary_path.unlink()
        raise


if __name__ == "__main__":
    try:
        main()
    except (SafetyStop, OSError, zipfile.BadZipFile) as error:
        print(f"STOPPED SAFELY: {error}")
        raise SystemExit(2)
