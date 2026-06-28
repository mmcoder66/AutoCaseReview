"""Shared data loader for AutoCaseReview.

Scans ``inputs/requirement_data/`` for xlsx files, normalises column names,
concatenates them into a single DataFrame, and provides grouping helpers.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

import pandas as pd

# ---------------------------------------------------------------------------
# Path conventions (resolved relative to this file so the loader works from
# any CWD).  Repository layout:
#   .claude/skills/AutoCaseReview/
#       inputs/requirement_data/*.xlsx
#       scripts/data_loader.py        <- this file
# ---------------------------------------------------------------------------
SKILL_ROOT = Path(__file__).resolve().parent.parent
INPUTS_DIR = SKILL_ROOT / "inputs"
TEMPLATES_DIR = INPUTS_DIR / "templates"
REQUIREMENT_DATA_DIR = INPUTS_DIR / "requirement_data"
OUTPUTS_DIR = SKILL_ROOT / "outputs"
CONFIG_DIR = SKILL_ROOT / "config"
FILENAME_TEMPLATES_PATH = CONFIG_DIR / "filename_templates.yaml"
CONTENT_RULES_PATH = CONFIG_DIR / "content_rules.yaml"
PROJECT_CONFIG_PATH = CONFIG_DIR / "project.yaml"
TEMPLATE_CONFIG_PATH = CONFIG_DIR / "templates.yaml"

# Canonical column names used everywhere downstream.
CANONICAL_COLUMNS = [
    "ID",
    "标题",
    "测试",
    "测试用例链接",
    "计划完成日期",
    "前端开发",
    "后端开发",
    "创建者",
    "所属迭代",
]

# Alternative spellings we accept from input files (best-effort normalisation).
COLUMN_ALIASES = {
    "id": "ID",
    "标题": "标题",
    "title": "标题",
    "测试": "测试",
    "tester": "测试",
    "测试用例链接": "测试用例链接",
    "用例链接": "测试用例链接",
    "case_url": "测试用例链接",
    "计划完成日期": "计划完成日期",
    "计划完成": "计划完成日期",
    "前端开发": "前端开发",
    "后端开发": "后端开发",
    "创建者": "创建者",
    "所属迭代": "所属迭代",
    "迭代": "所属迭代",
    "iteration": "所属迭代",
}

# Columns that hold to-do items (each contains free text with @mentions).
# The input can contain any number of columns such as 代办事项1@责任人,
# 代办事项2@责任人, ...; discover them dynamically instead of hard-coding a max.
TODO_COLUMN_RE = re.compile(r"^代办事项(\d+)@责任人$")

# Matches substrings like "@黄美玲" or "@Zhang San".
# Names: 2-4 Chinese chars, or 1-30 ASCII letters/dots/hyphens/apostrophes
# (we intentionally do NOT allow whitespace, so the match stops as soon as a
# space follows the name, e.g. "@黄美玲 修改用例" yields "黄美玲").
MENTION_RE = re.compile(
    r"@([\u4e00-\u9fa5]{2,4}|[A-Za-z][A-Za-z.\-']{0,29})"
)


# ---------------------------------------------------------------------------
# Public helpers
# ---------------------------------------------------------------------------
def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for col in df.columns:
        key = str(col).strip()
        if key in COLUMN_ALIASES:
            rename_map[col] = COLUMN_ALIASES[key]
    df = df.rename(columns=rename_map)

    missing = [c for c in CANONICAL_COLUMNS if c not in df.columns]
    if missing:
        # Add empty columns for optional fields so downstream code can rely on
        # the full schema.  Critical missing fields are reported by caller.
        for col in missing:
            df[col] = ""
    todo_columns = get_todo_columns(df)
    return df[CANONICAL_COLUMNS + todo_columns]


def get_todo_columns(df: pd.DataFrame) -> list[str]:
    """Return all ``代办事项N@责任人`` columns sorted by N, then original order."""
    indexed: list[tuple[int, int, str]] = []
    for position, col in enumerate(df.columns):
        name = str(col).strip()
        match = TODO_COLUMN_RE.match(name)
        if match:
            indexed.append((int(match.group(1)), position, name))
    return [name for _, _, name in sorted(indexed)]


def _coerce_dates(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return ""
        # Ones exports dates as strings like "2026-06-24 00:00:00".  Parse and
        # reformat so downstream cells show "2026-06-24" only.
        try:
            return pd.to_datetime(text).strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            return text
    try:
        return pd.to_datetime(value).strftime("%Y-%m-%d")
    except Exception:
        return str(value)


def previous_day(date_str: str) -> str:
    """Return the day before *date_str* (YYYY-MM-DD).  Empty on failure.

    Used for both ``会议时间`` and ``计划解决日期``, both of which must be the
    day before the requirement's ``计划完成日期``.
    """
    if not date_str:
        return ""
    try:
        return (pd.to_datetime(date_str) - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
    except Exception:
        return date_str


def load_requirement_file(path: Path) -> pd.DataFrame:
    """Load a single xlsx file into a canonicalised DataFrame."""
    # Header is on the first row of every Ones export we have seen so far.
    df = pd.read_excel(path, dtype=str)
    df = _normalise_columns(df)
    df["__source_file"] = path.name
    df["计划完成日期"] = df["计划完成日期"].map(_coerce_dates)
    df = df.fillna("")
    # Drop fully-empty rows (artefacts of Ones export trailing blanks).
    df = df[df["ID"].str.strip().astype(bool) | df["标题"].str.strip().astype(bool)]
    return df.reset_index(drop=True)


def load_all_requirements(
    directory: Path | str | None = None,
    iteration: str | None = None,
    data_file_keyword: str | None = None,
) -> pd.DataFrame:
    """Load and concatenate every xlsx in *directory*.

    Parameters
    ----------
    directory:
        Defaults to ``inputs/requirement_data``.  ``.xlsx`` files that start
        with ``~$`` (Excel lock files) are skipped.
    iteration:
        Optional filter on ``所属迭代`` (exact match after strip).
    data_file_keyword:
        Optional case-insensitive filename filter.  For example, ``SP8`` loads
        only requirement files whose names contain ``SP8``.
    """
    target_dir = Path(directory) if directory else REQUIREMENT_DATA_DIR
    if not target_dir.exists():
        raise FileNotFoundError(f"requirement_data directory not found: {target_dir}")

    files = sorted(p for p in target_dir.glob("*.xlsx") if not p.name.startswith("~$"))
    if data_file_keyword:
        keyword = data_file_keyword.strip().casefold()
        files = [p for p in files if keyword in p.name.casefold()]
    if not files:
        suffix = f" matching filename keyword {data_file_keyword!r}" if data_file_keyword else ""
        raise FileNotFoundError(f"no .xlsx files found under {target_dir}{suffix}")

    frames = [load_requirement_file(p) for p in files]
    combined = pd.concat(frames, ignore_index=True)

    if iteration:
        combined = combined[combined["所属迭代"].str.strip() == iteration.strip()].reset_index(drop=True)
    return combined


# ---------------------------------------------------------------------------
# Derived helpers used by generators
# ---------------------------------------------------------------------------
def list_iterations(df: pd.DataFrame) -> list[str]:
    seen: list[str] = []
    for value in df["所属迭代"].tolist():
        value = str(value).strip()
        if value and value not in seen:
            seen.append(value)
    return seen


def collect_participants(
    df: pd.DataFrame,
    extra: Iterable[str] | None = None,
    roles: Iterable[str] | None = None,
    separator: str = "、",
) -> str:
    """Build a de-duplicated participant string from data + extras.

    Parameters
    ----------
    roles:
        Column names whose values are concatenated.  Defaults to the four
        roles used by the test-case review workflow (测试 / 前端开发 / 后端开发 / 创建者).
        When called from ``main.py`` this comes from ``content_rules.yaml``.
    separator:
        Delimiter inserted between names.  Also configurable via YAML.
    """
    role_cols = tuple(roles) if roles else ("测试", "前端开发", "后端开发", "创建者")
    participants: list[str] = []
    for col in role_cols:
        for value in df[col].dropna().tolist():
            for name in re.split(r"[,，、\s/]+", str(value)):
                name = name.strip()
                if name and name not in participants:
                    participants.append(name)
    if extra:
        for name in extra:
            for n in re.split(r"[,，、\s/]+", str(name)):
                n = n.strip()
                if n and n not in participants:
                    participants.append(n)
    return separator.join(participants)


def extract_mentions(text: str) -> list[str]:
    """Return @-mentions found inside *text*, preserving order, no duplicates."""
    out: list[str] = []
    for match in MENTION_RE.finditer(text or ""):
        name = match.group(1).strip()
        if name and name not in out:
            out.append(name)
    return out


def strip_mentions(text: str) -> str:
    """Remove @-mentions from *text* and tidy up leftover whitespace.

    Example: ``"预览不可编辑 @黄美玲 修改用例"`` → ``"预览不可编辑 修改用例"``
    """
    if not text:
        return ""
    cleaned = MENTION_RE.sub("", text)
    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    return cleaned.strip()


def expand_todos(df: pd.DataFrame) -> list[dict]:
    """Flatten ``代办事项N@责任人`` cells into one record per todo item.

    Each returned dict has: ``seq``, ``description`` (original text), ``owners``
    (list of @-mentions), ``plan_date``, ``req_id``, ``req_title``,
    ``iteration``.
    """
    items: list[dict] = []
    seq = 0
    for _, row in df.iterrows():
        for col in get_todo_columns(df):
            text = str(row.get(col, "")).strip()
            if not text:
                continue
            seq += 1
            items.append(
                {
                    "seq": seq,
                    "description": text,
                    "owners": extract_mentions(text),
                    # 原始 计划完成日期；展示成什么由 generate_svn_word + content_rules 决定
                    "plan_date_raw": str(row.get("计划完成日期", "")).strip(),
                    "req_id": str(row.get("ID", "")).strip(),
                    "req_title": str(row.get("标题", "")).strip(),
                    "iteration": str(row.get("所属迭代", "")).strip(),
                }
            )
    return items


# ---------------------------------------------------------------------------
# Per-requirement iteration & filename helpers
# ---------------------------------------------------------------------------
_FILENAME_ILLEGAL_RE = re.compile(r'[\\/:*?"<>|\r\n\t]+')


def sanitize_filename(text: str, max_length: int = 60) -> str:
    """Strip characters that are unsafe in file names on Windows/macOS/Linux.

    Chinese punctuation such as ``【】（）`` is preserved.  The result is
    truncated to *max_length* characters.
    """
    safe = _FILENAME_ILLEGAL_RE.sub("_", str(text or "")).strip()
    safe = re.sub(r"[\s_]+", "_", safe).strip("_. ")
    return safe[:max_length] or "untitled"


def iter_requirements(df: pd.DataFrame):
    """Yield ``(req_id, req_title, single_row_df)`` for every requirement row.

    Used by the SVN generators which produce one Excel + one Word *per
    requirement*.  The yielded DataFrame keeps the original column schema so
    downstream code does not need to change.
    """
    for _, row in df.iterrows():
        rid = str(row.get("ID", "")).strip() or "(no-id)"
        title = str(row.get("标题", "")).strip()
        yield rid, title, df.loc[[row.name]].copy()


class _SafeFormatDict(dict):
    """Dict that returns "" for missing keys, so str.format_map never fails."""

    def __missing__(self, key):  # noqa: D401 - simple sentinel
        return ""


def _load_yaml_mapping(path: Path | str) -> dict:
    """Load a required YAML mapping file.

    YAML files are the single source of truth for defaults.  Missing or invalid
    config should fail early instead of silently falling back to stale Python
    constants.
    """
    target = Path(path)
    if not target.exists():
        raise FileNotFoundError(f"config file not found: {target}")
    try:
        import yaml  # type: ignore
    except ImportError as exc:
        raise RuntimeError("pyyaml is required to load AutoCaseReview config") from exc
    try:
        with target.open(encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
    except (OSError, yaml.YAMLError) as exc:
        raise RuntimeError(f"failed to load config file: {target}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"config file must contain a mapping: {target}")
    return data


def load_project_config(path: Path | str | None = None) -> dict:
    """Load project-level defaults such as product name."""
    return _load_yaml_mapping(Path(path) if path else PROJECT_CONFIG_PATH)


def get_default_product() -> str:
    """Return the configured product name."""
    return str(load_project_config()["product"])


def load_template_names(path: Path | str | None = None) -> dict:
    """Load template filenames from config/templates.yaml."""
    data = _load_yaml_mapping(Path(path) if path else TEMPLATE_CONFIG_PATH)
    required = ("svn_excel", "svn_word", "email_reference")
    missing = [key for key in required if not data.get(key)]
    if missing:
        raise KeyError(f"missing template config keys: {', '.join(missing)}")
    return data


def load_filename_templates(path: Path | str | None = None) -> dict:
    """Load filename templates (and output subdirs) from YAML.

    The YAML file is the single source of truth for output naming.
    """
    target = Path(path) if path else FILENAME_TEMPLATES_PATH
    data = _load_yaml_mapping(target)
    required = ("svn_excel", "svn_word", "email_word", "subdirs")
    missing = [key for key in required if key not in data]
    if missing:
        raise KeyError(f"missing filename template keys: {', '.join(missing)}")
    if not isinstance(data.get("subdirs"), dict):
        raise ValueError("filename_templates.yaml::subdirs must be a mapping")
    return data


def get_output_subdir(file_type: str, templates: dict | None = None) -> str:
    """Return the configured output subdir name for *file_type*.

    *file_type* is one of ``svn_excel`` / ``svn_word`` / ``email_word``.
    Returns ``""`` when no subdir is configured.
    """
    templates = templates if templates is not None else load_filename_templates()
    return str(templates.get("subdirs", {}).get(file_type, "") or "")


def render_filename(template: str, **context) -> str:
    """Render a filename template.

    Unknown placeholders become empty strings.  Every string value is
    sanitised (illegal filename chars replaced, truncated) before insertion.
    """
    safe_context = {
        key: sanitize_filename(str(value)) if isinstance(value, str) else value
        for key, value in context.items()
    }
    return template.format_map(_SafeFormatDict(safe_context))


def load_content_rules(path: Path | str | None = None) -> dict:
    """Load content rules from YAML."""
    target = Path(path) if path else CONTENT_RULES_PATH
    return _load_yaml_mapping(target)


def resolve_date_by_strategy(spec: dict, source) -> str:
    """Resolve a date cell according to a strategy spec from content_rules.

    *source* may be a ``pd.Series`` (full requirement row) or a plain dict;
    in both cases ``source[source_field]`` must return the value.

    Supported strategies:
      * ``day_before`` — take ``source[source_field]`` and return the previous day
      * ``explicit``   — return ``spec["explicit_value"]`` verbatim
    """
    strategy = (spec or {}).get("strategy", "explicit")
    if strategy == "day_before":
        source_field = spec.get("source", "计划完成日期")
        raw = source.get(source_field, "") if hasattr(source, "get") else ""
        return previous_day(str(raw).strip())
    if strategy == "explicit":
        return str(spec.get("explicit_value", ""))
    return ""


def resolve_field_by_strategy(spec: dict, row, *, default: str = "") -> str:
    """Generic resolver for any content field driven by content_rules.yaml.

    Supported strategies:
      * ``column``       — return ``row[source]`` verbatim
      * ``columns_join`` — collect ``row[sources...]`` and de-duplicate using
                           ``separator``
      * ``day_before``   — previous day of ``row[source]``
      * ``explicit``     — return ``spec["explicit_value"]`` verbatim

    Works for both ``pd.Series`` (single requirement row) and dict inputs.
    """
    if not spec:
        return default
    strategy = spec.get("strategy", "explicit")

    if strategy == "column":
        source_field = spec.get("source", "")
        raw = row.get(source_field, "") if hasattr(row, "get") else ""
        return str(raw).strip()

    if strategy == "columns_join":
        sources = spec.get("sources", []) or []
        separator = spec.get("separator", "、")
        # Build a single-row DataFrame so collect_participants keeps working.
        if hasattr(row, "to_frame"):
            df_row = row.to_frame().T
        else:
            df_row = pd.DataFrame([row])
        return collect_participants(df_row, roles=sources, separator=separator)

    if strategy == "day_before":
        return resolve_date_by_strategy(spec, row)

    if strategy == "explicit":
        return str(spec.get("explicit_value", ""))

    return default


def render_meeting_name(template: str, **context) -> str:
    """Render the meeting name (no filename sanitisation — keep full title)."""
    return template.format_map(_SafeFormatDict(context))


# ---------------------------------------------------------------------------
# Default metadata used when the caller doesn't pass explicit values.
# ---------------------------------------------------------------------------
def derive_meeting_name(df: pd.DataFrame, product: str | None = None) -> str:
    product_value = product or get_default_product()
    iterations = list_iterations(df)
    if len(iterations) == 1:
        suffix = iterations[0]
    elif len(iterations) > 1:
        suffix = "多迭代汇总"
    else:
        suffix = "需求评审"
    return f"{product_value} {suffix} 测试用例评审会议纪要"


def ensure_outputs_dir() -> Path:
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    return OUTPUTS_DIR
