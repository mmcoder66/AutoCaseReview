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
REQUIREMENT_DATA_DIR = INPUTS_DIR / "requirement_data"
OUTPUTS_DIR = SKILL_ROOT / "outputs"
CONFIG_DIR = SKILL_ROOT / "config"
FILENAME_TEMPLATES_PATH = CONFIG_DIR / "filename_templates.yaml"
CONTENT_RULES_PATH = CONFIG_DIR / "content_rules.yaml"

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
    "代办事项1@责任人",
    "代办事项2@责任人",
    "代办事项3@责任人",
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
TODO_COLUMNS = [
    "代办事项1@责任人",
    "代办事项2@责任人",
    "代办事项3@责任人",
]

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
    return df[CANONICAL_COLUMNS]


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
) -> pd.DataFrame:
    """Load and concatenate every xlsx in *directory*.

    Parameters
    ----------
    directory:
        Defaults to ``inputs/requirement_data``.  ``.xlsx`` files that start
        with ``~$`` (Excel lock files) are skipped.
    iteration:
        Optional filter on ``所属迭代`` (exact match after strip).
    """
    target_dir = Path(directory) if directory else REQUIREMENT_DATA_DIR
    if not target_dir.exists():
        raise FileNotFoundError(f"requirement_data directory not found: {target_dir}")

    files = sorted(p for p in target_dir.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise FileNotFoundError(f"no .xlsx files found under {target_dir}")

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
        for col in TODO_COLUMNS:
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


# ---------------------------------------------------------------------------
# Filename templates (loaded once from config/filename_templates.yaml)
# ---------------------------------------------------------------------------
DEFAULT_FILENAME_TEMPLATES = {
    "svn_excel": "LC-SOP-RC-007-R02_测试用例评审_{product}_{title}_{version}.xlsx",
    "svn_word": "LC-SOP-RC-003-M01_会议纪要_{product}_测试用例评审_{title}_{version}.docx",
    "email_word": "{product}_{iteration}_测试用例评审_邮件.docx",
    "subdirs": {
        "svn_excel": "测试用例评审",
        "svn_word": "会议纪要",
        "email_word": "邮件",
    },
}


class _SafeFormatDict(dict):
    """Dict that returns "" for missing keys, so str.format_map never fails."""

    def __missing__(self, key):  # noqa: D401 - simple sentinel
        return ""


def load_filename_templates(path: Path | str | None = None) -> dict:
    """Load filename templates (and output subdirs) from YAML.

    Missing keys fall back to :data:`DEFAULT_FILENAME_TEMPLATES`.  The
    ``subdirs`` value is a nested dict and is merged key-by-key rather than
    stringified.
    """
    target = Path(path) if path else FILENAME_TEMPLATES_PATH
    result = dict(DEFAULT_FILENAME_TEMPLATES)
    if not target.exists():
        return result
    try:
        import yaml  # type: ignore
    except ImportError:
        return result
    try:
        with target.open(encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
    except (OSError, yaml.YAMLError):
        return result
    for key in result:
        if key not in data or data[key] is None:
            continue
        if isinstance(result[key], dict):
            # Merge nested dict (e.g. subdirs) instead of replacing.
            merged = dict(result[key])
            if isinstance(data[key], dict):
                merged.update(data[key])
            result[key] = merged
        else:
            result[key] = str(data[key])
    return result


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


# ---------------------------------------------------------------------------
# Content rules (loaded from config/content_rules.yaml)
# ---------------------------------------------------------------------------
DEFAULT_CONTENT_RULES = {
    "meeting": {
        "name_template": "{title} 测试用例评审会议纪要",
        "place": "线上会议",
        "time": {"strategy": "day_before", "source": "计划完成日期", "explicit_value": ""},
    },
    "participants": {
        "source_columns": ["测试", "前端开发", "后端开发", "创建者"],
        "separator": "、",
    },
    "todo": {
        "plan_date": {"strategy": "day_before", "source": "计划完成日期", "explicit_value": ""},
        "status": "已完成",
        "note": "",
    },
}


def _deep_merge(defaults: dict, overrides: dict) -> dict:
    """Recursively merge *overrides* into *defaults* (defaults win on missing keys)."""
    out = dict(defaults)
    for key, value in (overrides or {}).items():
        if isinstance(value, dict) and isinstance(out.get(key), dict):
            out[key] = _deep_merge(out[key], value)
        elif value is not None:
            out[key] = value
    return out


def load_content_rules(path: Path | str | None = None) -> dict:
    """Load content rules from YAML, with hardcoded fallback defaults.

    Missing sections / keys fall back to :data:`DEFAULT_CONTENT_RULES`, so a
    partially-edited YAML file never breaks generation.
    """
    target = Path(path) if path else CONTENT_RULES_PATH
    try:
        import yaml  # type: ignore
    except ImportError:
        return dict(DEFAULT_CONTENT_RULES)
    if not target.exists():
        return dict(DEFAULT_CONTENT_RULES)
    try:
        with target.open(encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
    except (OSError, yaml.YAMLError):
        return dict(DEFAULT_CONTENT_RULES)
    return _deep_merge(DEFAULT_CONTENT_RULES, data)


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
def derive_meeting_name(df: pd.DataFrame, product: str = "iBatchInsight") -> str:
    iterations = list_iterations(df)
    if len(iterations) == 1:
        suffix = iterations[0]
    elif len(iterations) > 1:
        suffix = "多迭代汇总"
    else:
        suffix = "需求评审"
    return f"{product} {suffix} 测试用例评审会议纪要"


def ensure_outputs_dir() -> Path:
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    return OUTPUTS_DIR
