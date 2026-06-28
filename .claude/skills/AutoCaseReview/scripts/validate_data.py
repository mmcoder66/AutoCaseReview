"""Validate AutoCaseReview input data before generation."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from data_loader import (  # noqa: E402
    CANONICAL_COLUMNS,
    REQUIREMENT_DATA_DIR,
    get_todo_columns,
    list_iterations,
    load_all_requirements,
)


REQUIRED_COLUMNS = [
    "ID",
    "标题",
    "测试",
    "计划完成日期",
    "所属迭代",
]


@dataclass
class ValidationResult:
    errors: list[str]
    warnings: list[str]

    @property
    def ok(self) -> bool:
        return not self.errors


def _source_files(directory: Path | str | None, data_file_keyword: str | None) -> list[Path]:
    target_dir = Path(directory) if directory else REQUIREMENT_DATA_DIR
    if not target_dir.exists():
        return []
    files = sorted(p for p in target_dir.glob("*.xlsx") if not p.name.startswith("~$"))
    if data_file_keyword:
        keyword = data_file_keyword.strip().casefold()
        files = [p for p in files if keyword in p.name.casefold()]
    return files


def _row_label(row) -> str:
    rid = str(row.get("ID", "")).strip() or "(no-id)"
    title = str(row.get("标题", "")).strip() or "(no-title)"
    source = str(row.get("__source_file", "")).strip()
    return f"{source}::{rid} {title}".strip(":")


def validate_requirements(
    df,
    *,
    mode: str,
    iteration: str | None = None,
    data_file_keyword: str | None = None,
    directory: Path | str | None = None,
) -> ValidationResult:
    """Return validation errors and warnings for loaded requirement data."""
    errors: list[str] = []
    warnings: list[str] = []

    files = _source_files(directory, data_file_keyword)
    if not files:
        errors.append("未找到匹配的 requirement_data xlsx 文件。")

    if df.empty:
        errors.append("需求数据为空，无法生成文档。")
        return ValidationResult(errors, warnings)

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        errors.append(f"缺少必填列：{', '.join(missing)}")

    unexpected_missing = [col for col in CANONICAL_COLUMNS if col not in df.columns]
    if unexpected_missing:
        warnings.append(f"缺少可选标准列，已按空值处理：{', '.join(unexpected_missing)}")

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            continue
        blank_rows = df[df[col].astype(str).str.strip() == ""]
        for _, row in blank_rows.iterrows():
            errors.append(f"必填字段 `{col}` 为空：{_row_label(row)}")

    iterations = list_iterations(df)
    if mode == "email":
        if iteration and iteration not in iterations:
            errors.append(
                f"指定迭代 `{iteration}` 不在数据中；可用迭代：{', '.join(iterations) or '(无)'}"
            )
        target_iterations = [iteration] if iteration else iterations
        if len(target_iterations) != 1:
            errors.append(
                "邮件模式必须对应唯一所属迭代；请使用文件名关键字或 --iteration 缩小范围。"
            )

    if mode == "all":
        # all mode includes email generation, so it still needs one email iteration.
        target_iterations = [iteration] if iteration else iterations
        if len(target_iterations) != 1:
            errors.append("all 模式包含邮件生成，必须对应唯一所属迭代。")

    todo_columns = get_todo_columns(df)
    if not todo_columns:
        warnings.append("未检测到 `代办事项N@责任人` 列。")

    malformed_todos = [
        col for col in df.columns
        if str(col).startswith("代办事项") and col not in todo_columns
    ]
    if malformed_todos:
        warnings.append(
            "以下代办列名不会被识别，请使用 `代办事项N@责任人` 格式："
            + ", ".join(map(str, malformed_todos))
        )

    if "测试用例链接" in df.columns:
        blank_links = df[df["测试用例链接"].astype(str).str.strip() == ""]
        for _, row in blank_links.iterrows():
            warnings.append(f"`测试用例链接` 为空：{_row_label(row)}")

    return ValidationResult(errors, warnings)


def print_validation_result(result: ValidationResult) -> None:
    if result.errors:
        print("Validation errors:")
        for item in result.errors:
            print(f"  - {item}")
    if result.warnings:
        print("Validation warnings:")
        for item in result.warnings:
            print(f"  - {item}")
    if result.ok:
        print("Validation OK")


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate AutoCaseReview requirement data.")
    parser.add_argument("--mode", choices=("all", "svn", "email"), default="all")
    parser.add_argument("--iteration", default=None)
    parser.add_argument("--data-dir", default=None)
    parser.add_argument("--data-file-keyword", default=None)
    args = parser.parse_args()

    df = load_all_requirements(
        args.data_dir,
        iteration=args.iteration,
        data_file_keyword=args.data_file_keyword,
    )
    result = validate_requirements(
        df,
        mode=args.mode,
        iteration=args.iteration,
        data_file_keyword=args.data_file_keyword,
        directory=args.data_dir,
    )
    print_validation_result(result)
    return 0 if result.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
