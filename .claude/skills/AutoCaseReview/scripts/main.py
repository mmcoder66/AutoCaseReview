"""Unified CLI entry for AutoCaseReview.

Examples
--------
# Generate everything (SVN Excel + SVN Word + email Word). Email uses the
# single iteration detected in the data (or --iteration).
python3 main.py --mode all

# Only the SVN deliverables, merging every iteration under requirement_data/.
python3 main.py --mode svn

# Email only for files whose names contain the requested SP keyword; iteration is inferred.
python3 main.py --mode email --data-file-keyword SP8
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Allow running both as ``python main.py`` and ``python -m main`` from
# the scripts directory; sibling modules are imported directly.
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from data_loader import (  # noqa: E402
    OUTPUTS_DIR,
    collect_participants,
    get_default_product,
    get_output_subdir,
    iter_requirements,
    list_iterations,
    load_all_requirements,
    load_content_rules,
    render_meeting_name,
    resolve_date_by_strategy,
    sanitize_filename,
)
from generate_email_word import generate as gen_email  # noqa: E402
from generate_svn_excel import generate as gen_excel  # noqa: E402
from generate_svn_word import generate as gen_word  # noqa: E402
from validate_data import print_validation_result, validate_requirements  # noqa: E402


def _normalise_version(version: str | None) -> str | None:
    """Prefix plain numeric versions with "v" for SVN filenames."""
    if not version:
        return version
    text = version.strip()
    if text[:1].isdigit():
        return f"v{text}"
    return text


def _clear_output_subdirs(args, file_types: list[str]) -> None:
    """Clear generated files for the selected output categories."""
    templates = None
    base_dir = Path(args.output_dir) if args.output_dir else OUTPUTS_DIR
    for file_type in file_types:
        subdir = get_output_subdir(file_type, templates)
        target_dir = base_dir / subdir if subdir else base_dir
        if not target_dir.exists():
            target_dir.mkdir(parents=True, exist_ok=True)
            continue
        for path in target_dir.iterdir():
            if path.is_file():
                path.unlink()


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="AutoCaseReview",
        description="Generate 测试用例评审 deliverables (SVN Excel + SVN Word + email Word).",
    )
    parser.add_argument("--mode", choices=("all", "svn", "email"), default="all",
                        help="all = SVN Excel + SVN Word + email Word; "
                             "svn = only SVN Excel + Word (cross-iteration); "
                             "email = only email Word (single iteration).")
    parser.add_argument("--iteration", default=None,
                        help="所属迭代 filter.  If omitted in email mode, the "
                             "single iteration is inferred from loaded data.")
    parser.add_argument("--product", default=get_default_product())
    # Meeting metadata shared by the SVN Word + Excel templates.
    parser.add_argument("--meeting-name", default=None,
                        help="Override 会议名称 template from content_rules.yaml.")
    parser.add_argument("--meeting-place", default=None,
                        help="Override 会议地点 (default from content_rules.yaml::meeting.place).")
    parser.add_argument("--meeting-time", default=None,
                        help="Override 会议时间 (default: day-before 计划完成日期 from content_rules.yaml).")
    parser.add_argument("--recorder", default="", help="记录人员")
    parser.add_argument("--initiator", default="", help="发起人 (Excel)")
    parser.add_argument("--host", default="", help="主持人 (Excel)")
    parser.add_argument("--reviewer", default="", help="评审人 (Excel)")
    parser.add_argument("--review-date", default="", help="评审时间 (Excel)")
    parser.add_argument("--participants", default=None,
                        help="Comma-separated; auto-derived from data if omitted (Word).")
    parser.add_argument("--version", default=None,
                        help="Release/version tag appended to every SVN filename "
                             "(e.g. 1.0.1 -> v1.0.1, SP7).  Required for --mode svn|all; "
                             "the skill asks the user interactively before running.")

    parser.add_argument("--data-dir", default=None,
                        help="Override requirement_data directory.")
    parser.add_argument("--data-file-keyword", default=None,
                        help="Only load requirement_data .xlsx files whose "
                             "filenames contain this keyword, e.g. SP7.")
    parser.add_argument("--output-dir", default=None)
    parser.add_argument("--validate-only", action="store_true",
                        help="Validate input data and exit without generating files.")
    parser.add_argument("--no-clear-output", action="store_true",
                        help="Do not clear target output subdirectories before generation.")
    return parser


def _resolve_email_iteration(df, requested: str | None) -> str:
    iterations = list_iterations(df)
    if not iterations:
        raise SystemExit("no iteration detected in data; cannot build email")
    if requested and requested not in iterations:
        raise SystemExit(
            f"iteration {requested!r} not found; available: {', '.join(iterations)}"
        )
    return requested or iterations[0]


def _run_email(args, df_all) -> Path:
    iteration = _resolve_email_iteration(df_all, args.iteration)
    df_email = df_all[df_all["所属迭代"].str.strip() == iteration.strip()].reset_index(drop=True)
    return gen_email(
        df_email,
        product=args.product,
        iteration=iteration,
        output_dir=args.output_dir,
    )


def _run_svn(args, df_all) -> list[Path]:
    """Generate one Excel + one Word *per requirement* (ID).

    All per-requirement fields (meeting name / time / participants) and all
    todo defaults (status / note) come from ``content_rules.yaml``; CLI flags
    like ``--meeting-time`` / ``--participants`` override them.
    """
    rules = load_content_rules()
    meeting_cfg = rules["meeting"]
    participants_cfg = rules["participants"]

    outputs: list[Path] = []
    for req_id, req_title, df_row in iter_requirements(df_all):
        subject = sanitize_filename(req_title) or req_id
        row = df_row.iloc[0]

        # 会议名称：渲染模板（默认 "{title} 测试用例评审会议纪要"）
        meeting_name = args.meeting_name or render_meeting_name(
            meeting_cfg["name_template"],
            title=req_title,
            req_id=req_id,
            iteration=str(row.get("所属迭代", "")).strip(),
            product=args.product,
        )

        # 会议地点：默认从 config；--meeting-place 显式覆盖
        meeting_place = args.meeting_place or meeting_cfg.get("place", "线上会议")

        # 会议时间：按 strategy 计算；--meeting-time 显式覆盖
        if args.meeting_time:
            meeting_time = args.meeting_time
        else:
            meeting_time = resolve_date_by_strategy(meeting_cfg.get("time", {}), row)

        # 参会人员：默认从配置的列里取；--participants 显式覆盖
        if args.participants is not None:
            participants = args.participants
        else:
            participants = collect_participants(
                df_row,
                roles=participants_cfg.get("source_columns"),
                separator=participants_cfg.get("separator", "、"),
            )
        recorder = args.recorder or str(row.get("测试", "")).strip()

        excel_path = gen_excel(
            df_row,
            product=args.product,
            version=args.version,
            initiator=args.initiator,
            host=args.host,
            reviewer=args.reviewer,
            review_date=args.review_date,
            subject=subject,
            output_dir=args.output_dir,
        )
        word_path = gen_word(
            df_row,
            product=args.product,
            version=args.version,
            meeting_name=meeting_name,
            meeting_place=meeting_place,
            meeting_time=meeting_time,
            recorder=recorder,
            participants=participants,
            subject=subject,
            output_dir=args.output_dir,
        )
        outputs.extend([excel_path, word_path])
    return outputs


def main() -> int:
    args = _build_parser().parse_args()

    if args.mode in ("svn", "all") and not args.version:
        raise SystemExit(
            "--version is required when --mode is svn or all.  "
            "Ask the user for the current release/version tag before running."
        )
    args.version = _normalise_version(args.version)

    df_all = load_all_requirements(
        args.data_dir,
        data_file_keyword=args.data_file_keyword,
    )

    print(f"loaded {len(df_all)} rows from requirement_data; "
          f"iterations={list_iterations(df_all)}")

    validation = validate_requirements(
        df_all,
        mode=args.mode,
        iteration=args.iteration,
        data_file_keyword=args.data_file_keyword,
        directory=args.data_dir,
    )
    print_validation_result(validation)
    if not validation.ok:
        return 1
    if args.validate_only:
        return 0

    if not args.no_clear_output:
        if args.mode == "email":
            _clear_output_subdirs(args, ["email_word"])
        elif args.mode == "svn":
            _clear_output_subdirs(args, ["svn_excel", "svn_word"])
        elif args.mode == "all":
            _clear_output_subdirs(args, ["svn_excel", "svn_word", "email_word"])

    outputs: list[Path] = []
    if args.mode in ("svn", "all"):
        outputs.extend(_run_svn(args, df_all))
    if args.mode in ("email", "all"):
        outputs.append(_run_email(args, df_all))

    print("\nGenerated:")
    for p in outputs:
        print(f"  - {p}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
