"""Generate the SVN Excel deliverable (LC-SOP-RC-007-R02).

Strategy: copy the template xlsx **byte-for-byte**, then patch ONLY the
specific data cells in ``测试用例评审记录`` (sheet3) using direct XML edit.

This bypasses openpyxl entirely, which would otherwise drop:
  * ``xl/embeddings/*.docx`` (signature Word docs on 封皮 / 签名页)
  * ``xl/media/*.png`` / ``*.emf`` (logos & drawings)
  * ``xl/drawings/*.vml`` + ``drawing*.xml`` (VML / drawings)

After patching, the cover & signature pages remain visually identical to the
template — embedded objects, images and formatting all preserved.
"""

from __future__ import annotations

import argparse
import re
import zipfile
import xml.sax.saxutils as saxutils
from pathlib import Path

from data_loader import (
    INPUTS_DIR,
    OUTPUTS_DIR,
    ensure_outputs_dir,
    get_output_subdir,
    list_iterations,
    load_content_rules,
    load_filename_templates,
    render_filename,
    resolve_field_by_strategy,
)

TEMPLATE_NAME = "LC-SOP-RC-007-R02_测试用例评审_iBatchInsight_需求名称_版本号.xlsx"
TARGET_SHEET_PATH = "xl/worksheets/sheet3.xml"   # 测试用例评审记录

# Cell coordinates we patch in 测试用例评审记录.
CELL_PRODUCT = "C3"          # merged C3:E3
CELL_ONES_REQS = "C4"        # merged C4:E4
CELL_CASE_LINKS = "C5"       # merged C5:E5
CELL_INITIATOR = "C6"
CELL_HOST = "E6"
CELL_REVIEWER = "C7"
CELL_REVIEW_DATE = "E7"


# ---------------------------------------------------------------------------
# XML patching helpers
# ---------------------------------------------------------------------------
def _build_inline_cell(coord: str, style: str, value: str) -> str:
    """Build an inline-string ``<c>`` element that replaces whatever was there.

    Preserves the cell's original style index (``s="N"``) so fonts / borders /
    fill / wrap-text all survive.
    """
    escaped = saxutils.escape(value or "", {'"': "&quot;"})
    # Encode \n as &#10; so Excel renders multi-line text inside the cell.
    escaped = escaped.replace("\n", "&#10;")
    style_attr = f' s="{style}"' if style else ""
    return (
        f'<c r="{coord}"{style_attr} t="inlineStr">'
        f"<is><t xml:space=\"preserve\">{escaped}</t></is></c>"
    )


def _extract_cell_style(sheet_xml: str, coord: str) -> str:
    """Return the existing ``s="N"`` value for *coord*, or ``""`` if absent."""
    m = re.search(r'<c r="' + coord + r'"[^>]*?\bs="(\d+)"', sheet_xml)
    return m.group(1) if m else ""


def _patch_sheet_cells(sheet_xml: str, cell_values: dict[str, str]) -> str:
    """Replace specific cells in *sheet_xml* with inline-string values.

    *cell_values* maps cell coordinates (e.g. ``"C4"``) to their new text.
    Each cell becomes an ``inlineStr`` cell — no need to touch sharedStrings.
    """
    out = sheet_xml
    for coord, value in cell_values.items():
        style = _extract_cell_style(out, coord)
        new_cell = _build_inline_cell(coord, style, value)
        # Match the entire <c r="X" .../> or <c r="X" ...>...</c> element.
        # The negative class [^/>] prevents crossing into the next element
        # boundary; the alternation handles both self-closing & paired forms.
        pattern = re.compile(
            r'<c r="' + coord + r'"[^/>]*?(?:/>|>.*?</c>)',
            re.DOTALL,
        )
        new_out, n = pattern.subn(lambda _: new_cell, out, count=1)
        if n == 0:
            raise RuntimeError(
                f"cell {coord} not found in {TARGET_SHEET_PATH}; "
                "is the template the right one?"
            )
        out = new_out
    return out


def _write_patched_xlsx(template_path: Path, output_path: Path,
                        cell_values: dict[str, str]) -> None:
    """Copy *template_path* to *output_path*, patching only target cells.

    Every zip member other than ``TARGET_SHEET_PATH`` is copied verbatim,
    preserving order and binary content (images / embeddings / drawings).
    """
    with zipfile.ZipFile(template_path, "r") as zin:
        # Capture (filename, bytes) in template's own ordering.
        entries = [(info.filename, zin.read(info.filename)) for info in zin.infolist()]

    patched = False
    for i, (name, data) in enumerate(entries):
        if name == TARGET_SHEET_PATH:
            sheet_xml = data.decode("utf-8")
            entries[i] = (
                name,
                _patch_sheet_cells(sheet_xml, cell_values).encode("utf-8"),
            )
            patched = True
            break

    if not patched:
        raise RuntimeError(
            f"{TARGET_SHEET_PATH} not found in template {template_path.name}"
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries:
            zout.writestr(name, data)


# ---------------------------------------------------------------------------
# Value assembly
# ---------------------------------------------------------------------------
def _multiline(values: list[str]) -> str:
    return "\n".join(v for v in values if v and v.strip())


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def generate(
    df,
    *,
    product: str = "iBatchInsight",
    version: str = "A1",
    initiator: str = "",
    host: str = "",
    reviewer: str = "",
    review_date: str = "",
    subject: str = "",
    output_name: str | None = None,
    output_dir: Path | str | None = None,
) -> Path:
    template_path = INPUTS_DIR / TEMPLATE_NAME
    if not template_path.exists():
        raise FileNotFoundError(f"template missing: {template_path}")

    # Per-requirement generation: df is always a single row.
    row = df.iloc[0]

    # Build the multiline cell contents from the dataframe.
    ones_lines: list[str] = []
    link_lines: list[str] = []
    for _, r in df.iterrows():
        rid = str(r.get("ID", "")).strip()
        title = str(r.get("标题", "")).strip()
        link = str(r.get("测试用例链接", "")).strip()
        if rid or title:
            ones_lines.append(f"{rid} {title}".strip())
        if link:
            link_lines.append(link)

    # Resolve the four data-driven fields.  CLI args (if non-empty) override
    # the strategies in content_rules.yaml::excel.
    excel_cfg = load_content_rules().get("excel", {})
    initiator_value = initiator or resolve_field_by_strategy(excel_cfg.get("initiator", {}), row)
    host_value = host or resolve_field_by_strategy(excel_cfg.get("host", {}), row)
    reviewer_value = reviewer or resolve_field_by_strategy(excel_cfg.get("reviewer", {}), row)
    review_date_value = review_date or resolve_field_by_strategy(excel_cfg.get("review_date", {}), row)
    case_link_value = resolve_field_by_strategy(excel_cfg.get("case_link", {}), row)
    case_links_text = case_link_value or _multiline(link_lines)

    cell_values = {
        CELL_PRODUCT: product or "iBatchInsight",
        CELL_ONES_REQS: _multiline(ones_lines),
        CELL_CASE_LINKS: case_links_text,
        CELL_INITIATOR: initiator_value,
        CELL_HOST: host_value,
        CELL_REVIEWER: reviewer_value,
        CELL_REVIEW_DATE: review_date_value,
    }

    # Filename: render from YAML template (see config/filename_templates.yaml).
    iterations = list_iterations(df)
    if len(iterations) == 1:
        iteration_value = iterations[0]
    elif iterations:
        iteration_value = "多迭代"
    else:
        iteration_value = ""
    req_id_value = ""
    if len(df) == 1:
        req_id_value = str(df.iloc[0].get("ID", "")).strip().lstrip("#")

    templates = load_filename_templates()
    name = output_name or render_filename(
        templates["svn_excel"],
        product=product,
        title=subject,
        iteration=iteration_value,
        version=version or "",
        req_id=req_id_value,
    )

    subdir = get_output_subdir("svn_excel", templates)
    base_dir = Path(output_dir) if output_dir else ensure_outputs_dir()
    out_dir = base_dir / subdir if subdir else base_dir
    out_path = out_dir / name

    _write_patched_xlsx(template_path, out_path, cell_values)
    return out_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate SVN Excel (LC-SOP-RC-007-R02).")
    parser.add_argument("--iteration", help="Filter on 所属迭代 before generating.")
    parser.add_argument("--product", default="iBatchInsight")
    parser.add_argument("--version", default="A1")
    parser.add_argument("--initiator", default="", help="发起人")
    parser.add_argument("--host", default="", help="主持人")
    parser.add_argument("--reviewer", default="", help="评审人")
    parser.add_argument("--review-date", default="", help="评审时间 (YYYY-MM-DD)")
    parser.add_argument("--data-dir", default=None, help="Override requirement_data directory.")
    parser.add_argument("--output-dir", default=None)
    parser.add_argument("--output-name", default=None)
    args = parser.parse_args()

    from data_loader import load_all_requirements

    df = load_all_requirements(args.data_dir, iteration=args.iteration)
    out = generate(
        df,
        product=args.product,
        version=args.version,
        initiator=args.initiator,
        host=args.host,
        reviewer=args.reviewer,
        review_date=args.review_date,
        output_name=args.output_name,
        output_dir=args.output_dir,
    )
    print(f"OK  rows={len(df)}  -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
