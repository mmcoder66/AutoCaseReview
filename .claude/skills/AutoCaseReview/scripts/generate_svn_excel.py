"""Generate the SVN Excel deliverable (LC-SOP-RC-007-R02).

Strategy: copy the template xlsx **byte-for-byte**, then patch ONLY the
specific data cells in ``测试用例评审记录`` (sheet3) using direct XML edit.

This bypasses openpyxl entirely, which would otherwise drop:
  * ``xl/embeddings/*.docx`` (signature Word docs on 封皮 / 签名页)
  * ``xl/media/*.png`` / ``*.emf`` (logos & drawings)
  * ``xl/drawings/*.vml`` + ``drawing*.xml`` (VML / drawings)

After patching, the cover & signature pages remain visually identical to the
template — embedded objects, images and formatting all preserved.

A second, opt-in XML patch targets the embedded signature Word doc
(``xl/embeddings/Microsoft_Word___1.docx``): the 起草 row's 签名 cell is
filled with the requirement's 测试 person.  This patch only inserts a single
``<w:r>`` into an existing empty ``<w:p>``, leaving every other style/element
intact.
"""

from __future__ import annotations

import argparse
import io
import re
import zipfile
import xml.sax.saxutils as saxutils
from pathlib import Path

from data_loader import (
    OUTPUTS_DIR,
    TEMPLATES_DIR,
    ensure_outputs_dir,
    get_default_product,
    get_output_subdir,
    list_iterations,
    load_content_rules,
    load_filename_templates,
    load_template_names,
    render_filename,
    resolve_field_by_strategy,
)

TARGET_SHEET_PATH = "xl/worksheets/sheet3.xml"   # 测试用例评审记录

# Cell coordinates we patch in 测试用例评审记录.
CELL_PRODUCT = "C3"          # merged C3:E3
CELL_ONES_REQS = "C4"        # merged C4:E4
CELL_CASE_LINKS = "C5"       # merged C5:E5
CELL_INITIATOR = "C6"
CELL_HOST = "E6"
CELL_REVIEWER = "C7"
CELL_REVIEW_DATE = "E7"

# Embedded signature Word doc inside the xlsx (signature page on sheet2).
SIGNATURE_DOC_PATH = "xl/embeddings/Microsoft_Word___1.docx"
SIGNATURE_DOC_XML = "word/document.xml"
SIGNATURE_ROW_INDEX = 1      # 起草行（0-based: 0=表头, 1=起草, 2=审核, 3=批准）
SIGNATURE_CELL_INDEX = 3     # 签名 cell（合并后 5 cell: 标签/部门/印刷体姓名/签名/日期）


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


# ---------------------------------------------------------------------------
# Embedded signature Word doc patching
# ---------------------------------------------------------------------------
def _patch_signature_doc_xml(doc_xml: str, signer_name: str) -> str:
    """Insert *signer_name* into the 起草 row's 签名 cell.

    The signature page's first table is laid out as:
        row 0: header (部门 / 印刷体姓名 / 签名 / 日期)
        row 1: 起草 (label col 0, 签名 cell index 3)
        row 2: 审核 ; row 3: 批准 ; ...

    The 签名 cell already contains an empty ``<w:p>`` whose ``<w:pPr>`` defines
    a paragraph-level ``<w:rPr>`` default.  We splice one ``<w:r>`` (carrying
    the same font but non-bold, per SOP convention for signature content) just
    before the closing ``</w:p>``.  Every other style, merge, or element is
    left untouched.
    """
    if not signer_name:
        return doc_xml

    # Locate the 起草 <w:tr> block.
    tr_pattern = re.compile(r"<w:tr[ >].*?</w:tr>", re.DOTALL)
    trs = list(tr_pattern.finditer(doc_xml))
    if len(trs) <= SIGNATURE_ROW_INDEX:
        return doc_xml  # template shape unexpected; skip silently

    row1_match = trs[SIGNATURE_ROW_INDEX]
    row1_xml = row1_match.group(0)

    # Locate the 签名 <w:tc> inside the 起草 row.
    tc_pattern = re.compile(r"<w:tc>.*?</w:tc>", re.DOTALL)
    tcs = list(tc_pattern.finditer(row1_xml))
    if len(tcs) <= SIGNATURE_CELL_INDEX:
        return doc_xml

    target_tc_match = tcs[SIGNATURE_CELL_INDEX]
    target_tc_xml = target_tc_match.group(0)

    # Build the run XML.  Font matches the paragraph's rPr default (宋体 21半点),
    # but without <w:b/> so the actual signature content stays non-bold.
    escaped = saxutils.escape(signer_name)
    run_xml = (
        '<w:r w:rsidRPr="00770068">'
        "<w:rPr>"
        '<w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:hint="eastAsia"/>'
        '<w:color w:val="000000" w:themeColor="text1"/>'
        '<w:kern w:val="0"/>'
        '<w:sz w:val="21"/><w:szCs w:val="21"/>'
        "</w:rPr>"
        f'<w:t xml:space="preserve">{escaped}</w:t>'
        "</w:r>"
    )

    # Splice the run before the closing </w:p> of the cell's first paragraph.
    p_match = re.search(r"<w:p[ >].*?</w:p>", target_tc_xml, re.DOTALL)
    if not p_match:
        return doc_xml

    p_xml = p_match.group(0)
    close_tag = "</w:p>"
    close_idx = p_xml.rfind(close_tag)
    new_p_xml = p_xml[:close_idx] + run_xml + p_xml[close_idx:]
    new_tc_xml = (
        target_tc_xml[:p_match.start()]
        + new_p_xml
        + target_tc_xml[p_match.end():]
    )
    new_row1_xml = (
        row1_xml[:target_tc_match.start()]
        + new_tc_xml
        + row1_xml[target_tc_match.end():]
    )
    return doc_xml[:row1_match.start()] + new_row1_xml + doc_xml[row1_match.end():]


def _patch_signature_docx(docx_bytes: bytes, signer_name: str) -> bytes:
    """Patch the embedded signature Word doc (zip-in-zip).

    Opens the docx bytes as a zip, patches ``word/document.xml`` to fill the
    起草 签名 cell, then re-zips.  All other docx members are copied verbatim
    so styles / numbering / settings remain untouched.
    """
    if not signer_name:
        return docx_bytes

    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
        entries = [(info.filename, zin.read(info.filename)) for info in zin.infolist()]

    patched = False
    for i, (name, data) in enumerate(entries):
        if name == SIGNATURE_DOC_XML:
            new_xml = _patch_signature_doc_xml(data.decode("utf-8"), signer_name)
            entries[i] = (name, new_xml.encode("utf-8"))
            patched = True
            break

    if not patched:
        return docx_bytes  # template without the expected document.xml; skip

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries:
            zout.writestr(name, data)
    return out.getvalue()


def _write_patched_xlsx(template_path: Path, output_path: Path,
                        cell_values: dict[str, str],
                        signature_name: str = "") -> None:
    """Copy *template_path* to *output_path*, patching only target cells.

    Every zip member other than ``TARGET_SHEET_PATH`` is copied verbatim,
    preserving order and binary content (images / embeddings / drawings).

    When *signature_name* is non-empty, the embedded signature Word doc
    (``SIGNATURE_DOC_PATH``) is also patched to fill the 起草 row's 签名 cell.
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
        elif signature_name and name == SIGNATURE_DOC_PATH:
            entries[i] = (name, _patch_signature_docx(data, signature_name))

    if not patched:
        raise RuntimeError(
            f"{TARGET_SHEET_PATH} not found in template {template_path.name}"
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        output_path.unlink()
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
    product: str | None = None,
    version: str = "A1",
    initiator: str = "",
    host: str = "",
    reviewer: str = "",
    review_date: str = "",
    signature_name: str = "",
    subject: str = "",
    output_name: str | None = None,
    output_dir: Path | str | None = None,
) -> Path:
    product_value = product or get_default_product()
    template_path = TEMPLATES_DIR / load_template_names()["svn_excel"]
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

    # 签名页 起草 签名 cell (default strategy: 测试 column).
    signature_value = signature_name or resolve_field_by_strategy(
        excel_cfg.get("signature", {}), row
    )

    cell_values = {
        CELL_PRODUCT: product_value,
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
        product=product_value,
        title=subject,
        iteration=iteration_value,
        version=version or "",
        req_id=req_id_value,
    )

    subdir = get_output_subdir("svn_excel", templates)
    base_dir = Path(output_dir) if output_dir else ensure_outputs_dir()
    out_dir = base_dir / subdir if subdir else base_dir
    out_path = out_dir / name

    _write_patched_xlsx(template_path, out_path, cell_values,
                        signature_name=signature_value)
    return out_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate SVN Excel (LC-SOP-RC-007-R02).")
    parser.add_argument("--iteration", help="Filter on 所属迭代 before generating.")
    parser.add_argument("--product", default=get_default_product())
    parser.add_argument("--version", default="A1")
    parser.add_argument("--initiator", default="", help="发起人")
    parser.add_argument("--host", default="", help="主持人")
    parser.add_argument("--reviewer", default="", help="评审人")
    parser.add_argument("--review-date", default="", help="评审时间 (YYYY-MM-DD)")
    parser.add_argument("--signature", default="",
                        help="签名页起草行签名列填充值；默认按 content_rules.yaml::excel.signature 策略派生")
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
        signature_name=args.signature,
        output_name=args.output_name,
        output_dir=args.output_dir,
    )
    print(f"OK  rows={len(df)}  -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
