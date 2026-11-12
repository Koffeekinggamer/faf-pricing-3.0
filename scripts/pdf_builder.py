"""Branded retail PDF output for Foothills Amish Furniture."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows
from scripts.pdf_table_layout import (
    should_flush_wide_block,
    sku_table_flowables,
    wide_table_flowables,
)
from scripts.pricing_lib import assert_no_approx

REPO_ROOT = Path(__file__).resolve().parent.parent
LOGO_PATH = REPO_ROOT / "assets" / "faf-logo.png"

FAF_MOSS = colors.HexColor("#454C4B")
FAF_BG = colors.HexColor("#F4EFE4")
FAF_TEXT = colors.HexColor("#1d2726")
FAF_MUTED = colors.HexColor("#7d8281")

MARGIN_L = 0.55 * inch
MARGIN_R = 0.55 * inch
MARGIN_TOP = 0.85 * inch
MARGIN_BOTTOM = 0.65 * inch


def _format_updated(iso_date: str) -> str:
    try:
        dt = datetime.strptime(iso_date, "%Y-%m-%d")
        return dt.strftime("%b %-d, %Y")
    except ValueError:
        return iso_date


def _p(text: str, style: ParagraphStyle) -> Paragraph:
    safe = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    assert_no_approx(safe)
    return Paragraph(safe, style)


def _prose_paragraph(cells: list, body: ParagraphStyle) -> Paragraph | None:
    parts: list[str] = []
    for c in cells:
        t = str(c).strip()
        if not t:
            continue
        if t.startswith(("-", "•", "*")):
            parts.append(t)
        else:
            parts.append(t)
    if not parts:
        return None
    if len(parts) == 1:
        joined = parts[0]
    elif all(p.startswith(("-", "•", "*")) for p in parts):
        joined = "<br/>".join(_esc_fragment(p) for p in parts)
    else:
        joined = "<br/><br/>".join(_esc_fragment(p) for p in parts)
    return _p(joined, body)


def _esc_fragment(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def build_retail_pdf(
    grid: list[list],
    *,
    builder_name: str,
    updated: str,
    output_path: Path,
) -> None:
    rows = [
        ProcessedRow(RowKind.PROSE, [str(c) for c in row if c not in (None, "")])
        for row in grid
        if any(c not in (None, "") for c in row)
    ]
    build_retail_pdf_processed(
        rows, builder_name=builder_name, updated=updated, output_path=output_path
    )


def build_retail_pdf_processed(
    processed: list[ProcessedRow],
    *,
    builder_name: str,
    updated: str,
    output_path: Path,
) -> int:
    assert_no_approx(builder_name, updated)
    processed = collapse_blank_rows(processed)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    styles = getSampleStyleSheet()
    body = ParagraphStyle(
        "fafBody",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        textColor=FAF_TEXT,
        alignment=TA_LEFT,
    )
    title_style = ParagraphStyle(
        "fafTitle",
        parent=body,
        fontSize=16,
        leading=20,
        alignment=TA_CENTER,
        spaceAfter=12,
        textColor=FAF_MOSS,
        fontName="Helvetica-Bold",
    )
    section_style = ParagraphStyle(
        "fafSection",
        parent=body,
        fontSize=12,
        leading=14,
        spaceBefore=8,
        spaceAfter=6,
        textColor=FAF_MOSS,
        fontName="Helvetica-Bold",
        alignment=TA_CENTER,
    )
    subhead_style = ParagraphStyle(
        "fafSub",
        parent=body,
        fontSize=10,
        leading=12,
        spaceBefore=4,
        spaceAfter=4,
        fontName="Helvetica-Bold",
    )

    story: list = []
    table_buf: list[list[str]] = []
    wide_buf: list[list[str]] = []

    def flush_sku_table():
        nonlocal table_buf
        if not table_buf:
            return
        story.extend(sku_table_flowables(table_buf, body_style=body))
        table_buf = []

    def flush_wide_block():
        nonlocal wide_buf
        if not wide_buf:
            return
        story.extend(wide_table_flowables(wide_buf, body_style=body))
        wide_buf = []

    page_count_box = {"n": 0}

    def draw_header_footer(canvas, _doc):
        canvas.saveState()
        page_w, page_h = canvas._pagesize
        header_h = 0.55 * inch
        canvas.setFillColor(FAF_MOSS)
        canvas.rect(0, page_h - header_h, page_w, header_h, fill=1, stroke=0)

        if LOGO_PATH.exists():
            try:
                logo_w = 1.65 * inch
                logo_h = logo_w * (170 / 500)
                canvas.drawImage(
                    str(LOGO_PATH),
                    0.45 * inch,
                    page_h - header_h + (header_h - logo_h) / 2,
                    width=logo_w,
                    height=logo_h,
                    mask="auto",
                    preserveAspectRatio=True,
                )
            except Exception:
                pass

        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 10)
        canvas.drawRightString(
            page_w - 0.55 * inch,
            page_h - header_h + 0.22 * inch,
            builder_name[:60],
        )
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(
            page_w - 0.55 * inch,
            page_h - header_h + 0.08 * inch,
            f"Updated {_format_updated(updated)}",
        )

        canvas.setFillColor(FAF_MUTED)
        canvas.setFont("Helvetica", 8)
        canvas.drawCentredString(
            page_w / 2, 0.35 * inch, f"Page {canvas.getPageNumber()}"
        )
        canvas.drawCentredString(
            page_w / 2, 0.22 * inch, "Foothills Amish Furniture · Landrum, SC"
        )
        page_count_box["n"] = canvas.getPageNumber()
        canvas.restoreState()

    pw, ph = letter
    lw, lh = landscape(letter)
    portrait_frame = Frame(
        MARGIN_L,
        MARGIN_BOTTOM,
        pw - MARGIN_L - MARGIN_R,
        ph - MARGIN_TOP - MARGIN_BOTTOM,
        id="portrait_frame",
    )
    landscape_frame = Frame(
        MARGIN_L,
        MARGIN_BOTTOM,
        lw - MARGIN_L - MARGIN_R,
        lh - MARGIN_TOP - MARGIN_BOTTOM,
        id="landscape_frame",
    )

    doc = BaseDocTemplate(
        str(output_path),
        pagesize=letter,
        pageTemplates=[
            PageTemplate(
                id="portrait",
                frames=[portrait_frame],
                onPage=draw_header_footer,
            ),
            PageTemplate(
                id="landscape",
                frames=[landscape_frame],
                pagesize=landscape(letter),
                onPage=draw_header_footer,
            ),
        ],
        title=f"{builder_name} — Foothills Amish Furniture",
    )

    for pr in processed:
        if pr.kind in (RowKind.SKU, RowKind.TABLE_HEADER):
            flush_wide_block()
            if pr.kind == RowKind.TABLE_HEADER and table_buf:
                flush_sku_table()
            table_buf.append([str(c) for c in pr.cells])
            continue

        if pr.kind == RowKind.BLANK:
            story.append(Spacer(1, 4))
            continue

        if pr.kind == RowKind.WIDE:
            flush_sku_table()
            cells = [str(c) for c in pr.cells]
            if should_flush_wide_block(wide_buf, cells):
                flush_wide_block()
            wide_buf.append(cells)
            continue

        flush_sku_table()
        flush_wide_block()

        if pr.kind == RowKind.TITLE:
            story.append(_p(pr.cells[0], title_style))
        elif pr.kind == RowKind.SECTION:
            story.append(_p(pr.cells[0], section_style))
        elif pr.kind == RowKind.SUBHEAD:
            story.append(_p(pr.cells[0], subhead_style))
        elif pr.kind == RowKind.INDEX:
            if len(pr.cells) >= 2:
                data = [pr.cells[i : i + 2] for i in range(0, len(pr.cells), 2)]
                for pair in data:
                    while len(pair) < 2:
                        pair.append("")
                t = Table(data, colWidths=[3.25 * inch, 3.25 * inch])
                t.setStyle(
                    TableStyle(
                        [
                            ("FONTSIZE", (0, 0), (-1, -1), 9),
                            ("TEXTCOLOR", (0, 0), (-1, -1), FAF_TEXT),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                        ]
                    )
                )
                story.append(t)
            elif pr.cells:
                story.append(_p(" · ".join(pr.cells), body))
        elif pr.kind == RowKind.PROSE:
            para = _prose_paragraph(pr.cells, body)
            if para is not None:
                story.append(para)
                story.append(Spacer(1, 4))

    flush_sku_table()
    flush_wide_block()

    doc.build(story)
    return page_count_box["n"]
