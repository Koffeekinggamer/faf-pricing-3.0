"""Branded retail PDF output for Foothills Amish Furniture."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows
from scripts.pricing_lib import assert_no_approx

REPO_ROOT = Path(__file__).resolve().parent.parent
LOGO_PATH = REPO_ROOT / "assets" / "faf-logo.png"

FAF_MOSS = colors.HexColor("#454C4B")
FAF_BG = colors.HexColor("#F4EFE4")
FAF_TEXT = colors.HexColor("#1d2726")
FAF_MUTED = colors.HexColor("#7d8281")


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


def build_retail_pdf(
    grid: list[list],
    *,
    builder_name: str,
    updated: str,
    output_path: Path,
) -> None:
    """Legacy flat-grid PDF (simple CSV / small sheets)."""
    from scripts.dealer_book import ProcessedRow

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
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=letter,
        leftMargin=0.55 * inch,
        rightMargin=0.55 * inch,
        topMargin=0.85 * inch,
        bottomMargin=0.65 * inch,
        title=f"{builder_name} — Foothills Amish Furniture",
    )

    styles = getSampleStyleSheet()
    body = ParagraphStyle(
        "fafBody",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        textColor=FAF_TEXT,
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

    story = []
    table_buf: list[list[str]] = []

    def flush_sku_table():
        nonlocal table_buf
        if not table_buf:
            return
        t = Table(table_buf, colWidths=[2.4 * inch, 1.1 * inch, 1.05 * inch, 1.05 * inch])
        t.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), FAF_MOSS),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("GRID", (0, 0), (-1, -1), 0.25, FAF_MUTED),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, FAF_BG]),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(t)
        story.append(Spacer(1, 6))
        table_buf = []

    for pr in processed:
        if pr.kind in (RowKind.SKU, RowKind.TABLE_HEADER):
            if pr.kind == RowKind.TABLE_HEADER and table_buf:
                flush_sku_table()
            table_buf.append([str(c) for c in pr.cells])
            continue

        flush_sku_table()

        if pr.kind == RowKind.BLANK:
            story.append(Spacer(1, 4))
        elif pr.kind == RowKind.TITLE:
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
        elif pr.kind == RowKind.WIDE:
            cells = [str(c) for c in pr.cells]
            while cells and not cells[-1].strip():
                cells.pop()
            if not any(c.strip() for c in cells):
                continue
            ncol = len(cells)
            usable = 7.4 * inch
            first_w = min(1.35 * inch, usable * 0.22)
            rest_w = (usable - first_w) / max(ncol - 1, 1)
            widths = [first_w] + [rest_w] * (ncol - 1) if ncol > 1 else [usable]
            t = Table([cells], colWidths=widths)
            t.setStyle(
                TableStyle(
                    [
                        ("FONTSIZE", (0, 0), (-1, -1), 6.5),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("TOPPADDING", (0, 0), (-1, -1), 1),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                    ]
                )
            )
            story.append(t)
        elif pr.kind == RowKind.PROSE:
            left = pr.cells[0] if pr.cells else ""
            right = pr.cells[1] if len(pr.cells) > 1 else ""
            if right:
                t = Table([[left, right]], colWidths=[3.25 * inch, 3.25 * inch])
                t.setStyle(
                    TableStyle(
                        [
                            ("FONTSIZE", (0, 0), (-1, -1), 8),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("TOPPADDING", (0, 0), (-1, -1), 1),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                        ]
                    )
                )
                story.append(t)
            else:
                story.append(_p(left, body))

    flush_sku_table()

    page_count_box = {"n": 0}

    def on_page(canvas, _doc):
        canvas.saveState()
        page_w, page_h = letter
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

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return page_count_box["n"]
