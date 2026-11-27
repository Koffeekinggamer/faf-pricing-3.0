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
    KeepTogether,
    NextPageTemplate,
    PageBreak,
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
from scripts.dealer_terms import filter_prose_cells, is_dealer_terms_text
from scripts.pdf_plain_text import clean_pdf_plain_text, escape_for_paragraph_markup
from scripts.pricing_lib import assert_no_approx

REPO_ROOT = Path(__file__).resolve().parent.parent
LOGO_PATH = REPO_ROOT / "assets" / "faf-logo.png"

FAF_MOSS = colors.HexColor("#454C4B")
FAF_BG = colors.HexColor("#F4EFE4")
FAF_TEXT = colors.HexColor("#1d2726")
FAF_MUTED = colors.HexColor("#7d8281")

MARGIN_L = 0.55 * inch
MARGIN_R = 0.55 * inch
# Keep clear of the 0.55" branded header bar so continued rows never clip into it
MARGIN_TOP = 1.05 * inch
MARGIN_BOTTOM = 0.65 * inch


def _format_updated(iso_date: str) -> str:
    try:
        dt = datetime.strptime(iso_date, "%Y-%m-%d")
        return dt.strftime("%b %-d, %Y")
    except ValueError:
        return iso_date


def _p(text: str, style: ParagraphStyle) -> Paragraph:
    plain = clean_pdf_plain_text(text)
    assert_no_approx(plain)
    return Paragraph(escape_for_paragraph_markup(plain), style)


def _prose_flowables(cells: list, body: ParagraphStyle) -> list:
    """One full-width paragraph per source cell — never merge columns."""
    parts = filter_prose_cells(cells)
    if not parts:
        return []
    flow: list = []
    for part in parts:
        if is_dealer_terms_text(part):
            continue
        flow.append(_p(part, body))
        flow.append(Spacer(1, 3))
    return flow


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
    all_landscape: bool = False,
    empty_table_cells_as_dash: bool = True,
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
        keepWithNext=True,
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
        keepWithNext=True,
    )
    subhead_style = ParagraphStyle(
        "fafSub",
        parent=body,
        fontSize=10,
        leading=12,
        spaceBefore=4,
        spaceAfter=4,
        fontName="Helvetica-Bold",
        keepWithNext=True,
    )

    story: list = []
    table_buf: list[list[str]] = []
    wide_buf: list[list[str]] = []
    # Deferred lead-ins: ("title"|"section", text)
    pending_lead: list[tuple[str, str]] = []
    # Survives mid-collection header flushes so continued wide blocks keep the title
    active_section: str | None = None
    # Prose/subheads under a pending section — kept with the next table
    pending_intro: list = []

    def _insert_after_page_switches(flow: list, lead: list) -> list:
        if not lead:
            return flow
        i = 0
        while i < len(flow) and isinstance(flow[i], (NextPageTemplate, PageBreak)):
            i += 1
        # Keep section intro + first real flowable together so titles aren't orphaned
        first = flow[i] if i < len(flow) else None
        rest = flow[i + 1 :] if i < len(flow) else []
        if first is not None and lead:
            glued = KeepTogether(lead + [first])
            return flow[:i] + [glued] + rest
        for j, el in enumerate(lead):
            flow.insert(i + j, el)
        return flow

    def flush_sku_table():
        nonlocal table_buf, pending_intro
        if not table_buf:
            return
        titles = [t for k, t in pending_lead if k == "title"]
        sections = [t for k, t in pending_lead if k == "section"]
        pending_lead.clear()
        intro = list(pending_intro)
        pending_intro = []
        flow = sku_table_flowables(table_buf, body_style=body)
        lead = [_p(t, title_style) for t in titles] + [
            _p(s, section_style) for s in sections
        ] + intro
        story.extend(_insert_after_page_switches(flow, lead))
        table_buf = []

    def flush_wide_block():
        nonlocal wide_buf, active_section, pending_intro
        if not wide_buf:
            return
        titles = [t for k, t in pending_lead if k == "title"]
        sections = [t for k, t in pending_lead if k == "section"]
        pending_lead.clear()
        if sections:
            active_section = sections[-1]
        # Mid-collection wood-group flushes clear pending_lead; reuse active_section
        section_title = sections[-1] if sections else active_section
        ncol = max((len(r) for r in wide_buf), default=0)
        # Narrow option tables start on a fresh portrait page (skipped when all-landscape)
        if section_title and ncol <= 3 and not all_landscape:
            story.append(NextPageTemplate("portrait"))
            story.append(PageBreak())
        intro = list(pending_intro)
        pending_intro = []
        flow = wide_table_flowables(
            wide_buf,
            body_style=body,
            section_title=section_title,
            lock_landscape=all_landscape,
            empty_as_dash=empty_table_cells_as_dash,
        )
        lead = [_p(t, title_style) for t in titles]
        for s in sections[:-1]:
            lead.append(_p(s, section_style))
        lead.extend(intro)
        story.extend(_insert_after_page_switches(flow, lead))
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

    landscape_tmpl = PageTemplate(
        id="landscape",
        frames=[landscape_frame],
        pagesize=landscape(letter),
        onPage=draw_header_footer,
    )
    portrait_tmpl = PageTemplate(
        id="portrait",
        frames=[portrait_frame],
        pagesize=letter,  # must be explicit — doc.pagesize may be swapped to landscape
        onPage=draw_header_footer,
    )
    if all_landscape:
        templates = [landscape_tmpl]
        initial_pagesize = landscape(letter)
    else:
        templates = [portrait_tmpl, landscape_tmpl]
        initial_pagesize = letter

    doc = BaseDocTemplate(
        str(output_path),
        pagesize=initial_pagesize,
        pageTemplates=templates,
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
            pending_lead.append(("title", pr.cells[0]))
            continue
        if pr.kind == RowKind.SECTION:
            pending_lead.append(("section", pr.cells[0]))
            active_section = pr.cells[0]
            continue

        # Hold section + intro prose until the next table so titles aren't orphaned
        if pr.kind in (RowKind.SUBHEAD, RowKind.PROSE) and any(
            k == "section" for k, _ in pending_lead
        ):
            if pr.kind == RowKind.SUBHEAD:
                pending_intro.append(_p(pr.cells[0], subhead_style))
            else:
                pending_intro.extend(_prose_flowables(pr.cells, body))
            continue

        # Non-table content: emit deferred section on a clean portrait page
        if pending_lead:
            titles = [t for k, t in pending_lead if k == "title"]
            sections = [t for k, t in pending_lead if k == "section"]
            pending_lead.clear()
            if sections:
                active_section = sections[-1]
            if not all_landscape:
                story.append(NextPageTemplate("portrait"))
                story.append(PageBreak())
            for t in titles:
                story.append(_p(t, title_style))
            for s in sections:
                story.append(_p(s, section_style))
            if pending_intro:
                story.extend(pending_intro)
                pending_intro = []

        if pr.kind == RowKind.SUBHEAD:
            story.append(_p(pr.cells[0], subhead_style))
        elif pr.kind == RowKind.INDEX:
            if len(pr.cells) >= 2:
                data = []
                for i in range(0, len(pr.cells), 2):
                    left = clean_pdf_plain_text(str(pr.cells[i]))
                    right = (
                        clean_pdf_plain_text(str(pr.cells[i + 1]))
                        if i + 1 < len(pr.cells)
                        else ""
                    )
                    data.append([_p(left, body), _p(right, body)])
                t = Table(data, colWidths=[3.25 * inch, 3.25 * inch])
                t.setStyle(
                    TableStyle(
                        [
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("TEXTCOLOR", (0, 0), (-1, -1), FAF_TEXT),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                        ]
                    )
                )
                story.append(t)
            elif pr.cells:
                story.append(_p(" · ".join(pr.cells), body))
        elif pr.kind == RowKind.PROSE:
            story.extend(_prose_flowables(pr.cells, body))

    flush_sku_table()
    flush_wide_block()
    if pending_lead or pending_intro:
        for k, t in pending_lead:
            story.append(_p(t, title_style if k == "title" else section_style))
        pending_lead.clear()
        story.extend(pending_intro)
        pending_intro = []

    # If the first table forced landscape via NextPageTemplate+PageBreak, start on
    # landscape instead of leaving a blank portrait cover page.
    while story and isinstance(story[0], NextPageTemplate):
        action = getattr(story[0], "action", None) or ()
        if len(action) > 1 and action[1] == "landscape":
            templates = list(doc.pageTemplates)
            land = [t for t in templates if t.id == "landscape"]
            port = [t for t in templates if t.id != "landscape"]
            if land:
                doc.pageTemplates = land + port
                doc.pagesize = landscape(letter)
        story.pop(0)
    while story and isinstance(story[0], PageBreak):
        story.pop(0)

    doc.build(story)
    return page_count_box["n"]
