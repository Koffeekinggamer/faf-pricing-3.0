"""Measure and lay out wide retail price tables for PDF output."""

from __future__ import annotations

import re

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle

from scripts.pdf_plain_text import escape_for_paragraph_markup
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import (
    NextPageTemplate,
    PageBreak,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

MIN_FONT_PT = 7.0
DEFAULT_FONT_PT = 8.0
PORTRAIT_USABLE = 7.4 * inch
LANDSCAPE_USABLE = 10.0 * inch
PRICE_RE = re.compile(r"^\$[\d,]+$")
CAT_ROW_RE = re.compile(r"^Cat\.\s*\d", re.I)

FAF_MOSS = colors.HexColor("#454C4B")
FAF_BG = colors.HexColor("#F4EFE4")
FAF_MUTED = colors.HexColor("#7d8281")

_DASH_ONLY = frozenset({"", "—", "-", "–", "‑"})


def dash_or_empty(text: str) -> bool:
    t = (text or "").strip()
    return t in _DASH_ONLY


def prune_empty_columns(rows: list[list[str]]) -> list[list[str]]:
    if not rows:
        return []
    ncol = max(len(r) for r in rows)
    padded = [list(r) + [""] * (ncol - len(r)) for r in rows]
    keep: list[int] = []
    for j in range(ncol):
        vals = [padded[i][j].strip() for i in range(len(padded))]
        if all(dash_or_empty(v) for v in vals):
            continue
        keep.append(j)
    if not keep:
        return [[(r[0] if r else "").strip()] for r in padded]
    return [[padded[i][j] for j in keep] for i in range(len(padded))]


def _prefix_column_count(rows: list[list[str]]) -> int:
    if not rows:
        return 1
    ncol = max(len(r) for r in rows)
    if ncol <= 3:
        return 1
    # Prefer header labels when present (Item # / Description, Size / Leaf)
    h0 = (rows[0][0] or "").strip().lower() if rows[0] else ""
    h1 = (rows[0][1] or "").strip().lower() if len(rows[0]) > 1 else ""
    if h0 in {"item #", "item", "size", "order #", "sku"} and h1 in {
        "description",
        "leaf / top",
        "desc",
        "product",
        "size",
    }:
        return 2
    if h0 == "size" and "leaf" in h1:
        return 2
    second_labels = {r[1].strip() for r in rows if len(r) > 1 and r[1].strip()}
    if second_labels and all(
        len(s) <= 12 and not PRICE_RE.match(s) for s in second_labels
    ):
        return 2
    return 1


def _natural_col_widths(
    rows: list[list[str]], font_name: str, font_size: float
) -> list[float]:
    ncol = max(len(r) for r in rows)
    widths = [0.42 * inch] * ncol
    for row in rows:
        for j, cell in enumerate(row):
            text = (cell or "").strip() or " "
            # Width from first line only so multi-line wood headers don't inflate columns
            first_line = text.split("\n", 1)[0]
            w = stringWidth(first_line, font_name, font_size) + 8
            if PRICE_RE.match(text):
                w = max(w, 0.55 * inch)
            elif j == 0:
                w = max(w, 0.9 * inch)
            elif j < 2:
                # Keep "Description" readable (never clip to Descriptio/n)
                w = max(w, stringWidth("Description", font_name, font_size) + 10)
                w = max(w, 0.95 * inch)
            else:
                w = max(w, 0.48 * inch)
            widths[j] = max(widths[j], min(w, 2.0 * inch))
    return widths


def _scale_widths(
    widths: list[float],
    target: float,
    *,
    floor: float,
    protect: int = 0,
) -> list[float]:
    """Scale columns to fit target; optionally protect the first `protect` columns."""
    total = sum(widths)
    if total <= target or total <= 0:
        return widths
    protected = list(widths[:protect])
    rest = list(widths[protect:])
    rest_target = max(target - sum(protected), floor * max(len(rest), 1))
    if not rest:
        return widths
    rest_total = sum(rest)
    if rest_total <= rest_target:
        return widths
    factor = rest_target / rest_total
    scaled_rest = [max(w * factor, floor) for w in rest]
    total2 = sum(protected) + sum(scaled_rest)
    if total2 > target and rest:
        overflow = total2 - target
        # Shrink rest a bit more
        shrink = (sum(scaled_rest) - overflow) / sum(scaled_rest)
        scaled_rest = [max(w * shrink, floor * 0.9) for w in scaled_rest]
    return protected + scaled_rest


def _cell_para(text: str, style: ParagraphStyle) -> Paragraph:
    t = (text or "").strip()
    if not t or dash_or_empty(t):
        t = "—"
    return Paragraph(escape_for_paragraph_markup(t), style)


_PREFIX_HEADER_LABELS = frozenset(
    {
        "item #",
        "item",
        "description",
        "size",
        "leaf / top",
        "price",
        "order #",
        "sku",
        "desc",
        "product",
    }
)


def _stack_multiword_wood_header(cell: str) -> str:
    """Prefer wrapping wood names at spaces (RUSTIC\\nWALNUT), not mid-word."""
    t = (cell or "").strip()
    if not t or "\n" in t or " " not in t:
        return t
    if t.lower() in _PREFIX_HEADER_LABELS or "/" in t:
        return t
    return "\n".join(t.split())


def _table_style(header_rows: int, font_size: float) -> TableStyle:
    rules: list[tuple] = [
        ("GRID", (0, 0), (-1, -1), 0.25, FAF_MUTED),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
    ]
    if header_rows > 0:
        rules.extend(
            [
                ("BACKGROUND", (0, 0), (-1, header_rows - 1), FAF_MOSS),
                ("TEXTCOLOR", (0, 0), (-1, header_rows - 1), colors.white),
                ("FONTNAME", (0, 0), (-1, header_rows - 1), "Helvetica-Bold"),
            ]
        )
        rules.append(
            ("ROWBACKGROUNDS", (0, header_rows), (-1, -1), [colors.white, FAF_BG])
        )
    else:
        rules.append(("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, FAF_BG]))
    return TableStyle(rules)


def _row_looks_like_header(row: list[str]) -> bool:
    if not row:
        return False
    prices = sum(1 for c in row if PRICE_RE.match((c or "").strip()))
    nonempty = sum(1 for c in row if (c or "").strip() and not dash_or_empty(c))
    return nonempty >= 3 and prices == 0


def chunks_for_split(widths: list[float], prefix: int, usable: float) -> list[tuple[int, int]]:
    """Disjoint column ranges [start, end) after prefix columns."""
    if prefix >= len(widths):
        return [(0, len(widths))]
    chunks: list[tuple[int, int]] = []
    i = prefix
    while i < len(widths):
        acc = sum(widths[:prefix])
        start = i
        while i < len(widths) and acc + widths[i] <= usable + 0.01:
            acc += widths[i]
            i += 1
        if i == start:
            acc += widths[i]
            i += 1
        chunks.append((start, i))
    return chunks


def non_key_headers_in_chunks(
    rows: list[list[str]],
    *,
    prefix: int,
    chunks: list[tuple[int, int]],
    header_rows: int,
) -> list[list[str]]:
    """Header labels shown in each chunk (columns after prefix)."""
    if header_rows <= 0 or not rows:
        return [[] for _ in chunks]
    header = rows[0]
    out: list[list[str]] = []
    for c0, c1 in chunks:
        labels = []
        for j in range(c0, c1):
            if j < prefix:
                continue
            t = (header[j] if j < len(header) else "").strip()
            if t and not dash_or_empty(t):
                labels.append(t)
        out.append(labels)
    return out


def assert_split_headers_unique(
    rows: list[list[str]],
    *,
    prefix: int,
    chunks: list[tuple[int, int]],
    header_rows: int,
) -> None:
    """Fail if a later sub-table repeats a non-key column header from an earlier chunk."""
    per_chunk = non_key_headers_in_chunks(
        rows, prefix=prefix, chunks=chunks, header_rows=header_rows
    )
    seen: set[str] = set()
    for ci, labels in enumerate(per_chunk):
        for lab in labels:
            key = lab.lower()
            if key in seen:
                raise ValueError(
                    f"Split chunk {ci + 1} repeats column header {lab!r} from an earlier chunk"
                )
            seen.add(key)


def should_flush_wide_block(buf: list[list[str]], new_row: list[str]) -> bool:
    if not buf:
        return False
    if _row_looks_like_header(new_row):
        return True
    if len(buf[-1]) > 1 and CAT_ROW_RE.match((buf[-1][1] or "").strip()):
        if len(new_row) > 1 and CAT_ROW_RE.match((new_row[1] or "").strip()):
            prev = (buf[-1][1] or "").strip()
            nxt = (new_row[1] or "").strip()
            if prev.startswith("Cat. 3") and nxt.startswith("Cat. 1"):
                return True
    return False


def wide_table_flowables(
    raw_rows: list[list[str]],
    *,
    body_style: ParagraphStyle,
    section_title: str | None = None,
) -> list:
    rows = prune_empty_columns([[str(c) for c in r] for r in raw_rows])
    rows = [
        r
        for r in rows
        if any((c or "").strip() and not dash_or_empty(c) for c in r)
    ]
    if not rows:
        return []

    col_header_rows = 1 if _row_looks_like_header(rows[0]) else 0
    # Optional collection title row that repeats with column headers on every page
    title_rows = 0
    if section_title and section_title.strip():
        ncol0 = max(len(r) for r in rows)
        title_row = [section_title.strip()] + [""] * (ncol0 - 1)
        rows = [title_row] + rows
        title_rows = 1
    header_rows = title_rows + col_header_rows

    prefix = _prefix_column_count(rows[title_rows:] if title_rows else rows)
    font_size = DEFAULT_FONT_PT
    font_name = body_style.fontName

    cell_style = ParagraphStyle(
        "wideCell",
        parent=body_style,
        fontSize=font_size,
        leading=font_size + 1.5,
        alignment=TA_LEFT,
    )
    header_cell_style = ParagraphStyle(
        "wideHeaderCell",
        parent=cell_style,
        fontName="Helvetica-Bold",
        textColor=colors.white,
        fontSize=font_size,
        leading=font_size + 1.5,
    )
    title_cell_style = ParagraphStyle(
        "wideTitleCell",
        parent=header_cell_style,
        fontSize=font_size + 1,
        leading=font_size + 2.5,
        alignment=TA_CENTER,
    )

    # Prefer space-wrap on multi-word wood headers (avoid RUSTIC WALNU/T)
    if col_header_rows:
        hdr_i = title_rows
        hdr = list(rows[hdr_i])
        for j in range(prefix, len(hdr)):
            hdr[j] = _stack_multiword_wood_header(hdr[j])
        rows[hdr_i] = hdr

    # Never break header words mid-glyph when a column is tight
    header_cell_style.splitLongWords = 0

    # Measure widths from column-header + body (ignore title row)
    measure_rows = rows[title_rows:] if title_rows else rows
    widths = _natural_col_widths(measure_rows, font_name, font_size)
    total = sum(widths)
    ncol = max(len(r) for r in rows)
    usable = PORTRAIT_USABLE
    need_landscape = total > usable or ncol >= 7
    if need_landscape:
        usable = LANDSCAPE_USABLE
        while total > usable and font_size > MIN_FONT_PT:
            font_size -= 0.5
            cell_style.fontSize = font_size
            cell_style.leading = font_size + 1.5
            header_cell_style.fontSize = font_size
            header_cell_style.leading = font_size + 1.5
            title_cell_style.fontSize = font_size + 1
            title_cell_style.leading = font_size + 2.5
            widths = _natural_col_widths(measure_rows, font_name, font_size)
            total = sum(widths)

    flow: list = []

    def append_table(sub_rows: list[list[str]], col_widths: list[float]) -> None:
        data = []
        for ri, row in enumerate(sub_rows):
            if title_rows and ri == 0:
                st = title_cell_style
            elif ri < header_rows:
                st = header_cell_style
            else:
                st = cell_style
            data.append([_cell_para(c, st) for c in row])
        tbl = Table(data, colWidths=col_widths, repeatRows=header_rows)
        style = _table_style(header_rows, font_size)
        if title_rows:
            style.add("SPAN", (0, 0), (-1, 0))
            style.add("ALIGN", (0, 0), (-1, 0), "CENTER")
        tbl.setStyle(style)
        flow.append(tbl)

    scaled = _scale_widths(widths, usable, floor=0.35 * inch, protect=min(prefix, 2))
    # Never clip the Description header (Descriptio/n) — measure bold + cell padding
    if prefix >= 2 and len(scaled) > 1:
        desc_min = (
            stringWidth("Description", "Helvetica-Bold", max(font_size, DEFAULT_FONT_PT))
            + 20
        )
        if scaled[1] < desc_min:
            deficit = desc_min - scaled[1]
            scaled[1] = desc_min
            for j in range(len(scaled) - 1, 1, -1):
                if deficit <= 0:
                    break
                take = min(deficit, max(0, scaled[j] - 0.42 * inch))
                scaled[j] -= take
                deficit -= take
    total_scaled = sum(scaled)

    if total_scaled <= usable + 0.01:
        if need_landscape:
            flow.append(NextPageTemplate("landscape"))
            flow.append(PageBreak())
        append_table(rows, scaled)
        flow.append(Spacer(1, 6))
        if need_landscape:
            # Next page becomes portrait; no PageBreak here (avoids blank interstitial pages)
            flow.append(NextPageTemplate("portrait"))
        return flow

    # Split measurement must use the same header layout
    chunks = chunks_for_split(scaled, prefix, usable)
    assert_split_headers_unique(
        measure_rows, prefix=prefix, chunks=chunks, header_rows=col_header_rows
    )

    flow.append(NextPageTemplate("landscape"))
    flow.append(PageBreak())

    for ci, (c0, c1) in enumerate(chunks):
        if ci > 0:
            flow.append(Spacer(1, 3))
        sub_rows = [row[:prefix] + row[c0:c1] for row in rows]
        sub_widths = _scale_widths(
            scaled[:prefix] + scaled[c0:c1],
            usable,
            floor=0.35 * inch,
            protect=min(prefix, 2),
        )
        append_table(sub_rows, sub_widths)

    flow.append(Spacer(1, 6))
    flow.append(NextPageTemplate("portrait"))
    return flow


def sku_table_flowables(
    table_buf: list[list[str]],
    *,
    body_style: ParagraphStyle,
) -> list:
    if not table_buf:
        return []
    if len(table_buf[0]) == 4:
        widths = [2.35 * inch, 1.15 * inch, 1.0 * inch, 1.0 * inch]
    else:
        widths = _scale_widths(
            _natural_col_widths(table_buf, "Helvetica", DEFAULT_FONT_PT),
            PORTRAIT_USABLE,
            floor=0.35 * inch,
        )
    cell_style = ParagraphStyle(
        "skuCell",
        parent=body_style,
        fontSize=DEFAULT_FONT_PT,
        leading=9,
    )
    header_style = ParagraphStyle(
        "skuHeader",
        parent=cell_style,
        fontName="Helvetica-Bold",
        textColor=colors.white,
    )
    data = []
    for ri, row in enumerate(table_buf):
        st = header_style if ri == 0 else cell_style
        data.append([_cell_para(str(c), st) for c in row])
    tbl = Table(data, colWidths=widths, repeatRows=1)
    tbl.setStyle(_table_style(1, DEFAULT_FONT_PT))
    return [tbl, Spacer(1, 4)]
