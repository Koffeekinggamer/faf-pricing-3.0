"""Country View Oak dealer workbook → retail ProcessedRows."""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows
from scripts.dealer_terms import is_dealer_terms_text
from scripts.pdf_plain_text import clean_pdf_plain_text
from scripts.pricing_lib import retail_whole_dollar, transform_text_prices
from scripts.retail_text import forbidden_phrases, sanitize_retail_text

BUILDER = "Country View Oak"
DOC_TITLE = "Country View Oak Retail Price List · April 15, 2025"
MULTIPLIER_DEFAULT = 2.7
BUILDER_MARKUP = 2.82  # wholesale × (1 + MARKUP!E28) on retail columns; never read those cols

WHOLESALE_COLS = (4, 6, 8, 10, 12)  # D, F, H, J, L
HEADER_WHOLESALE_COLS = WHOLESALE_COLS

DROP_LINE_RE = re.compile(
    r"(?:"
    r"owen\s+mast|"
    r"8368\s+state\s+route|"
    r"fredericksburg\s*,?\s*oh|"
    r"for\s+unfinished\s+tables\s+deduct|"
    r"unfinished\s+tables\s+deduct|"
    r"^\s*finished\s+price\s+list\s*$|"
    r"^\s*april\s+15\s*,\s*2025\s*$"
    r")",
    re.I,
)

TABLE_HEADER_RE = re.compile(r"^item\s*$", re.I)

SECTION_RE = re.compile(
    r"^(?:\d+\"\s+)?[A-Za-z0-9].*(?:"
    r"Tables|Beds|Chairs|Stools|Benches|Desks|Hutches|Servers|Cabinets|"
    r"Bookcases|Chests|Mirrors|Racks|Gliders|Crates|Stands)"
    r"(?:\s*\(cont\.\))?\s*$",
    re.I,
)


def is_country_view_oak_workbook(path: Path) -> bool:
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return False
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            names = set(wb.sheetnames)
            if "RETAIL" not in names or "MARKUP" not in names:
                return False
            sh = wb["RETAIL"]
            a35 = sh.cell(35, 1).value
            return isinstance(a35, str) and "finished price list" in a35.lower()
        finally:
            wb.close()
    except Exception:
        return False


def is_country_view_oak_pdf(path: Path) -> bool:
    """Legacy: builder PDF upload with finished list header (prefer workbook)."""
    if path.suffix.lower() != ".pdf":
        return False
    try:
        import pdfplumber

        with pdfplumber.open(path) as pdf:
            if len(pdf.pages) < 80:
                return False
            head = (pdf.pages[0].extract_text() or "").lower()
            return (
                "finished price list" in head
                and "april 15, 2025" in head
                and ("owen mast" in head or "8368" in head)
            )
    except Exception:
        return False


def _cell_str(v: object) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _wholesale_number(v: object) -> float | None:
    if not isinstance(v, (int, float)) or isinstance(v, bool):
        return None
    n = float(v)
    if n <= 0:
        return None
    if n < 1:
        return None
    return n


def sparse_wholesale_values(row_values: tuple[object, ...]) -> list[float]:
    """Wholesale cells from D/F/H/J/L that are not a full five-group table row."""
    idx = (3, 5, 7, 9, 11)
    return [
        n
        for i in idx
        if i < len(row_values) and (n := _wholesale_number(row_values[i])) is not None
    ]


def wholesale_values_five_group(row_values: tuple[object, ...]) -> list[float] | None:
    """All five wood-group wholesales when this is a full price row."""
    idx = [3, 5, 7, 9, 11]
    nums: list[float] = []
    for i in idx:
        if i >= len(row_values):
            return None
        n = _wholesale_number(row_values[i])
        if n is None:
            return None
        nums.append(n)
    return nums


def retail_from_wholesale(wholesale: float, multiplier: float = MULTIPLIER_DEFAULT) -> str:
    return f"${retail_whole_dollar(wholesale, multiplier):,}"


def row_label(row_values: tuple[object, ...]) -> str:
    parts: list[str] = []
    for i in range(3):
        s = _cell_str(row_values[i] if i < len(row_values) else None)
        if s:
            parts.append(s)
    return " ".join(parts)


def _transform_prose(text: str, multiplier: float) -> str:
    text = clean_pdf_plain_text(text)
    if not text or DROP_LINE_RE.search(text) or is_dealer_terms_text(text):
        return ""
    text = re.sub(r"\bwholesale\b", "", text, flags=re.I)
    text = re.sub(r"\s{2,}", " ", text).strip()
    if not text:
        return ""
    if "$" in text:
        text = transform_text_prices(text, multiplier)
    return sanitize_retail_text(text, multiplier, allow_percent=True)


def _rows_from_retail_sheet(sh, multiplier: float) -> list[ProcessedRow]:
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]
    header_buf: list[str] = []

    for r in range(1, sh.max_row + 1):
        row = tuple(sh.cell(r, c).value for c in range(1, 14))
        if not any(v is not None and str(v).strip() for v in row):
            continue

        a = _cell_str(row[0])
        b = _cell_str(row[1]) if len(row) > 1 else ""
        joined = clean_pdf_plain_text(" ".join(x for x in (a, b) if x))

        if DROP_LINE_RE.search(joined):
            continue

        five = wholesale_values_five_group(row)
        sparse = sparse_wholesale_values(row) if five is None else None

        if TABLE_HEADER_RE.match(a):
            sub_parts = [a, _cell_str(row[1]), _cell_str(row[2])]
            for ci in HEADER_WHOLESALE_COLS:
                h = sh.cell(r, ci).value
                if h:
                    sub_parts.append(clean_pdf_plain_text(str(h).replace("\n", " ")))
            hdr = sanitize_retail_text(" — ".join(p for p in sub_parts if p), multiplier)
            if hdr:
                out.append(ProcessedRow(RowKind.SUBHEAD, [hdr]))
            header_buf = []
            continue

        if five:
            label = row_label(row)
            cells = [label] if label else []
            cells.extend(retail_from_wholesale(w, multiplier) for w in five)
            if cells:
                out.append(ProcessedRow(RowKind.WIDE, cells))
            header_buf = []
            continue

        if sparse:
            label = row_label(row) or joined
            cells = [label]
            cells.extend(retail_from_wholesale(w, multiplier) for w in sparse)
            out.append(ProcessedRow(RowKind.WIDE, cells))
            header_buf = []
            continue

        if SECTION_RE.match(a) and "$" not in joined:
            hdr = sanitize_retail_text(a, multiplier)
            if hdr:
                out.append(ProcessedRow(RowKind.SECTION, [hdr]))
            header_buf = []
            continue

        low = joined.lower()
        if (
            any(
                x in low
                for x in (
                    "brown maple",
                    "rustic cherry",
                    "cherry sap",
                    "wormy maple",
                    "white oak",
                    "walnut",
                    "number of",
                )
            )
            and "$" not in joined
        ):
            header_buf.append(joined)
            continue

        text = a or b
        if not text and joined:
            text = joined
        prose = _transform_prose(text, multiplier)
        if prose:
            if prose.lower().startswith("options"):
                out.append(ProcessedRow(RowKind.SUBHEAD, [prose]))
            elif prose.lower().startswith("standard"):
                out.append(ProcessedRow(RowKind.SUBHEAD, [prose]))
            else:
                out.append(ProcessedRow(RowKind.PROSE, [prose]))

    return out


def process_country_view_oak_workbook(
    path: Path,
    *,
    multiplier: float = MULTIPLIER_DEFAULT,
) -> tuple[list[ProcessedRow], dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        sh = wb["RETAIL"]
        rows = _rows_from_retail_sheet(sh, multiplier)
    finally:
        wb.close()

    price_rows = sum(1 for pr in rows if pr.kind == RowKind.WIDE)
    dollar_cells = sum(
        len([c for c in pr.cells if str(c).startswith("$")]) for pr in rows if pr.kind == RowKind.WIDE
    )
    stats = {
        "processed_rows": len(rows),
        "wide_rows": price_rows,
        "retail_dollar_cells": dollar_cells,
    }
    return collapse_blank_rows(rows), stats


def process_country_view_oak_pdf(
    path: Path,
    *,
    multiplier: float = MULTIPLIER_DEFAULT,
) -> tuple[list[ProcessedRow], dict]:
    raise ValueError(
        "Country View Oak must be imported from the .xlsx workbook "
        "(RETAIL wholesale columns D/F/H/J/L), not the builder PDF."
    )


def country_view_oak_forbidden(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            bad.extend(forbidden_phrases(str(c)))
            if is_dealer_terms_text(str(c)):
                bad.append(str(c))
    return bad


def spot_check_36x48_solid(rows: list[ProcessedRow]) -> dict[str, str]:
    """Map wood group column index → site retail for 36x48 Solid row."""
    for pr in rows:
        if pr.kind != RowKind.WIDE or not pr.cells:
            continue
        label = str(pr.cells[0])
        if "36" in label and "48" in label and "solid" in label.lower():
            prices = [c for c in pr.cells[1:] if str(c).startswith("$")]
            if len(prices) >= 5:
                names = ("oak", "cherry", "qshw", "brown_maple_group", "walnut")
                return dict(zip(names, prices, strict=True))
    raise AssertionError('36" X 48" Solid row not found in output')