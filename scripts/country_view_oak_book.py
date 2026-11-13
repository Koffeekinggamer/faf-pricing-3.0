"""Country View Oak finished wholesale PDF → retail ProcessedRows."""

from __future__ import annotations

import re
from pathlib import Path

import pdfplumber

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows
from scripts.dealer_terms import is_dealer_terms_text
from scripts.pdf_plain_text import clean_pdf_plain_text
from scripts.pricing_lib import retail_whole_dollar, transform_text_prices
from scripts.retail_text import forbidden_phrases, sanitize_retail_text

BUILDER = "Country View Oak"
DOC_TITLE = "Country View Oak Retail Price List · April 15, 2025"
MULTIPLIER_DEFAULT = 2.7

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

DOLLAR_RE = re.compile(
    r"\$\s*(\d{1,3}(?:,\d{3})*)(?:\.(\d{2}))?"
)

TABLE_HEADER_RE = re.compile(r"item\s+leaves\s+legs", re.I)

SECTION_RE = re.compile(
    r"^(?:\d+\"\s+)?[A-Za-z0-9].*(?:"
    r"Tables|Beds|Chairs|Stools|Benches|Desks|Hutches|Servers|Cabinets|"
    r"Bookcases|Chests|Mirrors|Racks|Gliders|Crates|Stands|Stools)"
    r"(?:\s*\(cont\.\))?\s*$",
    re.I,
)


def is_country_view_oak_pdf(path: Path) -> bool:
    if path.suffix.lower() != ".pdf":
        return False
    try:
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


def _wholesale_amounts(line: str) -> list[float]:
    out: list[float] = []
    for m in DOLLAR_RE.finditer(line):
        whole = m.group(1).replace(",", "")
        cents = m.group(2) or "00"
        out.append(float(f"{whole}.{cents}"))
    return out


def _transform_price_line(line: str, multiplier: float) -> list[str]:
    line = clean_pdf_plain_text(line)
    if not line or "$" not in line:
        return []
    idx = line.index("$")
    label = sanitize_retail_text(line[:idx].strip(), multiplier)
    if not label and not _wholesale_amounts(line):
        return []
    cells: list[str] = []
    if label:
        cells.append(label)
    for wholesale in _wholesale_amounts(line):
        cells.append(f"${retail_whole_dollar(wholesale, multiplier):,}")
    return cells if len(cells) > 1 or (cells and "$" in cells[-1]) else []


def _transform_prose_line(line: str, multiplier: float) -> str:
    line = clean_pdf_plain_text(line)
    if DROP_LINE_RE.search(line) or is_dealer_terms_text(line):
        return ""
    line = re.sub(r"\bwholesale\b", "", line, flags=re.I)
    line = re.sub(r"\s{2,}", " ", line).strip()
    if not line:
        return ""
    if "$" in line:
        line = transform_text_prices(line, multiplier)
    return sanitize_retail_text(line, multiplier, allow_percent=True)


def extract_pdf_lines(path: Path) -> list[str]:
    lines: list[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw in text.splitlines():
                s = clean_pdf_plain_text(raw).strip()
                if s:
                    lines.append(s)
    return lines


def count_wholesale_dollars(lines: list[str]) -> int:
    n = 0
    for line in lines:
        if DROP_LINE_RE.search(line):
            continue
        n += len(_wholesale_amounts(line))
    return n


def process_country_view_oak_pdf(
    path: Path,
    *,
    multiplier: float = MULTIPLIER_DEFAULT,
) -> tuple[list[ProcessedRow], dict]:
    source_lines = extract_pdf_lines(path)
    expected_dollars = count_wholesale_dollars(source_lines)
    out: list[ProcessedRow] = [
        ProcessedRow(RowKind.TITLE, [DOC_TITLE]),
    ]
    output_dollars = 0
    header_buf: list[str] = []

    for line in source_lines:
        if DROP_LINE_RE.search(line):
            continue
        if is_dealer_terms_text(line):
            continue

        amounts = _wholesale_amounts(line)
        if amounts:
            cells = _transform_price_line(line, multiplier)
            if cells:
                out.append(ProcessedRow(RowKind.WIDE, cells))
                output_dollars += len(amounts)
                header_buf = []
            continue

        if TABLE_HEADER_RE.search(line):
            header_buf.append(line)
            combined = " — ".join(header_buf) if header_buf else line
            hdr = sanitize_retail_text(combined, multiplier)
            if hdr:
                out.append(ProcessedRow(RowKind.SUBHEAD, [hdr]))
            header_buf = []
            continue

        if SECTION_RE.match(line) and "$" not in line:
            hdr = sanitize_retail_text(line, multiplier)
            if hdr:
                out.append(ProcessedRow(RowKind.SECTION, [hdr]))
            header_buf = []
            continue

        low = line.lower()
        if any(
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
        ) and "$" not in line:
            header_buf.append(line)
            continue

        prose = _transform_prose_line(line, multiplier)
        if prose:
            output_dollars += len(_wholesale_amounts(line))
            if prose.lower().startswith("options"):
                out.append(ProcessedRow(RowKind.SUBHEAD, [prose]))
            elif prose.lower().startswith("standard"):
                out.append(ProcessedRow(RowKind.SUBHEAD, [prose]))
            else:
                out.append(ProcessedRow(RowKind.PROSE, [prose]))

    stats = {
        "source_lines": len(source_lines),
        "expected_wholesale_dollars": expected_dollars,
        "output_wholesale_dollars": output_dollars,
        "dollar_parity": expected_dollars == output_dollars,
    }
    if not stats["dollar_parity"]:
        raise ValueError(
            f"Country View Oak extraction mismatch: "
            f"expected {expected_dollars} wholesale dollar cells, "
            f"transformed {output_dollars}"
        )
    return collapse_blank_rows(out), stats


def country_view_oak_forbidden(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            bad.extend(forbidden_phrases(str(c)))
            if is_dealer_terms_text(str(c)):
                bad.append(str(c))
    return bad
