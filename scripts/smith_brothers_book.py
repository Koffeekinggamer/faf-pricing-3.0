"""Smith Brothers FAF retail book PDF → customer catalog ProcessedRows (1.0× pass-through)."""

from __future__ import annotations

import re
from pathlib import Path

import pdfplumber

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows
from scripts.dealer_terms import is_dealer_terms_text
from scripts.pdf_plain_text import clean_pdf_plain_text
from scripts.pricing_lib import transform_text_prices
from scripts.retail_text import forbidden_phrases, sanitize_retail_text

BUILDER = "Smith Brothers"
DOC_TITLE = "Smith Brothers Retail Price List · Fall 2025 (Jan 12, 2026)"
MULTIPLIER = 1.0

FOOTPRINT_RE = re.compile(
    r"^\s*\d+\s+Foothills Amish Furniture Retail Prices-FALL 2025---01-12-2026\s+\d+\s*$",
    re.I,
)
PAGE_NUM_ONLY_RE = re.compile(r"^\s*\d{1,3}\s+\d{1,3}\s*$")
SKU_ROW_RE = re.compile(r"^\d{3}-\d{2}(?:\s+\d{3}-\d{2})+\s*$")
SKU_GROUP_RE = re.compile(r"^\d{3}\s+Group\s*$", re.I)
DROP_LINE_RE = re.compile(
    r"(?:\bwholesale\b|preferred\s+dealer|\bapprox\b|\bdealer\s+price)",
    re.I,
)


def _looks_reversed(line: str) -> bool:
    if len(line) < 20:
        return False
    letters = re.sub(r"[^a-zA-Z]", "", line)
    if len(letters) < 12:
        return False
    return letters == letters[::-1]


def _clean_line(raw: str, multiplier: float) -> str:
    line = clean_pdf_plain_text(raw).strip()
    if not line:
        return ""
    if FOOTPRINT_RE.match(line) or PAGE_NUM_ONLY_RE.match(line):
        return ""
    if _looks_reversed(line):
        return ""
    if DROP_LINE_RE.search(line) or is_dealer_terms_text(line):
        return ""
    line = re.sub(r"\bapproximately\b", "about", line, flags=re.I)
    line = re.sub(r"\bapproximate\b", "estimated", line, flags=re.I)
    line = re.sub(r"\bapprox\.?\b", "", line, flags=re.I)
    line = re.sub(r"\s{2,}", " ", line).strip()
    line = transform_text_prices(line, multiplier)
    return sanitize_retail_text(line, multiplier, allow_percent=True)


def process_smith_brothers_pdfs(
    paths: list[Path],
    *,
    multiplier: float = MULTIPLIER,
) -> tuple[list[ProcessedRow], dict]:
    ordered = sorted(paths, key=lambda p: p.name.casefold())
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]
    lines_kept = 0
    pages = 0

    for path in ordered:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                pages += 1
                text = page.extract_text() or ""
                for raw in text.splitlines():
                    line = _clean_line(raw, multiplier)
                    if not line:
                        continue
                    lines_kept += 1
                    if SKU_ROW_RE.match(line):
                        out.append(ProcessedRow(RowKind.SECTION, [line]))
                    elif SKU_GROUP_RE.match(line):
                        out.append(ProcessedRow(RowKind.SUBHEAD, [line]))
                    elif (
                        re.match(r"^[A-Z][A-Z\s/\-]{8,}$", line)
                        and "$" not in line
                        and len(line.split()) <= 8
                    ):
                        out.append(ProcessedRow(RowKind.SUBHEAD, [line]))
                    else:
                        out.append(ProcessedRow(RowKind.PROSE, [line]))

    stats = {"source_pages": pages, "content_lines": lines_kept, "parts": len(ordered)}
    return collapse_blank_rows(out), stats


def smith_brothers_forbidden(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            bad.extend(forbidden_phrases(str(c)))
            if is_dealer_terms_text(str(c)):
                bad.append(str(c))
    return bad


def pdf_contains_prices(pdf_path: Path, needles: list[str]) -> dict[str, bool]:
    with pdfplumber.open(pdf_path) as pdf:
        full = "\n".join((p.extract_text() or "") for p in pdf.pages)
    return {n: n in full for n in needles}