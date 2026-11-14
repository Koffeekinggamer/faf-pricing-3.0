"""Viztech workbooks: import explicit finished wholesale columns only (avoid retail sheets)."""

from __future__ import annotations

import re
from typing import Iterable

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.dealer_terms import is_builder_contact_text, is_dealer_terms_text
from scripts.pdf_plain_text import clean_pdf_plain_text
from scripts.retail_text import sanitize_retail_text
from scripts.retail_text import retail_price_cell
from scripts.viztech_book import (
    SpotSample,
    ViztechBuildResult,
    _cell_str,
    _is_wholesale_price,
    viztech_forbidden,
)


def _drop_line(joined: str, patterns: Iterable[re.Pattern]) -> bool:
    if is_dealer_terms_text(joined) or is_builder_contact_text(joined):
        return True
    for pat in patterns:
        if pat.search(joined):
            return True
    return False


def _transform_finished_row(
    row: list,
    *,
    finished_cols: tuple[int, ...],
    multiplier: float,
    spots: list[SpotSample],
) -> list[str] | None:
    label_parts: list[str] = []
    for j in range(min(4, len(row))):
        s = _cell_str(row[j])
        if s and not _is_wholesale_price(row[j]):
            label_parts.append(s)
    label = sanitize_retail_text(" — ".join(label_parts), multiplier)
    if not label:
        return None
    cells = [label]
    got = False
    for j in finished_cols:
        if j >= len(row):
            continue
        v = row[j]
        if _is_wholesale_price(v):
            retail = retail_price_cell(v, multiplier)
            cells.append(retail)
            got = True
            if len(spots) < 6:
                spots.append(
                    SpotSample(
                        item=label[:60],
                        wholesale=float(v),
                        retail=retail,
                    )
                )
        elif isinstance(row[j], str) and row[j].strip() in ("√", "Yes", "Standard"):
            cells.append(row[j].strip())
    if not got:
        return None
    return cells


def process_finished_wholesale_sheet(
    rows: list[list],
    *,
    finished_cols: tuple[int, ...],
    multiplier: float,
    drop_patterns: Iterable[re.Pattern],
    spots: list[SpotSample],
    include_prose: bool = True,
) -> list[ProcessedRow]:
    out: list[ProcessedRow] = []
    sku_re = re.compile(r"^(?:IT|JR)[A-Z]{0,2}-", re.I)
    for row in rows:
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        joined = clean_pdf_plain_text(" ".join(_cell_str(c) for c in row if _cell_str(c)))
        if not joined:
            continue
        first = _cell_str(row[0]) if row else ""
        is_price_row = bool(sku_re.match(first)) and any(
            _is_wholesale_price(c) for c in row
        )
        if is_price_row:
            if any(pat.search(joined) for pat in drop_patterns):
                continue
        elif _drop_line(joined, drop_patterns):
            continue
        wide = _transform_finished_row(
            row, finished_cols=finished_cols, multiplier=multiplier, spots=spots
        )
        if wide:
            out.append(ProcessedRow(RowKind.WIDE, wide))
            continue
        first = _cell_str(row[0]) if row else ""
        if first and len(first) <= 40 and first.isupper() and not sku_re.match(first):
            if first in {"ID. NO.", "SKU"}:
                continue
            hdr = sanitize_retail_text(first, multiplier)
            if hdr:
                out.append(ProcessedRow(RowKind.SECTION, [hdr]))
            continue
        if include_prose and joined and "$" not in joined and "%" in joined:
            prose = sanitize_retail_text(joined, multiplier, allow_percent=True)
            if prose:
                out.append(ProcessedRow(RowKind.PROSE, [prose]))
        elif include_prose and joined and not _is_wholesale_price(row[0] if row else None):
            if len(joined) > 12 and not re.search(r"\bapprox", joined, re.I):
                prose = sanitize_retail_text(joined, multiplier, allow_percent=True)
                if prose and not is_builder_contact_text(prose):
                    out.append(ProcessedRow(RowKind.PROSE, [prose]))
    return out


def build_viztech_result(
    rows: list[ProcessedRow],
    spots: list[SpotSample],
    *,
    skipped: list[str] | None = None,
) -> ViztechBuildResult:
    data = [r for r in rows if r.kind not in (RowKind.BLANK,)]
    if not data:
        return ViztechBuildResult([], spots, skipped or [], ["No rows produced"])
    return ViztechBuildResult(collapse_blank_rows(rows), spots, skipped or [], [])


__all__ = [
    "process_finished_wholesale_sheet",
    "build_viztech_result",
    "viztech_forbidden",
]