"""Farmside Wood multi-sheet wholesale workbook → retail PDF rows."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.retail_text import (
    WHOLESALE_PRICELIST_RE,
    forbidden_phrases,
    retail_price_cell,
    sanitize_retail_text,
)

ITEM_RE = re.compile(r"^#\w+", re.I)
SKIP_VISIBLE_SHEETS = frozenset({"MarkUp"})
DEALER_HEADER_RE = re.compile(
    r"farmside wood|harrison|apple creek|viztech|support@farmsidewood|330-695",
    re.I,
)
TERMS_RE = re.compile(r"terms\s*:|net\s*30|finance charge", re.I)
DOC_TITLE = "Farmside Woodworking Retail Price List"


def is_farmside_workbook(path: Path) -> bool:
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return False
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return "Title Page" in names and "MarkUp" in names and any("Collection" in n for n in names)


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    return str(c).strip()


def _joined(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def _max_used_col(row: list) -> int:
    m = -1
    for j, c in enumerate(row):
        if _cell_str(c) or (isinstance(c, (int, float)) and not isinstance(c, bool)):
            m = max(m, j)
    return max(m, 0)


def _drop_dealer_row(row: list) -> bool:
    j = _joined(row)
    if DEALER_HEADER_RE.search(j):
        return True
    if TERMS_RE.search(j):
        return True
    if re.search(r"pricelist effective", j, re.I) and re.search(r"phone|harrison", j, re.I):
        return True
    return False


def _is_item_price_row(row: list) -> bool:
    c0 = _cell_str(row[0] if row else "")
    return bool(ITEM_RE.match(c0))


def _transform_row(row: list, multiplier: float, width: int) -> list[str]:
    price_row = _is_item_price_row(row)
    out: list[str] = []
    for j in range(width + 1):
        c = row[j] if j < len(row) else None
        if price_row and j >= 2 and isinstance(c, (int, float)) and not isinstance(c, bool):
            out.append(retail_price_cell(c, multiplier))
        elif isinstance(c, (int, float)) and not isinstance(c, bool):
            out.append("")
        elif isinstance(c, str) or c is not None:
            t = sanitize_retail_text(_cell_str(c), multiplier) if c is not None else ""
            if j == 1 and re.fullmatch(r"[\d\-–—]+", t):
                t = ""
            out.append(t)
        else:
            out.append("")
    while out and not out[-1]:
        out.pop()
    return out


def _process_title_page(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out: list[ProcessedRow] = []
    title_done = False
    in_toc = False
    for row in rows:
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        joined = _joined(row)
        if _drop_dealer_row(row):
            continue
        if "table of contents" in joined.lower():
            in_toc = True
            out.append(ProcessedRow(RowKind.SECTION, ["Table of Contents"]))
            continue
        if WHOLESALE_PRICELIST_RE.search(joined):
            if not title_done:
                out.append(ProcessedRow(RowKind.TITLE, [DOC_TITLE]))
                title_done = True
            continue
        if in_toc and _cell_str(row[0]):
            name = sanitize_retail_text(_cell_str(row[0]), multiplier)
            if name:
                out.append(ProcessedRow(RowKind.PROSE, [name]))
            continue
        if "all prices subject" in joined.lower():
            continue
        cells = []
        for c in row:
            if c is None:
                continue
            if isinstance(c, (int, float)):
                continue
            t = sanitize_retail_text(_cell_str(c), multiplier)
            if t:
                cells.append(t)
        if cells:
            out.append(ProcessedRow(RowKind.PROSE, cells))
    if not title_done:
        out.insert(0, ProcessedRow(RowKind.TITLE, [DOC_TITLE]))
    return out


def process_farmside_workbook(
    path: Path,
    *,
    builder_name: str,
    multiplier: float,
) -> list[ProcessedRow]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    processed: list[ProcessedRow] = []
    first_sheet = True

    for name in wb.sheetnames:
        ws = wb[name]
        if ws.sheet_state != "visible":
            continue
        if name in SKIP_VISIBLE_SHEETS:
            continue

        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if name == "Title Page":
            processed.extend(_process_title_page(rows, multiplier))
            first_sheet = False
            continue

        if first_sheet:
            processed.append(ProcessedRow(RowKind.TITLE, [DOC_TITLE]))
            first_sheet = False

        processed.append(ProcessedRow(RowKind.SECTION, [name.replace(" Collection", "")]))

        width = max((_max_used_col(r) for r in rows), default=11)
        skip_header = True
        for row in rows:
            if row_is_blank(row):
                processed.append(ProcessedRow(RowKind.BLANK, []))
                continue
            if _drop_dealer_row(row):
                skip_header = False
                continue
            joined = _joined(row)
            if skip_header and DEALER_HEADER_RE.search(joined):
                continue
            skip_header = False

            cells = _transform_row(row, multiplier, width)
            if not any(cells):
                continue
            if len([c for c in cells if c]) == 1 and cells[0].endswith("COLLECTION"):
                processed.append(ProcessedRow(RowKind.SUBHEAD, [cells[0]]))
            else:
                processed.append(ProcessedRow(RowKind.WIDE, cells))

    wb.close()
    return collapse_blank_rows(processed)


def farmside_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            bad.extend(forbidden_phrases(str(c)))
    return bad
