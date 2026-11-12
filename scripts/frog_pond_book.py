"""Frog Pond Furniture workbook → retail PDF."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.retail_text import forbidden_phrases, retail_price_cell, sanitize_retail_text

BUILDER = "Frog Pond Furniture"
DOC_TITLE = "Frog Pond Furniture Retail Price List"
SKIP_SHEETS = frozenset({"Markup"})
COVER_SHEETS = frozenset({"Index", "Drawer Cover", "TV Stands Cover"})
DIMENSION_RE = re.compile(r'^\d{1,3}(\.\d+)?$')


def is_frog_pond_workbook(path: Path) -> bool:
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return False
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return "Markup" in names and any(n.strip().startswith("1 Weston") for n in names)


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    return str(c).strip()


def _joined(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def _is_dimension(row: list, j: int, val: Any) -> bool:
    if j in (2, 3, 4) and isinstance(val, (int, float)) and not isinstance(val, bool):
        if 0 < float(val) <= 120:
            return True
    if isinstance(val, str) and re.search(r"\d+\"", val):
        return True
    return False


def _is_finished_price(row: list, j: int, val: Any) -> bool:
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return False
    if float(val) == 1.0:
        return False
    if float(val) <= 0:
        return False
    if _is_dimension(row, j, val):
        return False
    c0 = _cell_str(row[0])
    if c0 in ("Item #", "H\"", "W\"", "D\""):
        return False
    if j >= 5 and float(val) >= 8:
        return True
    if j == 5 and _cell_str(row[0]) and float(val) >= 8:
        return True
    # Options numeric column 3 (index) or 5
    if _joined(row).lower().startswith("options") or "add" in _joined(row).lower():
        if j in (3, 5) and float(val) >= 1 and float(val) != 1.0:
            return True
    return False


def _transform_row(row: list, multiplier: float) -> list[str]:
    price_cols = {j for j, c in enumerate(row) if _is_finished_price(row, j, c)}
    width = max(
        (j for j in range(len(row)) if _cell_str(row[j]) or j in price_cols),
        default=0,
    )
    out: list[str] = []
    for j in range(width + 1):
        c = row[j] if j < len(row) else None
        if j in price_cols:
            out.append(retail_price_cell(c, multiplier))
        elif isinstance(c, (int, float)) and not isinstance(c, bool):
            if c == 1 or _is_dimension(row, j, c):
                out.append(_cell_str(c) if c != 1 else "")
            else:
                out.append(_cell_str(c))
        else:
            t = sanitize_retail_text(_cell_str(c), multiplier) if c not in (None, "") else ""
            out.append(t)
    while out and not out[-1]:
        out.pop()
    return out


def process_frog_pond_workbook(path: Path, *, multiplier: float) -> list[ProcessedRow]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    title = DOC_TITLE
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        if ws.title == "Index":
            for row in ws.iter_rows(values_only=True):
                for c in row:
                    if c and re.search(r"effective", str(c), re.I):
                        title = f"{DOC_TITLE} · {sanitize_retail_text(str(c), multiplier)}"
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [title])]

    for ws in wb.worksheets:
        if ws.sheet_state != "visible" or ws.title in SKIP_SHEETS:
            continue
        if ws.title in COVER_SHEETS:
            out.append(ProcessedRow(RowKind.SECTION, [ws.title]))
            for row in ws.iter_rows(values_only=True):
                row = list(row)
                if row_is_blank(row):
                    continue
                cells = [
                    sanitize_retail_text(_cell_str(c), multiplier)
                    for c in row
                    if _cell_str(c) and not re.match(r"^\d+$", _cell_str(c))
                ]
                cells = [c for c in cells if c and "wholesale" not in c.lower()]
                if cells:
                    out.append(ProcessedRow(RowKind.PROSE, cells[:6]))
            continue

        out.append(
            ProcessedRow(RowKind.SECTION, [sanitize_retail_text(ws.title.strip(), multiplier)])
        )
        for row in ws.iter_rows(values_only=True):
            row = list(row)
            if row_is_blank(row):
                out.append(ProcessedRow(RowKind.BLANK, []))
                continue
            joined = _joined(row)
            if "wholesale" in joined.lower():
                continue
            cells = _transform_row(row, multiplier)
            if cells:
                out.append(ProcessedRow(RowKind.WIDE, cells))

    wb.close()
    return collapse_blank_rows(out)


def frog_pond_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            s = str(c)
            bad.extend(forbidden_phrases(s))
            if re.search(r"\bwholesale\b", s, re.I):
                bad.append(s)
    return bad
