"""Lamb Woodworking Wholesale sheet → retail PDF (ignore retail tabs)."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.retail_text import forbidden_phrases, retail_price_cell, sanitize_retail_text

DOC_TITLE = "Lamb Woodworking Retail Price List · Effective May 1, 2026"
SHEET_NAME = "Wholesale"
LA_ITEM_RE = re.compile(r"^LA-", re.I)
DEALER_RE = re.compile(
    r"shipshewana|5510\s*w\s*200|260-768|260-264|send orders to:|"
    r"info@lambww\.com|companyartwork\.com",
    re.I,
)
WHOLESALE_TITLE_RE = re.compile(r"^\s*2026\s+wholesale", re.I)


def is_lamb_workbook(path: Path) -> bool:
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return False
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return (
        "Wholesale" in names
        and "Retail Finished" in names
        and "Retail Unfinished" in names
    )


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    return str(c).strip()


def _joined(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def _drop_dealer_row(row: list) -> bool:
    return bool(DEALER_RE.search(_joined(row)))


def _max_used_col(row: list) -> int:
    m = -1
    for j, c in enumerate(row):
        if _cell_str(c) or (isinstance(c, (int, float)) and not isinstance(c, bool)):
            m = max(m, j)
    return max(m, 0)


def _is_section_head(row: list) -> bool:
    c0 = _cell_str(row[0] if row else "")
    if not c0 or len(c0) > 80:
        return False
    if LA_ITEM_RE.match(c0):
        return False
    if c0 in ("Item No.", "LIGHTING", "HARDWARE", "SEATING"):
        return True
    if c0.isupper() and not c0.startswith("LA-") and len(c0) < 50:
        if any(ch.isalpha() for ch in c0):
            return True
    return False


def _is_price_cell(row: list, j: int, val: float | int) -> bool:
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return False
    if val == 0:
        return False
    joined = _joined(row).lower()
    if j == 1 and val == int(val) and 1 <= int(val) <= 20 and _cell_str(row[2] if len(row) > 2 else ""):
        return False
    if val in (0.1, 0.5) and ("knot" in joined or "walnut" in joined):
        return False
    if val < 1 and "%" in joined:
        return False
    return val >= 1 or (0 < val < 1 and "per" in joined)


def _format_non_price_number(val: float | int) -> str:
    if val == 0.5:
        return "50%"
    if val == 0.1 and val < 1:
        return "10%"
    return _cell_str(val)


def _transform_row(row: list, multiplier: float) -> list[str]:
    width = _max_used_col(row)
    out: list[str] = []
    for j in range(width + 1):
        c = row[j] if j < len(row) else None
        if isinstance(c, (int, float)) and not isinstance(c, bool):
            if _is_price_cell(row, j, c):
                out.append(retail_price_cell(c, multiplier))
            elif c in (0.1, 0.5):
                out.append(_format_non_price_number(c))
            else:
                out.append(_cell_str(c) if c != 0 else "")
        else:
            t = sanitize_retail_text(_cell_str(c), multiplier) if c not in (None, "") else ""
            if WHOLESALE_TITLE_RE.search(t):
                t = ""
            out.append(t)
    while out and not out[-1]:
        out.pop()
    return out


def _load_wholesale_rows(path: Path) -> list[list]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if SHEET_NAME not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Missing {SHEET_NAME!r} sheet")
    ws = wb[SHEET_NAME]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def process_lamb_workbook(path: Path, *, multiplier: float) -> list[ProcessedRow]:
    rows = _load_wholesale_rows(path)
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]

    for row in rows:
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        if _drop_dealer_row(row):
            continue
        joined = _joined(row)
        if WHOLESALE_TITLE_RE.search(joined):
            continue
        if re.fullmatch(r"effective may 1, 2026", joined, re.I):
            continue
        if _is_section_head(row):
            out.append(
                ProcessedRow(
                    RowKind.SUBHEAD,
                    [sanitize_retail_text(_cell_str(row[0]), multiplier)],
                )
            )
            continue
        cells = _transform_row(row, multiplier)
        if not cells:
            continue
        if len(cells) == 1 and len(cells[0]) > 120:
            out.append(ProcessedRow(RowKind.PROSE, cells))
        else:
            out.append(ProcessedRow(RowKind.WIDE, cells))

    return collapse_blank_rows(out)


def lamb_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            s = str(c)
            bad.extend(forbidden_phrases(s))
            if re.search(r"\bwholesale\b", s, re.I):
                bad.append(s)
            if "companyartwork" in s.lower():
                bad.append(s)
    return bad
