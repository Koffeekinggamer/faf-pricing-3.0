"""Meadow Wood multi-sheet customer workbook → retail PDF."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.retail_text import forbidden_phrases, retail_price_cell, sanitize_retail_text

BUILDER_NAME = "Meadow Wood"
DOC_TITLE = "Meadow Wood Retail Price List · Effective July 1, 2026"
SKIP_SHEETS = frozenset({"Instructions"})
DEALER_RE = re.compile(
    r"pleasant view finishing|apple creek|viztech|client portal|"
    r"excel pricelist password|register\.|existing dealer|square one marketing|"
    r"picked up at:",
    re.I,
)
DIMENSION_RE = re.compile(r'\d+\s*[wWdD]\s*x\s*\d+', re.I)


def is_meadow_wood_workbook(path: Path) -> bool:
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return False
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return "Instructions" in names and "Inside Cover" in names


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    return str(c).strip()


def _joined(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def _drop_row(row: list) -> bool:
    return bool(DEALER_RE.search(_joined(row)))


def _is_add_label_cell(row: list, j: int) -> bool:
    if j > 0:
        prev = _cell_str(row[j - 1])
        if prev.endswith("Add:") or prev.endswith("Add"):
            return True
    return False


def _should_retail_number(row: list, j: int, val: float | int) -> bool:
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return False
    if val <= 0:
        return False
    if _is_add_label_cell(row, j):
        return True
    c0 = _cell_str(row[0])
    if c0 in ("Item #", "Description"):
        return False
    if j in (2, 3) and (DIMENSION_RE.search(_cell_str(row[j])) or (val < 120 and j in (2, 3))):
        if isinstance(row[j], str):
            return False
        if isinstance(val, (int, float)) and val < 120 and "Add" not in _joined(row):
            return False
    if j >= 4 and val >= 10:
        return True
    if j == 3 and _cell_str(row[2]).endswith("Add:"):
        return True
    return False


def _transform_row(row: list, multiplier: float) -> list[str]:
    width = max(
        (j for j, c in enumerate(row) if _cell_str(c) or isinstance(c, (int, float))),
        default=0,
    )
    out: list[str] = []
    for j in range(width + 1):
        c = row[j] if j < len(row) else None
        if isinstance(c, (int, float)) and not isinstance(c, bool):
            if _should_retail_number(row, j, c):
                out.append(retail_price_cell(c, multiplier))
            elif j in (2, 3) and c < 120:
                out.append(_cell_str(c))
            else:
                out.append(_cell_str(c) if c else "")
        else:
            t = sanitize_retail_text(_cell_str(c), multiplier) if c not in (None, "") else ""
            if "wholesale" in t.lower():
                t = ""
            out.append(t)
    while out and not out[-1]:
        out.pop()
    return out


def process_meadow_wood_workbook(path: Path, *, multiplier: float) -> list[ProcessedRow]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]

    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        if ws.title in SKIP_SHEETS:
            continue
        out.append(
            ProcessedRow(
                RowKind.SECTION,
                [sanitize_retail_text(ws.title, multiplier)],
            )
        )
        for row in ws.iter_rows(values_only=True):
            row = list(row)
            if row_is_blank(row):
                out.append(ProcessedRow(RowKind.BLANK, []))
                continue
            if _drop_row(row):
                continue
            joined = _joined(row)
            if "wholesale" in joined.lower() and "markup" in joined.lower():
                continue
            cells = _transform_row(row, multiplier)
            if not cells:
                continue
            if len(cells) == 1 and len(cells[0]) > 100:
                out.append(ProcessedRow(RowKind.PROSE, cells))
            elif _cell_str(row[0]) == "Item #" or _looks_like_sku(_cell_str(row[0])):
                out.append(ProcessedRow(RowKind.WIDE, cells))
            elif any(isinstance(row[j], (int, float)) for j in range(len(row)) if j >= 4):
                out.append(ProcessedRow(RowKind.WIDE, cells))
            else:
                out.append(ProcessedRow(RowKind.PROSE, cells))

    wb.close()
    return collapse_blank_rows(out)


def _looks_like_sku(s: str) -> bool:
    return bool(s and re.match(r"^[A-Z]{2,}[-\w]", s, re.I))


def meadow_wood_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            s = str(c)
            bad.extend(forbidden_phrases(s))
            if re.search(r"\bwholesale\b", s, re.I):
                bad.append(s)
    return bad


def meadow_builder_display_name() -> str:
    return BUILDER_NAME
