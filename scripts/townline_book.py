"""Townline Furniture Finished sheet → retail PDF (text-only, no embedded images)."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.retail_text import forbidden_phrases, retail_price_cell, sanitize_retail_text

DOC_TITLE = "Townline Furniture Retail Price List · Effective March 1, 2026"
SHEET = "Finished"
LEFT_PRICE_COLS = range(2, 8)
OPTION_PRICE_COL = 7
DEALER_RE = re.compile(
    r"lagrange|46761|499-487|sales@townline|terms:\s*2%|net\s*30|wholesale|"
    r"add markup|overdue invoices",
    re.I,
)
SKU_RE = re.compile(r"^TL-\d+", re.I)


def is_townline_workbook(path: Path) -> bool:
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return False
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return "Finished" in names and "Unfinished" in names


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    if hasattr(c, "isoformat"):
        return ""
    s = str(c).strip()
    if s.startswith("#"):
        return ""
    return s


def _joined(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def _safe_text(text: str, multiplier: float) -> str:
    t = re.sub(r"\bapproximately\b", "about", text, flags=re.I)
    t = re.sub(r"\bapprox\.?\b", "about", t, flags=re.I)
    return sanitize_retail_text(t, multiplier)


def _option_price(val: Any, multiplier: float) -> str:
    if isinstance(val, (int, float)) and not isinstance(val, bool) and val > 0:
        return retail_price_cell(val, multiplier)
    s = _cell_str(val).upper()
    if s in ("N/C", "NC"):
        return "N/C"
    return _safe_text(_cell_str(val), multiplier)


def _transform_sku_row(row: list, multiplier: float) -> list[str]:
    out: list[str] = []
    for j in range(8):
        c = row[j] if j < len(row) else None
        if j in LEFT_PRICE_COLS and isinstance(c, (int, float)) and not isinstance(c, bool):
            out.append(retail_price_cell(c, multiplier) if c > 0 else "")
        elif j <= 1:
            out.append(_safe_text(_cell_str(c), multiplier) if c not in (None, "") else "")
        else:
            out.append("")
    return out


def process_townline_workbook(path: Path, *, multiplier: float) -> list[ProcessedRow]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[SHEET]
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]

    for row in ws.iter_rows(values_only=True):
        row = list(row)
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        joined = _joined(row)
        if DEALER_RE.search(joined):
            continue
        if "wholesale price sheet" in joined.lower():
            continue
        if "cd." in joined.lower() and "pricing featured" in joined.lower():
            continue

        c0 = _cell_str(row[0])
        if SKU_RE.match(c0):
            out.append(ProcessedRow(RowKind.WIDE, _transform_sku_row(row, multiplier)))
            continue

        if c0 and not c0.startswith("•") and _cell_str(row[1]) == "" and not any(
            isinstance(row[j], (int, float)) for j in LEFT_PRICE_COLS if j < len(row)
        ):
            out.append(ProcessedRow(RowKind.SUBHEAD, [_safe_text(c0, multiplier)]))
            continue

        if _cell_str(row[1]) == "STANDARD FEATURES:":
            continue

        opt_name = _cell_str(row[3]) if len(row) > 3 else ""
        price_val = row[OPTION_PRICE_COL] if len(row) > OPTION_PRICE_COL else None
        if opt_name.startswith("*") or (
            opt_name and price_val is not None and _cell_str(price_val)
        ):
            feature = _safe_text(_cell_str(row[1]), multiplier)
            option = _safe_text(opt_name, multiplier)
            price = _option_price(price_val, multiplier)
            if feature or option:
                out.append(
                    ProcessedRow(
                        RowKind.WIDE,
                        [feature, option, price],
                    )
                )
            continue

        if "•" in joined or "%" in joined:
            prose = _safe_text(joined, multiplier)
            if prose:
                out.append(ProcessedRow(RowKind.PROSE, [prose]))

    wb.close()
    return collapse_blank_rows(out)


def townline_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            s = str(c)
            bad.extend(forbidden_phrases(s))
            if re.search(r"\bwholesale\b", s, re.I):
                bad.append(s)
    return bad
