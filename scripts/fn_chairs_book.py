"""FN Chairs multi-sheet workbook → retail PDF rows."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.retail_text import forbidden_phrases, retail_price_cell, sanitize_retail_text

DOC_TITLE = "FN Chairs Retail Price List · 2027 · Effective September 1, 2026"
SKIP_SHEETS = frozenset({"PL With Markup", "PCL Color List", "PL To Export"})
SKIP_SHEET_PREFIX = "Extra Blank Page"

ADDRESS_RE = re.compile(
    r"lagrange|463\.8938|499\.4946|sales@fnchairs\.net|2175 west",
    re.I,
)
TERMS_RE = re.compile(r"terms\s*:|net\s*30|finance charge", re.I)
INTERNET_POLICY_RE = re.compile(r"internet policy|mininum pricing|minimum pricing", re.I)
FINISHED_CAT_LABELS = frozenset({"Cat. 1", "Cat. 2", "Cat. 3"})


def _resolve_workbook(path: Path) -> Path | None:
    if path.suffix.lower() in (".xlsx", ".xlsm") and path.is_file():
        return path
    if path.suffix.lower() == ".zip" and path.is_file():
        base = path.parent / f"{path.stem}_fn_extract"
        base.mkdir(exist_ok=True)
        with zipfile.ZipFile(path, "r") as zf:
            zf.extractall(base)
        for p in base.rglob("*.xlsm"):
            return p
        for p in base.rglob("*.xlsx"):
            return p
    return None


def is_fn_chairs_workbook(path: Path) -> bool:
    xlsx = _resolve_workbook(path)
    if not xlsx:
        return False
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return "PL Print" in names and "Cover Page" in names


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    if c == "":
        return ""
    return str(c).strip()


def _joined(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def _drop_dealer_row(row: list) -> bool:
    j = _joined(row)
    if ADDRESS_RE.search(j):
        return True
    if TERMS_RE.search(j):
        return True
    return False


def _is_wood_header_row(row: list) -> bool:
    return (
        bool(_cell_str(row[0]))
        and row[1] is None
        and isinstance(row[2], str)
        and "/" in row[2]
    )


def _is_pl_data_row(row: list) -> bool:
    return _cell_str(row[1]) in FINISHED_CAT_LABELS


def _transform_pl_row(row: list, multiplier: float, width: int) -> list[str]:
    is_data = _is_pl_data_row(row)
    out: list[str] = []
    for j in range(width + 1):
        c = row[j] if j < len(row) else None
        if is_data and j >= 2:
            if c == "" or c is None:
                out.append("—")
            elif isinstance(c, (int, float)) and not isinstance(c, bool):
                out.append(retail_price_cell(c, multiplier))
            else:
                out.append(_cell_str(c))
        elif isinstance(c, (int, float)) and not isinstance(c, bool):
            out.append("")
        else:
            t = sanitize_retail_text(_cell_str(c), multiplier) if c not in (None, "") else ""
            out.append(t)
    while out and not out[-1]:
        out.pop()
    return out


def _process_cover(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    return [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]


def _process_notes_sheet(
    rows: list[list], multiplier: float, *, drop_head: int = 0
) -> list[ProcessedRow]:
    out: list[ProcessedRow] = []
    for i, row in enumerate(rows):
        if i < drop_head and _drop_dealer_row(row):
            continue
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        if _drop_dealer_row(row):
            continue
        cells = []
        for c in row:
            if c is None or isinstance(c, (int, float)):
                if isinstance(c, (int, float)) and c:
                    cells.append(str(c))
                continue
            t = sanitize_retail_text(_cell_str(c), multiplier)
            if t:
                cells.append(t)
        if cells:
            out.append(ProcessedRow(RowKind.PROSE, cells))
    return out


def _process_internet_warranty_only(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out: list[ProcessedRow] = []
    keep = False
    for row in rows:
        joined = _joined(row)
        if re.search(r"warranty\s*&\s*wood specs", joined, re.I):
            keep = True
            out.append(ProcessedRow(RowKind.SECTION, ["Warranty & Wood Specs"]))
            continue
        if not keep:
            if INTERNET_POLICY_RE.search(joined):
                continue
            continue
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        cells = []
        for c in row:
            if c is None:
                continue
            t = sanitize_retail_text(_cell_str(c), multiplier)
            if t and not INTERNET_POLICY_RE.search(t):
                cells.append(t)
        if cells:
            out.append(ProcessedRow(RowKind.PROSE, cells))
    return out


def _process_pl_print(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out: list[ProcessedRow] = []
    width = 23
    current_product = ""
    for row in rows:
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        if _is_wood_header_row(row):
            current_product = ""
            style = sanitize_retail_text(_cell_str(row[0]), multiplier)
            out.append(ProcessedRow(RowKind.SUBHEAD, [style]))
            cells = _transform_pl_row(row, multiplier, width)
            if any(cells[2:]):
                out.append(ProcessedRow(RowKind.WIDE, cells))
            continue
        if _cell_str(row[0]) and not _is_pl_data_row(row):
            current_product = _cell_str(row[0])
        if _cell_str(row[1]) == "Unf":
            continue
        if _is_pl_data_row(row) or _cell_str(row[0]):
            cells = _transform_pl_row(row, multiplier, width)
            if _is_pl_data_row(row) and current_product and not cells[0].strip():
                cells[0] = current_product
            if any(cells):
                out.append(ProcessedRow(RowKind.WIDE, cells))
    return out


def process_fn_chairs_workbook(
    path: Path,
    *,
    builder_name: str,
    multiplier: float,
) -> list[ProcessedRow]:
    xlsx = _resolve_workbook(path)
    if not xlsx:
        raise FileNotFoundError(f"FN Chairs workbook not found in {path}")
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)
    processed: list[ProcessedRow] = []

    for name in wb.sheetnames:
        ws = wb[name]
        if ws.sheet_state != "visible":
            continue
        if name in SKIP_SHEETS or name.startswith(SKIP_SHEET_PREFIX):
            continue

        rows = [list(r) for r in ws.iter_rows(values_only=True)]

        if name == "Cover Page":
            processed.extend(_process_cover(rows, multiplier))
            continue
        if name == "Retired Designs":
            processed.append(ProcessedRow(RowKind.SECTION, ["Retired & New Designs"]))
            processed.extend(_process_notes_sheet(rows, multiplier))
            continue
        if name == "Price List Notes":
            processed.append(ProcessedRow(RowKind.SECTION, ["Price List Notes"]))
            processed.extend(_process_notes_sheet(rows, multiplier, drop_head=3))
            continue
        if name == "Internet Policy - Warranty":
            processed.extend(_process_internet_warranty_only(rows, multiplier))
            continue
        if name == "PL Print":
            processed.append(ProcessedRow(RowKind.SECTION, ["Prices by style"]))
            processed.extend(_process_pl_print(rows, multiplier))

    wb.close()
    return collapse_blank_rows(processed)


def fn_chairs_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            s = str(c)
            bad.extend(forbidden_phrases(s))
            if INTERNET_POLICY_RE.search(s):
                bad.append(s)
    return bad
