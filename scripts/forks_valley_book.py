"""Forks Valley Woodworking: combine four dealer files into one retail PDF."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any, Collection

import openpyxl
import xlrd

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.pricing_lib import retail_whole_dollar
from scripts.retail_text import forbidden_phrases, retail_price_cell, sanitize_retail_text

DOC_TITLE = "Forks Valley Woodworking Retail Price List · 2026"
COLOR_BAKED_MARKUP = 1.08  # workbook stores base_upcharge * 1.08 when Mark-Up = 1
SEARCH_SKIP_RE = re.compile(
    r"item search|search houston|search all options|^total$",
    re.I,
)
DEALER_INFO_RE = re.compile(
    r"middlebury|terms\s*:|net\s*\d|discount|std\s*terms|"
    r"\d{3}[\s\-]?\d{3}[\s\-]?\d{4}|@forksvalley",
    re.I,
)
ID_ROW_RE = re.compile(r"^ID\s*NO\.?$", re.I)
SECTION_BREAK_RE = re.compile(r"^\s*-+\s*continued\s*$|^[A-Z][A-Z\s/&-]+$", re.I)


def is_forks_valley_bundle(path: Path) -> bool:
    if path.is_dir():
        return any(path.glob("*main_price_book*.xls"))
    if path.suffix.lower() == ".zip" and path.is_file():
        with zipfile.ZipFile(path, "r") as zf:
            return any("main_price_book" in n for n in zf.namelist())
    return False


def resolve_forks_sources(path: Path) -> list[Path]:
    if path.is_dir():
        base = path
    else:
        base = path.parent / "_forks_extract"
        base.mkdir(exist_ok=True)
        with zipfile.ZipFile(path, "r") as zf:
            zf.extractall(base)
    order = [
        "3_main_price_book.xls",
        "4_quick_ship.xlsx",
        "1_glass_mullion_options.xls",
        "2_special_color_upcharge.xls",
    ]
    out = []
    for name in order:
        p = base / name
        if not p.is_file():
            raise FileNotFoundError(f"Missing Forks Valley file: {name}")
        out.append(p)
    return out


def _xls_rows(path: Path, sheet: int | str = 0) -> list[list]:
    book = xlrd.open_workbook(str(path))
    sh = book.sheet_by_name(sheet) if isinstance(sheet, str) else book.sheet_by_index(sheet)
    return [list(sh.row_values(i)) for i in range(sh.nrows)]


def _xlsx_rows(path: Path, sheet: str) -> list[list]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    s = str(c).strip()
    return s


def _is_std(c: Any) -> bool:
    return _cell_str(c).upper() == "STD"


def _transform_price(value: Any, multiplier: float, *, color_sheet: bool = False) -> str:
    if _is_std(value):
        return "STD"
    if isinstance(value, str) and value.upper() in ("N/A", "NA"):
        return "N/A"
    if color_sheet and isinstance(value, (int, float)) and not isinstance(value, bool):
        base = float(value) / COLOR_BAKED_MARKUP
        return f"${retail_whole_dollar(base, multiplier):,}"
    return retail_price_cell(value, multiplier)


def _row_to_wide(
    row: list,
    multiplier: float,
    width: int,
    *,
    price_cols: Collection[int],
    color_sheet: bool = False,
) -> list[str]:
    out: list[str] = []
    for j in range(width + 1):
        c = row[j] if j < len(row) else None
        if j in price_cols and isinstance(c, (int, float)) and not isinstance(c, bool):
            out.append(_transform_price(c, multiplier, color_sheet=color_sheet))
        elif isinstance(c, (int, float)) and not isinstance(c, bool):
            out.append("")
        elif c is not None and c != "":
            out.append(sanitize_retail_text(_cell_str(c), multiplier))
        else:
            out.append("")
    while out and not out[-1]:
        out.pop()
    return out


def _process_information(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out: list[ProcessedRow] = []
    out.append(ProcessedRow(RowKind.TITLE, [DOC_TITLE]))
    out.append(ProcessedRow(RowKind.SECTION, ["Information"]))
    for row in rows:
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        joined = " ".join(_cell_str(c) for c in row)
        if DEALER_INFO_RE.search(joined):
            continue
        if "wholesale" in joined.lower():
            continue
        cells = []
        for c in row:
            if c is None or isinstance(c, (int, float)):
                continue
            t = sanitize_retail_text(_cell_str(c), multiplier)
            if t:
                cells.append(t)
        if cells:
            out.append(ProcessedRow(RowKind.PROSE, cells))
    return out


def _skip_search_row(row: list) -> bool:
    joined = " ".join(_cell_str(c) for c in row)
    if SEARCH_SKIP_RE.search(joined):
        return True
    if _cell_str(row[2] if len(row) > 2 else "") == "Total":
        return True
    return False


def _finished_wholesale_start(rows: list[list]) -> int:
    for i, row in enumerate(rows):
        if _cell_str(row[0]) == "LIVING ROOM OCCASIONALS":
            return i
        if ID_ROW_RE.match(_cell_str(row[0])) and "Description" in _cell_str(row[1]):
            if i > 20:
                return max(0, i - 1)
    return 27


def _process_finished_wholesale(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out: list[ProcessedRow] = []
    out.append(ProcessedRow(RowKind.SECTION, ["Finished product price list"]))
    start = _finished_wholesale_start(rows)
    price_cols = range(3, 7)
    width = 6
    for row in rows[start:]:
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        if _skip_search_row(row):
            continue
        joined = " ".join(_cell_str(c) for c in row)
        if "markup" in joined.lower() and len(joined) < 20:
            continue
        c0 = _cell_str(row[0])
        if c0 and not c0.startswith("FV") and ID_ROW_RE.match(c0) is False:
            if SECTION_BREAK_RE.match(c0) or "continued" in c0.lower():
                out.append(ProcessedRow(RowKind.SUBHEAD, [sanitize_retail_text(c0, multiplier)]))
                continue
        if "%" in joined and not c0.startswith("FV"):
            note = sanitize_retail_text(joined.replace("markup", ""), multiplier)
            if note:
                out.append(ProcessedRow(RowKind.PROSE, [note]))
            continue
        if c0.startswith("FV") or (c0 and row[1]):
            trimmed = row[: width + 1]
            cells = _row_to_wide(trimmed, multiplier, width, price_cols=price_cols)
            if any(cells):
                out.append(ProcessedRow(RowKind.WIDE, cells))
        elif c0 == "ID NO.":
            cells = _row_to_wide(row[: width + 1], multiplier, width, price_cols=range(99))
            out.append(ProcessedRow(RowKind.WIDE, cells))
    return out


def _process_glass(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out = [ProcessedRow(RowKind.SECTION, ["Glass & mullion options"])]
    width = 5
    price_cols = range(1, 6)
    for row in rows:
        if row_is_blank(row):
            continue
        joined = " ".join(_cell_str(c) for c in row)
        if "wholesale" in joined.lower() or "mark-up" in joined.lower():
            if "Glass" in joined or "Mullion" in joined:
                out.append(ProcessedRow(RowKind.SUBHEAD, [sanitize_retail_text(joined, multiplier)]))
            continue
        cells = _row_to_wide(row[: width + 1], multiplier, width, price_cols=price_cols)
        if any(cells):
            out.append(ProcessedRow(RowKind.WIDE, cells))
    return out


def _color_block_wide(
    row: list, start: int, multiplier: float
) -> list[str] | None:
    width = 4
    slice_end = start + width + 1
    block = row[start:slice_end]
    while len(block) < width + 1:
        block.append("")
    desc = _cell_str(block[0])
    has_price = any(
        isinstance(block[j], (int, float)) and not isinstance(block[j], bool)
        for j in range(1, 5)
    )
    if not desc and not has_price:
        return None
    rel_prices = range(1, 5)
    out: list[str] = []
    for j in range(width + 1):
        c = block[j] if j < len(block) else None
        abs_j = start + j
        if j in rel_prices and isinstance(c, (int, float)) and not isinstance(c, bool):
            out.append(_transform_price(c, multiplier, color_sheet=True))
        elif isinstance(c, (int, float)) and not isinstance(c, bool):
            out.append("")
        elif c is not None and c != "":
            out.append(sanitize_retail_text(_cell_str(c), multiplier))
        else:
            out.append("")
    while out and not out[-1]:
        out.pop()
    return out if (desc or has_price) else None


def _process_color(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out = [ProcessedRow(RowKind.SECTION, ["Special color upcharges"])]
    for row in rows:
        if row_is_blank(row):
            continue
        joined = " ".join(_cell_str(c) for c in row)
        if "price sheet" in joined.lower() and "2026" in joined:
            continue
        if _cell_str(row[9] if len(row) > 9 else "") and "markup" in joined.lower():
            if len(joined) < 40:
                continue
        if "add these amounts" in joined.lower():
            continue
        for start in (0, 5):
            cells = _color_block_wide(row, start, multiplier)
            if cells:
                out.append(ProcessedRow(RowKind.WIDE, cells))
    return out


def _process_quick_ship(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out = [ProcessedRow(RowKind.SECTION, ["Quick ship finished items"])]
    started = False
    for row in rows:
        if row_is_blank(row):
            continue
        joined = " ".join(_cell_str(c) for c in row)
        if "wholesale" in joined.lower():
            continue
        if not started and _cell_str(row[0]) == "ID NO.":
            started = True
        if not started:
            if "Quick Ship" in joined:
                out.append(ProcessedRow(RowKind.SUBHEAD, [sanitize_retail_text(joined, multiplier)]))
            continue
        width = 4
        cells = _row_to_wide(row[: width + 1], multiplier, width, price_cols={4})
        if _cell_str(cells[0]).startswith("FV") or cells[0] == "ID NO.":
            out.append(ProcessedRow(RowKind.WIDE, cells))
    return out


def process_forks_valley_bundle(
    path: Path,
    *,
    multiplier: float,
) -> list[ProcessedRow]:
    sources = resolve_forks_sources(path)
    main_xls, quick_xlsx, glass_xls, color_xls = sources

    processed: list[ProcessedRow] = []

    info_rows = _xls_rows(main_xls, "Information Sheet")
    processed.extend(_process_information(info_rows, multiplier))

    fin_rows = _xls_rows(main_xls, "Finished Wholesale")
    processed.extend(_process_finished_wholesale(fin_rows, multiplier))

    qs_rows = _xlsx_rows(quick_xlsx, "Unfinished Wholesale")
    processed.extend(_process_quick_ship(qs_rows, multiplier))

    glass_rows = _xls_rows(glass_xls, "Glass Price List WHOLESALE")
    processed.extend(_process_glass(glass_rows, multiplier))

    color_rows = _xls_rows(color_xls, "Special Color Upcharge")
    processed.extend(_process_color(color_rows, multiplier))

    return collapse_blank_rows(processed)


def forks_valley_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            s = str(c)
            bad.extend(forbidden_phrases(s))
            if re.search(r"\bwholesale\b", s, re.I):
                bad.append(s)
    return bad
