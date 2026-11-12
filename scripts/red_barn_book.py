"""Red Barn Woodworking .xls workbook → one retail PDF."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any, Collection

import xlrd

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.retail_text import forbidden_phrases, retail_price_cell, sanitize_retail_text

DOC_TITLE = "Red Barn Woodworking Retail Price List"
WOOD_FACTORS = frozenset({1.1, 1.35, 1.85, 1.05})
SKIP_SHEETS = frozenset(
    {
        "Markup",
        "Options 1",
        "Options 2",
        "Options 4",
        "Options2",
        "Master",
        "Master2",
        "Master3",
        "VLOOKUP Increase Log",
    }
)
DEALER_RE = re.compile(
    r"wholesale price list|dalton|zuercher|44618|330-698|fax:",
    re.I,
)
ITEM_ID_RE = re.compile(
    r"^(C\d|EH\d|GO-|LA-|BM-|#|\d)",
    re.I,
)


def is_red_barn_workbook(path: Path) -> bool:
    xls = _resolve_xls(path)
    if xls is None:
        return False
    try:
        book = xlrd.open_workbook(str(xls))
        return "Markup" in book.sheet_names() and "Options 3" in book.sheet_names()
    except Exception:
        return False


def _resolve_xls(path: Path) -> Path | None:
    if path.suffix.lower() == ".xls" and path.is_file():
        return path
    if path.suffix.lower() == ".zip" and path.is_file():
        base = path.parent / f"{path.stem}_rb_extract"
        base.mkdir(exist_ok=True)
        with zipfile.ZipFile(path, "r") as zf:
            zf.extractall(base)
        for p in base.rglob("*.xls"):
            return p
    if path.is_dir():
        for p in path.rglob("*.xls"):
            return p
    return None


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    return str(c).strip()


def _joined(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def _looks_like_item_id(s: str) -> bool:
    s = s.strip()
    if not s or len(s) > 40:
        return False
    return bool(ITEM_ID_RE.match(s))


def _is_factor(val: Any) -> bool:
    return isinstance(val, (int, float)) and not isinstance(val, bool) and float(val) in WOOD_FACTORS


def _is_dimension_cell(j: int, val: Any) -> bool:
    if j not in (1, 2, 3):
        return False
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return 0 < float(val) < 120
    if isinstance(val, str) and re.search(r"\d+\s*[wWdDhH]", val):
        return True
    return False


def _price_columns_for_row(row: list) -> Collection[int]:
    prices: list[int] = []
    for j, c in enumerate(row):
        if _is_dimension_cell(j, c) or _is_factor(c):
            continue
        if isinstance(c, (int, float)) and not isinstance(c, bool) and float(c) > 10:
            prices.append(j)
    # Drop mirrored duplicate block (cols 10+ when 4-7 exist and equal)
    left = [row[j] for j in range(4, 8) if j < len(row)]
    if len(left) == 4 and all(isinstance(x, (int, float)) for x in left):
        mirror = [row[j] for j in range(10, 14) if j < len(row)]
        if mirror == left:
            prices = [j for j in prices if j < 9]
    return prices


def _primary_item_ids(rows: list[list]) -> set[str]:
    ids: set[str] = set()
    for row in rows:
        c0 = _cell_str(row[0])
        if c0 and _looks_like_item_id(c0.split()[0]):
            ids.add(c0.split()[0])
    return ids


def _secondary_block(row: list, primary_ids: set[str]) -> list | None:
    """Return a synthetic left-style row for a right-side item list, or None."""
    for c in range(8, min(len(row), 12)):
        cid = _cell_str(row[c])
        if not _looks_like_item_id(cid):
            continue
        key = cid.split()[0]
        if key in primary_ids:
            return None
        desc = _cell_str(row[c + 1]) if c + 1 < len(row) else ""
        nums = [
            (j, row[j])
            for j in range(c + 2, len(row))
            if isinstance(row[j], (int, float))
            and not isinstance(row[j], bool)
            and float(row[j]) > 10
        ]
        if len(nums) < 4:
            return None
        price_idxs = [j for j, _ in nums[:4]]
        out = [key, desc, "", ""] + [row[j] for j in price_idxs]
        return out
    return None


def _trim_office_mirror(row: list) -> list:
    left = _cell_str(row[0])
    if len(row) > 10 and left and _cell_str(row[10]) == left:
        return row[:10]
    return row


def _transform_row(row: list, multiplier: float) -> list[str]:
    row = _trim_office_mirror(row)
    price_cols = set(_price_columns_for_row(row))
    width = max(
        (j for j in range(len(row)) if _cell_str(row[j]) or j in price_cols),
        default=0,
    )
    out: list[str] = []
    for j in range(width + 1):
        c = row[j] if j < len(row) else None
        if j in price_cols and isinstance(c, (int, float)) and not isinstance(c, bool):
            out.append(retail_price_cell(c, multiplier))
        elif _is_factor(c):
            out.append("")
        elif isinstance(c, (int, float)) and not isinstance(c, bool):
            if _is_dimension_cell(j, c):
                out.append(_cell_str(c))
            else:
                out.append("")
        else:
            t = sanitize_retail_text(_cell_str(c), multiplier) if c not in (None, "") else ""
            out.append(t)
    while out and not out[-1]:
        out.pop()
    return out


def _process_options(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out = [ProcessedRow(RowKind.SECTION, ["Finish options"])]
    for row in rows:
        if row_is_blank(row):
            continue
        joined = _joined(row)
        if "customizing chart" in joined.lower() and "330-388" in joined:
            continue
        if "call 330" in joined.lower():
            continue
        cells = _transform_row(row, multiplier)
        if cells:
            out.append(ProcessedRow(RowKind.WIDE, cells))
    return out


def _process_collection(
    name: str, rows: list[list], multiplier: float
) -> list[ProcessedRow]:
    out = [ProcessedRow(RowKind.SECTION, [sanitize_retail_text(name, multiplier)])]
    primary_ids = _primary_item_ids(rows)
    for row in rows:
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        if DEALER_RE.search(_joined(row)):
            continue
        if _joined(row) in ("Depth", "Width", "Height") and not _cell_str(row[0]):
            continue
        if all(_is_factor(c) for c in row if isinstance(c, (int, float))):
            continue
        cells = _transform_row(row, multiplier)
        if cells:
            if len(cells) == 1 and len(str(cells[0])) > 60:
                out.append(ProcessedRow(RowKind.PROSE, cells))
            else:
                out.append(ProcessedRow(RowKind.WIDE, cells))
        sec = _secondary_block(row, primary_ids)
        if sec:
            sec_cells = _transform_row(sec, multiplier)
            if sec_cells:
                out.append(ProcessedRow(RowKind.WIDE, sec_cells))
    return out


def process_red_barn_workbook(path: Path, *, multiplier: float) -> list[ProcessedRow]:
    xls = _resolve_xls(path)
    if xls is None:
        raise FileNotFoundError("Red Barn .xls not found")
    book = xlrd.open_workbook(str(xls))
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]

    opt = book.sheet_by_name("Options 3")
    opt_rows = [list(opt.row_values(i)) for i in range(opt.nrows)]
    out.extend(_process_options(opt_rows, multiplier))

    for name in book.sheet_names():
        if name in SKIP_SHEETS or name == "Options 3":
            continue
        sh = book.sheet_by_name(name)
        rows = [list(sh.row_values(i)) for i in range(sh.nrows)]
        out.extend(_process_collection(name, rows, multiplier))

    return collapse_blank_rows(out)


def red_barn_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            s = str(c)
            bad.extend(forbidden_phrases(s))
            if re.search(r"\bwholesale\b", s, re.I):
                bad.append(s)
    return bad
