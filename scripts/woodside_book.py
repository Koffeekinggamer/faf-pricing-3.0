"""Woodside Woodworks .xls → retail PDF (finished columns only)."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any

import xlrd

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.finished_only import UNFINISHED_NOTE_RE, drop_unfinished_note_text
from scripts.retail_text import forbidden_phrases, retail_price_cell, sanitize_retail_text

BUILDER = "Woodside Woodworks"
DOC_TITLE = "Woodside Woodworks Retail Price List"
SKIP_SHEETS = frozenset({"Paper Version", "Dashboard", "Conditions"})
FINISHED_START = 8
FINISHED_END = 13
MIRROR_START = 25
ERROR_VALS = frozenset({"#REF!", "#N/A", "#VALUE!", "#DIV/0!", "NA"})


def is_woodside_workbook(path: Path) -> bool:
    xls = _resolve_xls(path)
    if not xls:
        return False
    try:
        book = xlrd.open_workbook(str(xls))
        return "2026 price list" in book.sheet_names()
    except Exception:
        return False


def _resolve_xls(path: Path) -> Path | None:
    if path.suffix.lower() == ".xls" and path.is_file():
        return path
    if path.suffix.lower() == ".zip" and path.is_file():
        base = path.parent / f"{path.stem}_ws_extract"
        base.mkdir(exist_ok=True)
        with zipfile.ZipFile(path, "r") as zf:
            zf.extractall(base)
        for p in base.rglob("*.xls"):
            return p
    return None


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    s = str(c).strip()
    if s in ERROR_VALS:
        return ""
    return s


def _joined(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def _is_price(val: Any) -> bool:
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return False
    return float(val) > 0


def _trim_left_block(row: list) -> list:
    return list(row[:MIRROR_START])


def _effective_date(book: xlrd.book.Book) -> str:
    if "Paper Version" not in book.sheet_names():
        return DOC_TITLE
    sh = book.sheet_by_name("Paper Version")
    for i in range(sh.nrows):
        row = sh.row_values(i)
        for c in row:
            if isinstance(c, str) and "effective" in c.lower():
                nxt = row[row.index(c) + 1] if row.index(c) + 1 < len(row) else ""
                if nxt:
                    return f"{DOC_TITLE} · Effective {nxt}"
    return DOC_TITLE


def _transform_price_row(row: list, multiplier: float) -> list[str]:
    row = _trim_left_block(row)
    out: list[str] = []
    for j, c in enumerate(row):
        if FINISHED_START <= j < FINISHED_END and _is_price(c):
            out.append(retail_price_cell(c, multiplier))
        elif j < FINISHED_START:
            out.append(sanitize_retail_text(_cell_str(c), multiplier))
        elif j == FINISHED_END or (j > FINISHED_START and j < FINISHED_END + 3):
            if _is_price(c):
                out.append(retail_price_cell(c, multiplier))
    while out and not out[-1]:
        out.pop()
    return out


def _process_price_list(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out: list[ProcessedRow] = []
    for row in rows:
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue
        joined = _joined(row)
        if UNFINISHED_NOTE_RE.search(joined):
            continue
        if _cell_str(row[3] if len(row) > 3 else "") == "Unfinished":
            continue
        if "rule" in joined.lower() and "round" in joined.lower():
            continue
        nums_finished = [
            row[j]
            for j in range(FINISHED_START, min(len(row), FINISHED_END + 5))
            if _is_price(row[j])
        ]
        if nums_finished and (_cell_str(row[0]) or _cell_str(row[1])):
            cells = _transform_price_row(row, multiplier)
            if any(cells):
                out.append(ProcessedRow(RowKind.WIDE, cells))
            continue
        if re.search(r"add\s+\d+|add\s+\$", joined, re.I):
            cells = _transform_price_row(row, multiplier)
            if cells:
                out.append(ProcessedRow(RowKind.WIDE, cells))
            continue
        if joined and not joined.startswith("-"):
            note = drop_unfinished_note_text(joined)
            if note:
                prose = sanitize_retail_text(note, multiplier)
                if prose:
                    out.append(ProcessedRow(RowKind.PROSE, [prose]))
    return out


def _process_options(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out = [ProcessedRow(RowKind.SECTION, ["Options"])]
    finished_zone = False
    for row in rows:
        if row_is_blank(row):
            continue
        joined = _joined(row)
        if "standard finished" in joined.lower():
            finished_zone = True
            continue
        if not finished_zone:
            continue
        if UNFINISHED_NOTE_RE.search(joined) or _cell_str(row[3] if len(row) > 3 else "") == "Unfinished":
            continue
        nums = [
            (j, row[j])
            for j in range(FINISHED_START, min(len(row), MIRROR_START))
            if _is_price(row[j])
        ]
        if nums:
            cells = _transform_price_row(row, multiplier)
            if cells:
                out.append(ProcessedRow(RowKind.WIDE, cells))
    return out


def _process_pcl(rows: list[list], multiplier: float) -> list[ProcessedRow]:
    out = [ProcessedRow(RowKind.SECTION, ["PCL Color List"])]
    for row in rows:
        if row_is_blank(row):
            continue
        cells = [
            sanitize_retail_text(_cell_str(c), multiplier)
            for c in row
            if _cell_str(c) and "$" not in _cell_str(c)
        ]
        cells = [c for c in cells if c and not re.match(r"^\d+$", c)]
        if cells:
            out.append(ProcessedRow(RowKind.PROSE, cells[:8]))
    return out


def process_woodside_workbook(path: Path, *, multiplier: float) -> list[ProcessedRow]:
    xls = _resolve_xls(path)
    if not xls:
        raise FileNotFoundError("Woodside.xls not found")
    book = xlrd.open_workbook(str(xls))
    title = _effective_date(book)
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [title])]

    sh = book.sheet_by_name("2026 price list")
    rows = [list(sh.row_values(i)) for i in range(sh.nrows)]
    out.extend(_process_price_list(rows, multiplier))

    if "Options" in book.sheet_names():
        opt = book.sheet_by_name("Options")
        out.extend(_process_options([list(opt.row_values(i)) for i in range(opt.nrows)], multiplier))

    if "PCL Color List" in book.sheet_names():
        pcl = book.sheet_by_name("PCL Color List")
        out.extend(_process_pcl([list(pcl.row_values(i)) for i in range(pcl.nrows)], multiplier))

    return collapse_blank_rows(out)


def woodside_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            s = str(c)
            bad.extend(forbidden_phrases(s))
            if re.search(r"\bwholesale\b", s, re.I):
                bad.append(s)
    return bad
