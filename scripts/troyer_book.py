"""J. Troyer and Co: four workbooks → one retail PDF."""

from __future__ import annotations

import re
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Collection

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.retail_text import forbidden_phrases, retail_price_cell, sanitize_retail_text

BUILDER = "J. Troyer and Co"
DOC_TITLE = "J. Troyer and Co Retail Price List"
FACTOR_VALUES = frozenset({1.0, 1.03, 1.05, 1.06, 1.1, 1.35, 1.85})
WORKBOOK_ORDER = ("troyer_2.xlsx", "troyer_3.xlsx", "troyer_1.xlsx", "troyer_4.xlsx")
SKIP_SHEETS = frozenset(
    {
        "MARKUP",
        "Markup",
        "Cover",
        "Index",
        "Master",
        "3% Increase Items",
        "Finley",
        "Easton",
        "Lincoln",
        "Northport",
    }
)
OPTIONS_SHEETS = frozenset({"Options", "2 Options"})
OVERLAP_FROM_T4 = frozenset(
    {
        "cassidy",
        "colorado river edge",
        "evan",
        "fiona",
        "rory",
        "alcova",
    }
)
DEALER_RE = re.compile(
    r"millersburg|674\.1050|jtroyercompany|wholesale price list",
    re.I,
)
FORMULA_RE = re.compile(r"formula line don't edit", re.I)
ERROR_TOKENS = frozenset({"#REF!", "#N/A", "#VALUE!", "#DIV/0!"})


def is_j_troyer_bundle(path: Path) -> bool:
    if path.is_dir():
        return (path / "troyer_2.xlsx").is_file()
    if path.suffix.lower() == ".zip" and path.is_file():
        with zipfile.ZipFile(path, "r") as zf:
            return any(n.endswith("troyer_2.xlsx") for n in zf.namelist())
    return path.name.startswith("troyer_") and path.suffix.lower() == ".xlsx"


def _resolve_dir(path: Path) -> Path:
    if path.is_dir():
        return path
    if path.suffix.lower() == ".zip":
        base = path.parent / f"{path.stem}_jt_extract"
        base.mkdir(exist_ok=True)
        with zipfile.ZipFile(path, "r") as zf:
            zf.extractall(base)
        return base
    return path.parent


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    return str(c).strip()


def _joined(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def _is_error(val: Any) -> bool:
    return _cell_str(val).upper() in ERROR_TOKENS


def _is_factor(val: Any) -> bool:
    return (
        isinstance(val, (int, float))
        and not isinstance(val, bool)
        and float(val) in FACTOR_VALUES
    )


def _is_price(val: Any) -> bool:
    if _is_error(val) or _is_factor(val):
        return False
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return False
    return float(val) > 0


def _truncate_before_mirror(row: list) -> list:
    price_idxs = [j for j, c in enumerate(row) if _is_price(c)]
    if len(price_idxs) < 2:
        return row
    start = price_idxs[0]
    block: list[int] = [start]
    for j in price_idxs[1:]:
        if j == block[-1] + 1 or (not block and j == start):
            block.append(j)
        elif j <= block[-1] + 3:
            block.append(j)
        else:
            break
    if len(block) < 2:
        end = max(price_idxs) + 1
        return row[: end + 1]
    seg = [row[j] for j in block]
    for j in range(block[-1] + 1, len(row) - len(block) + 1):
        chunk = row[j : j + len(block)]
        if chunk == seg:
            return row[:j]
    return row[: block[-1] + 1]


def _transform_row(row: list, multiplier: float) -> list[str]:
    row = _truncate_before_mirror(list(row))
    price_cols: set[int] = set()
    for j, c in enumerate(row):
        if _is_price(c):
            price_cols.add(j)
    width = max(
        (j for j in range(len(row)) if _cell_str(row[j]) or j in price_cols),
        default=0,
    )
    out: list[str] = []
    for j in range(width + 1):
        c = row[j] if j < len(row) else None
        if _is_error(c) or (c == 0 or c == 0.0):
            out.append("")
        elif j in price_cols:
            out.append(retail_price_cell(c, multiplier))
        elif isinstance(c, (int, float)) and not isinstance(c, bool):
            if _is_factor(c):
                out.append("")
            else:
                out.append(_cell_str(c))
        else:
            t = sanitize_retail_text(_cell_str(c), multiplier) if c not in (None, "") else ""
            if DEALER_RE.search(t):
                t = ""
            out.append(t)
    while out and not out[-1]:
        out.pop()
    return out


def _overlap_key(sheet: str) -> str | None:
    s = sheet.lower().replace("occasional table", "").strip()
    for key in OVERLAP_FROM_T4:
        if key in s:
            return key
    return None


def _should_skip_t4_sheet(sheet: str, t1_newer: bool) -> bool:
    if not t1_newer:
        return False
    key = _overlap_key(sheet)
    return key is not None and sheet != "Occasional Collection"


def _parse_effective_month(text: str) -> datetime | None:
    text = text.lower()
    m = re.search(r"effective[:\s]*([a-z]+)\s*,?\s*(\d{4})", text)
    if not m:
        m = re.search(r"effective[:\s]*([a-z]+)\s+(\d{4})", text)
    if not m:
        return None
    months = {
        "january": 1,
        "february": 2,
        "march": 3,
        "april": 4,
        "may": 5,
        "june": 6,
        "july": 7,
        "august": 8,
        "september": 9,
        "october": 10,
        "november": 11,
        "december": 12,
    }
    mon = months.get(m.group(1)[:3] if len(m.group(1)) > 3 else m.group(1))
    if not mon:
        mon = months.get(m.group(1))
    if not mon:
        return None
    return datetime(int(m.group(2)), mon, 1)


def _workbook_effective(path: Path) -> datetime | None:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "Cover" not in wb.sheetnames:
        wb.close()
        return None
    texts: list[str] = []
    for row in wb["Cover"].iter_rows(values_only=True):
        for c in row:
            if c:
                texts.append(str(c))
    wb.close()
    for t in texts:
        dt = _parse_effective_month(t)
        if dt:
            return dt
    return None


def troyer_overlap_report(base: Path) -> str:
    t1 = _workbook_effective(base / "troyer_1.xlsx")
    t4 = _workbook_effective(base / "troyer_4.xlsx")
    kept = "troyer_1 (October 2026)" if (t1 and t4 and t1 > t4) else "troyer_4"
    return (
        f"Overlapping collections (Cassidy, Alcova, Evan, Fiona, Rory, Colorado River Edge): "
        f"kept {kept}; e.g. Cassidy M1284 wholesale "
        f"{'1034' if t1 and t4 and t1 > t4 else '969'} from {kept.split()[0]}."
    )


def _process_workbook(
    path: Path,
    multiplier: float,
    *,
    skip_overlap_t4: bool,
) -> list[ProcessedRow]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out: list[ProcessedRow] = []
    for ws in wb.worksheets:
        title = ws.title
        if ws.sheet_state != "visible":
            continue
        if title in SKIP_SHEETS and title not in OPTIONS_SHEETS:
            continue
        if path.name == "troyer_4.xlsx" and title.lower() == "options":
            continue
        if skip_overlap_t4 and _should_skip_t4_sheet(title, True):
            continue
        out.append(ProcessedRow(RowKind.SECTION, [sanitize_retail_text(title, multiplier)]))
        for row in ws.iter_rows(values_only=True):
            row = list(row)
            if row_is_blank(row):
                out.append(ProcessedRow(RowKind.BLANK, []))
                continue
            if FORMULA_RE.search(_joined(row)):
                continue
            if DEALER_RE.search(_joined(row)) and "deduct" not in _joined(row).lower():
                continue
            cells = _transform_row(row, multiplier)
            if not cells:
                continue
            if len(cells) == 1 and len(str(cells[0])) > 80:
                out.append(ProcessedRow(RowKind.PROSE, cells))
            else:
                out.append(ProcessedRow(RowKind.WIDE, cells))
    wb.close()
    return out


def process_j_troyer_bundle(path: Path, *, multiplier: float) -> list[ProcessedRow]:
    base = _resolve_dir(path)
    t1 = _workbook_effective(base / "troyer_1.xlsx")
    t4 = _workbook_effective(base / "troyer_4.xlsx")
    t1_newer = bool(t1 and t4 and t1 > t4)
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]
    for name in WORKBOOK_ORDER:
        xlsx = base / name
        if not xlsx.is_file():
            raise FileNotFoundError(name)
        skip_t4 = name == "troyer_4.xlsx" and t1_newer
        out.extend(_process_workbook(xlsx, multiplier, skip_overlap_t4=skip_t4))
    return collapse_blank_rows(out)


def j_troyer_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            s = str(c)
            bad.extend(forbidden_phrases(s))
            if re.search(r"\bwholesale\b", s, re.I):
                bad.append(s)
            if any(tok in s for tok in ERROR_TOKENS):
                bad.append(s)
    return bad
