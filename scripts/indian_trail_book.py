"""Indian Trail Furniture — WHOLESALE sheet only (RETAIL equals wholesale; markup 0)."""

from __future__ import annotations

import re
from pathlib import Path

from scripts.dealer_book import ProcessedRow, RowKind
from scripts.viztech_book import MULTIPLIER_DEFAULT, ViztechBuildResult, load_workbook_sheets
from scripts.viztech_wholesale_grid import (
    build_viztech_result,
    process_finished_wholesale_sheet,
    viztech_forbidden,
)

BUILDER = "Indian Trail Furniture"
DOC_TITLE = "Indian Trail Furniture Retail Price List · Pricing in Groups 2026"
# Finished wood columns only (0-based); cols 8–10 in XLSX; skip unfinished D–F.
FINISHED_COLS = (7, 8, 9)

DROP = (
    re.compile(r"to our valued dealers", re.I),
    re.compile(r"terms:\s*2%\s*after\s*30", re.I),
    re.compile(r"william\s*&\s*phyllis|troyer", re.I),
    re.compile(r"alan jay|john jay|lavern jay", re.I),
    re.compile(r"purchase order", re.I),
    re.compile(r"finish shop", re.I),
    re.compile(r"viztech", re.I),
    re.compile(r"net\s*30", re.I),
)


def is_indian_trail_workbook(path: Path) -> bool:
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return False
    try:
        sheets = load_workbook_sheets(path)
        return "WHOLESALE" in sheets and "RETAIL" in sheets
    except Exception:
        return False


def process_indian_trail_workbook(
    path: Path,
    *,
    multiplier: float = MULTIPLIER_DEFAULT,
) -> ViztechBuildResult:
    sheets = load_workbook_sheets(path)
    rows_raw = sheets.get("WHOLESALE")
    if not rows_raw:
        return ViztechBuildResult([], [], list(sheets), ["Missing WHOLESALE sheet"])
    spots: list = []
    body = process_finished_wholesale_sheet(
        rows_raw,
        finished_cols=FINISHED_COLS,
        multiplier=multiplier,
        drop_patterns=DROP,
        spots=spots,
    )
    title = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]
    edition = [
        r
        for r in body
        if r.kind == RowKind.PROSE
        and "pricing in groups 2026" in str(r.cells[0]).lower()
    ]
    if not edition:
        for row in rows_raw:
            joined = " ".join(str(c) for c in row if c)
            if "pricing in groups 2026" in joined.lower():
                title.append(
                    ProcessedRow(RowKind.SUBHEAD, ["Pricing in Groups 2026"])
                )
                break
    return build_viztech_result(title + body, spots, skipped=["RETAIL", "Master"])