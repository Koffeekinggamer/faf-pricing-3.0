"""J&R Woodworking — wholesale from '2025 Price List markup' (markup=1), finished columns only."""

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

BUILDER = "J&R Woodworking"
DOC_TITLE = "J&R Woodworking Retail Price List · 2026"
PRICE_SHEET = "2025 Price List markup"
# Finished Spectrum columns (0-based indices 10–13); unfinished 7–9 skipped.
FINISHED_COLS = (10, 11, 12, 13)

DROP = (
    re.compile(r"miller\s*&\s*family|jerry\s*&\s*rosetta", re.I),
    re.compile(r"lavon,\s*jason|janey", re.I),
    re.compile(r"net\s*30", re.I),
    re.compile(r"2%\s*per\s*month", re.I),
    re.compile(r"lead time is 6-8", re.I),
    re.compile(r"finish shop", re.I),
    re.compile(r"viztech\s*password", re.I),
    re.compile(r"invoiced from|payment terms", re.I),
    re.compile(r"@[a-z0-9.-]+\.[a-z]{2,}", re.I),
)


def is_jr_woodworking_workbook(path: Path) -> bool:
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return False
    try:
        sheets = load_workbook_sheets(path)
        return PRICE_SHEET in sheets
    except Exception:
        return False


def process_jr_woodworking_workbook(
    path: Path,
    *,
    multiplier: float = MULTIPLIER_DEFAULT,
) -> ViztechBuildResult:
    sheets = load_workbook_sheets(path)
    rows_raw = sheets.get(PRICE_SHEET)
    if not rows_raw:
        return ViztechBuildResult([], [], list(sheets), [f"Missing {PRICE_SHEET}"])
    spots: list = []
    body = process_finished_wholesale_sheet(
        rows_raw,
        finished_cols=FINISHED_COLS,
        multiplier=multiplier,
        drop_patterns=DROP,
        spots=spots,
    )
    all_rows = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])] + body
    skip = [n for n in sheets if re.search(r"backup|\(bk\)", n, re.I)]
    skip.append("2025 Price List")
    return build_viztech_result(all_rows, spots, skipped=skip)