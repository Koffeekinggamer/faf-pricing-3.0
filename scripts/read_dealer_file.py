"""Load dealer price files into row/column grids."""

from __future__ import annotations

import csv
from pathlib import Path

import openpyxl
import pdfplumber

try:
    import xlrd
except ImportError:
    xlrd = None  # type: ignore


class DealerFileError(Exception):
    pass


def _grid_from_rows(rows: list[list]) -> list[list]:
    return [[cell if cell is not None else "" for cell in row] for row in rows]


def read_csv(path: Path) -> list[list]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        return _grid_from_rows(list(csv.reader(f)))


def read_xlsx(path: Path, sheet_name: str | None = None) -> list[list]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    elif "Wholesale" in wb.sheetnames:
        ws = wb["Wholesale"]
    else:
        ws = wb["Pricelist"] if "Pricelist" in wb.sheetnames else wb.active
    rows = [list(row) for row in ws.iter_rows(values_only=True)]
    wb.close()
    if not rows:
        raise DealerFileError(f"No rows found in {path}")
    return _grid_from_rows(rows)


def read_xls(path: Path) -> list[list]:
    if xlrd is None:
        raise DealerFileError("xlrd is required for .xls files")
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    rows = [sheet.row_values(r) for r in range(sheet.nrows)]
    if not rows:
        raise DealerFileError(f"No rows found in {path}")
    return _grid_from_rows(rows)


def read_pdf_as_grid(path: Path) -> list[list]:
    tables: list[list[list]] = []
    lines: list[str] = []
    with pdfplumber.open(path) as pdf:
        if not pdf.pages:
            raise DealerFileError(
                f"PDF {path} has no pages. Export a spreadsheet or use a text-based PDF."
            )
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if table and any(any(c for c in row) for row in table):
                    tables.append(table)
            text = page.extract_text() or ""
            if text.strip():
                lines.extend(text.splitlines())

    if tables:
        best = max(tables, key=lambda t: len(t) * max(len(r) for r in t))
        return _grid_from_rows(best)

    if lines:
        return [[line] for line in lines]

    raise DealerFileError(
        f"Could not extract text or tables from PDF {path}. "
        "This may be a scanned image PDF — convert to .xlsx or .csv, or obtain a digital export."
    )


def read_dealer_file(path: Path, sheet_name: str | None = None) -> list[list]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return read_csv(path)
    if suffix == ".xlsx":
        return read_xlsx(path, sheet_name=sheet_name)
    if suffix == ".xls":
        return read_xls(path)
    if suffix == ".pdf":
        return read_pdf_as_grid(path)
    raise DealerFileError(f"Unsupported file type: {suffix}")


def is_dealer_book_grid(grid: list[list]) -> bool:
    for row in grid[:80]:
        for c in row:
            if isinstance(c, str) and "price list index" in c.lower():
                return True
    return False


def apply_multiplier_to_grid(grid: list[list], multiplier: float) -> list[list]:
    from scripts.pricing_lib import is_price_column, transform_cell_value

    if not grid:
        return grid

    header_row_idx = 0
    headers: list[str | None] = []
    for i, row in enumerate(grid[:5]):
        text_cells = sum(1 for c in row if str(c).strip())
        if text_cells >= 2:
            header_row_idx = i
            headers = [str(c).strip() if c else None for c in row]
            break

    out: list[list] = []
    for ridx, row in enumerate(grid):
        if ridx < header_row_idx:
            out.append(row)
            continue
        new_row = []
        for cidx, cell in enumerate(row):
            header = headers[cidx] if cidx < len(headers) else None
            price_col = is_price_column(header) if ridx > header_row_idx else False
            new_row.append(
                transform_cell_value(
                    cell, multiplier, header, force_price=price_col
                )
            )
        out.append(new_row)
    return out
