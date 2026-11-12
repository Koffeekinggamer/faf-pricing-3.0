"""Retail transform for wide wholesale books (e.g. Artisan Chairs)."""

from __future__ import annotations

import re
from typing import Any

from scripts.dealer_book import (
    FOB_RE,
    PREFERRED_RE,
    WHOLESALE_RE,
    ProcessedRow,
    RowKind,
    _format_retail_price,
    _sanitize_text,
    row_is_blank,
)

EFFECTIVE_RE = re.compile(
    r"prices\s+effective\s+starting\s+"
    r"(January|February|March|April|May|June|July|August|September|October|November|December)"
    r"\s+(\d{1,2}),\s*(\d{4})",
    re.I,
)
LOGIN_RE = re.compile(r"username\s*:|password\s*:", re.I)


def is_artisan_style_grid(grid: list[list]) -> bool:
    for row in grid[:40]:
        for c in row:
            if isinstance(c, str) and "finished and unfinished" in c.lower():
                return True
    return False


def extract_effective_label(grid: list[list]) -> str | None:
    for row in grid:
        for c in row:
            if not isinstance(c, str):
                continue
            m = EFFECTIVE_RE.search(c)
            if m:
                return f"Effective {m.group(1)} {int(m.group(2))}, {m.group(3)}"
    return None


def retail_title(builder_name: str, effective: str | None) -> str:
    if effective:
        return f"{builder_name} Retail Price List · {effective}"
    return f"{builder_name} Retail Price List"


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    return str(c).strip()


def _joined(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def _max_used_col(grid: list[list]) -> int:
    m = 0
    for row in grid:
        for j, c in enumerate(row):
            if _cell_str(c):
                m = max(m, j)
    return m


def _is_add_row(row: list) -> bool:
    return any(_cell_str(c).upper() == "ADD" for c in row)


def _is_price_product_row(row: list) -> bool:
    if not _cell_str(row[0] if row else ""):
        return False
    nums = [
        c
        for j, c in enumerate(row)
        if j >= 3
        and isinstance(c, (int, float))
        and not isinstance(c, bool)
        and c >= 15
    ]
    return len(nums) >= 2


def _transform_wide_cells(
    row: list, multiplier: float, *, is_add: bool, is_price: bool, width: int
) -> list[str]:
    out: list[str] = []
    for j in range(width + 1):
        c = row[j] if j < len(row) else None
        if c is None or c == "":
            out.append("")
            continue
        if isinstance(c, (int, float)) and not isinstance(c, bool):
            if is_add and j == 14:
                out.append(_format_retail_price(c, multiplier))
            elif is_price and j >= 3:
                out.append(_format_retail_price(c, multiplier))
            else:
                out.append("")
            continue
        t = _sanitize_text(_cell_str(c), multiplier)
        if t.upper() == "ADD":
            out.append("ADD")
        elif t:
            out.append(t)
        else:
            out.append("")
    return out


def artisan_forbidden_text(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            s = str(c).lower()
            if "wholesale" in s:
                bad.append(str(c))
            if "preferred dealer" in s:
                bad.append(str(c))
            if FOB_RE.search(s):
                bad.append(str(c))
            if LOGIN_RE.search(s):
                bad.append(str(c))
    return bad


def process_artisan_book(
    grid: list[list],
    *,
    builder_name: str,
    multiplier: float,
) -> tuple[list[ProcessedRow], str]:
    effective = extract_effective_label(grid)
    doc_title = retail_title(builder_name, effective)
    width = _max_used_col(grid)

    out: list[ProcessedRow] = []
    title_written = False

    for row in grid:
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue

        joined = _joined(row)
        if FOB_RE.search(joined) or LOGIN_RE.search(joined):
            continue

        if WHOLESALE_RE.search(joined) and "price list" in joined.lower():
            if not title_written:
                out.append(ProcessedRow(RowKind.TITLE, [doc_title]))
                title_written = True
            continue

        if EFFECTIVE_RE.search(joined):
            continue

        if "finished and unfinished" in joined.lower():
            text = _sanitize_text(joined, multiplier)
            out.append(ProcessedRow(RowKind.SUBHEAD, [text]))
            continue

        if _is_add_row(row):
            out.append(
                ProcessedRow(
                    RowKind.WIDE,
                    _transform_wide_cells(
                        row, multiplier, is_add=True, is_price=False, width=width
                    ),
                )
            )
            continue

        if _is_price_product_row(row):
            out.append(
                ProcessedRow(
                    RowKind.WIDE,
                    _transform_wide_cells(
                        row, multiplier, is_add=False, is_price=True, width=width
                    ),
                )
            )
            continue

        filled = [(j, c) for j, c in enumerate(row) if _cell_str(c)]
        if len(filled) == 1 and filled[0][0] == 0:
            text = _sanitize_text(_cell_str(filled[0][1]), multiplier)
            if text and len(text) < 60 and not text.lower().startswith("options"):
                out.append(ProcessedRow(RowKind.SECTION, [text]))
                continue

        has_num = any(
            isinstance(c, (int, float)) and not isinstance(c, bool) and c >= 10
            for c in row
        )
        if not has_num and len(filled) >= 2:
            cells = _transform_wide_cells(
                row, multiplier, is_add=False, is_price=False, width=width
            )
            if any(cells):
                out.append(ProcessedRow(RowKind.WIDE, cells))
            continue

        if "unfinished" in joined.lower():
            out.append(
                ProcessedRow(
                    RowKind.WIDE,
                    _transform_wide_cells(
                        row, multiplier, is_add=False, is_price=False, width=width
                    ),
                )
            )
            continue

        cells = []
        for c in row:
            if c is None or isinstance(c, (int, float)):
                continue
            t = _sanitize_text(_cell_str(c), multiplier)
            if t:
                cells.append(t)
        if cells:
            out.append(ProcessedRow(RowKind.PROSE, cells))

    return out, doc_title
