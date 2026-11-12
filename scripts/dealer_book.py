"""Transform formatted dealer price books (e.g. Millcraft) into retail grids."""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum
from typing import Any

from scripts.pricing_lib import (
    assert_no_approx,
    retail_whole_dollar,
    transform_text_prices,
)

SKU_RE = re.compile(r"^[A-Z]{2,4}[0-9][A-Z0-9\-]*$", re.IGNORECASE)
INDEX_START_RE = re.compile(r"price\s*list\s*index", re.I)
INDEX_END_RE = re.compile(r"standard\s*options", re.I)
TITLE_DATE_RE = re.compile(
    r"(January|February|March|April|May|June|July|August|September|October|November|December)"
    r"\s+(\d{1,2})(?:st|nd|rd|th)?,\s*(\d{4})",
    re.I,
)
FOB_RE = re.compile(r"f\.?\s*o\.?\s*b", re.I)
WHOLESALE_RE = re.compile(r"\bwholesale\b", re.I)
PREFERRED_RE = re.compile(r"preferred\s+dealer", re.I)
PRICE_LIST_RE = re.compile(r"price\s*list", re.I)


class RowKind(str, Enum):
    BLANK = "blank"
    DROP = "drop"
    TITLE = "title"
    SECTION = "section"
    SUBHEAD = "subhead"
    TABLE_HEADER = "table_header"
    SKU = "sku"
    INDEX = "index"
    PROSE = "prose"
    WIDE = "wide"


@dataclass
class ProcessedRow:
    kind: RowKind
    cells: list[Any]


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    return str(c).strip()


def row_is_blank(row: list) -> bool:
    return not any(_cell_str(c) for c in row)


def _sanitize_text(text: str, multiplier: float) -> str:
    if not text:
        return text
    if FOB_RE.search(text):
        return ""
    t = PREFERRED_RE.sub("", text)
    t = WHOLESALE_RE.sub("", t)
    t = re.sub(r"\s{2,}", " ", t).strip(" ·-")
    t = re.sub(r"\s{2,}", " ", t).strip()
    if not t:
        return ""
    if "$" in t:
        t = transform_text_prices(t, multiplier)
    assert_no_approx(t)
    return t


def _format_retail_price(value: Any, multiplier: float) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value <= 0:
            return ""
        return f"${retail_whole_dollar(float(value), multiplier):,}"
    text = _cell_str(value)
    if not text:
        return ""
    if "$" in text:
        return transform_text_prices(text, multiplier)
    try:
        num = float(text.replace(",", "").replace("$", ""))
        if num > 0:
            return f"${retail_whole_dollar(num, multiplier):,}"
    except ValueError:
        pass
    return text


def extract_catalog_date(grid: list[list]) -> str | None:
    for row in grid:
        for c in row:
            if not isinstance(c, str):
                continue
            m = TITLE_DATE_RE.search(c)
            if m:
                from datetime import datetime

                month, day, year = m.group(1), int(m.group(2)), int(m.group(3))
                try:
                    dt = datetime.strptime(f"{month} {day} {year}", "%B %d %Y")
                    return dt.strftime("%B %-d, %Y")
                except ValueError:
                    return f"{month} {day}, {year}"
    return None


def retail_title(builder_name: str, catalog_date: str | None) -> str:
    if catalog_date:
        return f"{builder_name} Retail Price List · {catalog_date}"
    return f"{builder_name} Retail Price List"


def _is_sku(cell: Any) -> bool:
    s = _cell_str(cell)
    return bool(s and SKU_RE.match(s))


def _joined_text(row: list) -> str:
    return " ".join(_cell_str(c) for c in row if _cell_str(c))


def process_dealer_book(
    grid: list[list],
    *,
    builder_name: str,
    multiplier: float,
    catalog_date: str | None = None,
) -> tuple[list[ProcessedRow], str]:
    catalog_date = catalog_date or extract_catalog_date(grid)
    doc_title = retail_title(builder_name, catalog_date)

    in_index = False
    out: list[ProcessedRow] = []
    title_written = False

    for row in grid:
        joined = _joined_text(row)

        if INDEX_START_RE.search(joined):
            in_index = True
        if INDEX_END_RE.search(joined):
            in_index = False

        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            continue

        if FOB_RE.search(joined):
            continue

        sku = row[2] if len(row) > 2 else None
        if _is_sku(sku):
            desc = _sanitize_text(_cell_str(row[1]), multiplier)
            sku_s = _cell_str(sku)
            std = _format_retail_price(row[3] if len(row) > 3 else None, multiplier)
            prem = _format_retail_price(row[4] if len(row) > 4 else None, multiplier)
            if (
                out
                and out[-1].kind != RowKind.TABLE_HEADER
                and out[-1].kind != RowKind.SKU
            ):
                out.append(
                    ProcessedRow(
                        RowKind.TABLE_HEADER,
                        [
                            "Description",
                            "Item Number",
                            "Standard Wood",
                            "Premium Wood",
                        ],
                    )
                )
            out.append(ProcessedRow(RowKind.SKU, [desc, sku_s, std, prem]))
            continue

        lowered = joined.lower()
        if (
            not title_written
            and PRICE_LIST_RE.search(joined)
            and (WHOLESALE_RE.search(joined) or PREFERRED_RE.search(joined))
        ):
            out.append(ProcessedRow(RowKind.TITLE, [doc_title]))
            title_written = True
            continue

        if (
            lowered
            == "description item number standard wood premium wood"
            or ("description" in lowered and "item number" in lowered)
        ):
            out.append(
                ProcessedRow(
                    RowKind.TABLE_HEADER,
                    [
                        "Description",
                        "Item Number",
                        "Standard Wood",
                        "Premium Wood",
                    ],
                )
            )
            continue

        if in_index and not INDEX_START_RE.search(joined):
            names = _strip_index_row(row, multiplier)
            if names:
                out.append(ProcessedRow(RowKind.INDEX, names))
            continue

        if INDEX_START_RE.search(joined) or joined.strip() == "B Y M I L L C R A F T™":
            text = _sanitize_text(joined, multiplier)
            if text:
                out.append(ProcessedRow(RowKind.SECTION, [text]))
            continue

        if "collection" in lowered and len([c for c in row if _cell_str(c)]) <= 2:
            text = _sanitize_text(joined, multiplier)
            if text:
                out.append(ProcessedRow(RowKind.SECTION, [text]))
            continue

        if len([c for c in row if _cell_str(c)]) == 1:
            text = _sanitize_text(_cell_str(next(c for c in row if _cell_str(c))), multiplier)
            if not text:
                continue
            if text.isupper() or "BED" in text.upper():
                out.append(ProcessedRow(RowKind.SUBHEAD, [text]))
            else:
                out.append(ProcessedRow(RowKind.SECTION, [text]))
            continue

        cells: list[str] = []
        for j, c in enumerate(row):
            if c is None or _cell_str(c) == "":
                continue
            if j in (3, 4) and isinstance(c, (int, float)) and not isinstance(c, bool):
                cells.append(_format_retail_price(c, multiplier))
                continue
            if isinstance(c, (int, float)) and not isinstance(c, bool):
                continue
            t = _sanitize_text(_cell_str(c), multiplier)
            if t:
                cells.append(t)
        if cells:
            out.append(ProcessedRow(RowKind.PROSE, cells))

    return out, doc_title


def _strip_index_row(row: list, multiplier: float) -> list[str]:
    names: list[str] = []
    for c in row:
        if c is None:
            continue
        if isinstance(c, (int, float)) and not isinstance(c, bool) and c < 300:
            continue
        s = _cell_str(c)
        if not s or s.isdigit():
            continue
        t = _sanitize_text(s, multiplier)
        if t:
            names.append(t)
    return names


def collapse_blank_rows(rows: list[ProcessedRow]) -> list[ProcessedRow]:
    collapsed: list[ProcessedRow] = []
    prev_blank = False
    for pr in rows:
        if pr.kind == RowKind.BLANK:
            if not prev_blank:
                collapsed.append(pr)
            prev_blank = True
        else:
            collapsed.append(pr)
            prev_blank = False
    return collapsed


def forbidden_retail_text(rows: list[ProcessedRow]) -> list[str]:
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
    return bad
