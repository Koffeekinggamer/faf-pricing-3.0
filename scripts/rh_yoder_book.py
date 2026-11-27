"""RH Yoder finished wholesale workbook → retail PDF rows.

Pricing: Excel displays cells as ``$"#,##0`` (half-up whole dollars). Retail is
``ceil(printed × 2.7)`` with Decimal math. Species multipliers are already baked
into each wood column — apply 2.7 once only.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP, ROUND_UP
from pathlib import Path
from typing import Any

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows
from scripts.pricing_lib import assert_no_approx
from scripts.retail_text import (
    NOT_AVAILABLE,
    forbidden_phrases,
    sanitize_retail_text,
)

SHEET_NAME = "Finished Prices From RH Yoder"
DOC_TITLE = "RH Yoder Retail Price List"
MULTIPLIER_DEFAULT = 2.7

# Wood-header tokens (chair / table / server labels)
WOOD_HINT_RE = re.compile(
    r"\b(?:oak|maple|cherry|walnut|hickory|qswo|rustic|elm|wormy|sap)\b",
    re.I,
)
SOLID_RE = re.compile(r"^\s*solid(?:\s*top)?\s*$", re.I)
LEAF_RE = re.compile(r"^\s*\d+\s*leaf\b", re.I)
DASH_RE = re.compile(r"^\s*[-–—]\s*$")
TABLE_RUSTIC_DEDUCT_RE = re.compile(
    r"deduct\s+10%?\s+for\s+rustic\s+species", re.I
)
INCLUDED_RE = re.compile(r"^\s*included\s*$", re.I)
DISCONTINUED_RE = re.compile(r"discontinued", re.I)
SIZE_LABEL_RE = re.compile(r"^\s*size\s*$", re.I)

# Series with no prices — skip entire blocks
DISCONTINUED_SERIES = frozenset(
    {
        "country sheaf",
        "douglas",
        "roosevelt",
        "wellbeck",
        "kirkby table",
    }
)

# Label sits one column left of prices for these table titles (pair by Solid order).
LABEL_SHIFT_TITLES = frozenset(
    {
        "keelan table",
        "korbyn table",
        "korbyn pub table",
        "niles single pedestal table",
    }
)

# Held for vendor confirmation (1-based Excel rows/cols from STRUCTURE.md)
HELD_ALBANY_RUSTIC_ROWS = frozenset(range(35, 42))  # G35:G41
HELD_SINCLAIR_BENCH_ROWS = frozenset({607, 608})  # B607:G608
HELD_TIFTON_LE_ROWS = frozenset({618, 619})  # B618:D619
HELD_NORTH_STAR_QSWO = (1488, 1497, 9, 13)  # I1488:M1497
HELD_IRON_FORGE_WALNUT = (1288, 1291, 10, 10)  # J1288:J1291
HELD_CRAYTON_OAK_1LEAF = (954, 3)  # C954

CHAIR_WOOD_LABELS = [
    "Oak",
    "B. Maple / Wormy Maple / Sap Cherry",
    "Cherry / H. Maple",
    "QSWO / Hickory / White Oak",
    "Walnut",
    "Rustic",
]
CHAIR_SEAT_LABELS = [
    "Std Fabric add",
    "Premium Fabric add",
    "Leather add",
]


@dataclass
class ZoneStats:
    wood_prices: int = 0
    seat_addons: int = 0
    dollar_addons: int = 0
    held_skipped: int = 0
    item_rows: int = 0

    def as_dict(self) -> dict[str, int]:
        return {
            "wood_prices": self.wood_prices,
            "seat_addons": self.seat_addons,
            "dollar_addons": self.dollar_addons,
            "held_skipped": self.held_skipped,
            "item_rows": self.item_rows,
            "priced_cells": self.wood_prices + self.seat_addons + self.dollar_addons,
        }


@dataclass
class RhYoderResult:
    rows: list[ProcessedRow]
    stats: dict[str, Any] = field(default_factory=dict)


def is_rh_yoder_workbook(path: Path) -> bool:
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return False
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return False
    try:
        names = wb.sheetnames
        return SHEET_NAME in names or any(
            "rh yoder" in n.casefold() and "finished" in n.casefold() for n in names
        )
    finally:
        wb.close()


def printed_whole_dollar(value: float | int) -> int:
    """Excel ``$"#,##0`` display: round half-up to the whole dollar."""
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def retail_from_raw(value: float | int, multiplier: float) -> str:
    printed = printed_whole_dollar(value)
    product = Decimal(printed) * Decimal(str(multiplier))
    retail = int(product.to_integral_value(rounding=ROUND_UP))
    return f"${retail:,}"


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    return str(c).strip()


def _is_dash(c: Any) -> bool:
    return isinstance(c, str) and bool(DASH_RE.match(c.strip()))


def _is_included(c: Any) -> bool:
    return isinstance(c, str) and bool(INCLUDED_RE.match(c.strip()))


def _is_number(c: Any) -> bool:
    return isinstance(c, (int, float)) and not isinstance(c, bool)


def _norm_series(name: str) -> str:
    return re.sub(r"\s+", " ", name).strip().casefold()


def _title_key(name: str) -> str:
    # Strip parenthetical height notes for LABEL_SHIFT matching
    t = re.sub(r"\([^)]*\)", "", name)
    t = re.sub(r"\s+", " ", t).strip().casefold()
    for prefix in LABEL_SHIFT_TITLES:
        if t.startswith(prefix) or prefix in t:
            return prefix
    return t


def _is_wood_label(text: str) -> bool:
    if not text or SIZE_LABEL_RE.match(text):
        return False
    if SOLID_RE.match(text) or LEAF_RE.match(text):
        return False
    if text.casefold() in {"add:", "add"}:
        return False
    if re.search(r"deduct\s+\d+\s*%|deduct\s+10", text, re.I):
        return False
    return bool(WOOD_HINT_RE.search(text)) or bool(
        re.search(r"barn\s*wood|reclaimed", text, re.I)
    )


def _append_wide_header(out: list[ProcessedRow], headers: list[str]) -> None:
    """Header row in the WIDE buffer so PDF layout keeps it with price rows."""
    out.append(ProcessedRow(RowKind.WIDE, headers))


def _table_column_headers(
    groups: list[tuple[str, int, int]], size_row: list
) -> list[str]:
    cols = ["Size"]
    for lab, start, end in groups:
        for j in range(start, end + 1):
            leaf = _cell_str(size_row[j]) if j < len(size_row) else ""
            if not leaf or _is_dash(leaf):
                leaf = "Solid"
            cols.append(f"{lab}\n{leaf}")
    return cols


def _drop_empty_wood_group(
    title: str,
    label: str,
    rows: list[list],
    hidden: set[int],
    data_start_i: int,
    data_end_i: int,
    start: int,
    end: int,
) -> bool:
    """Omit wood groups that are entirely held or have no numeric prices."""
    t = title.casefold()
    l = label.casefold()
    if "north star" in t and any(
        x in l for x in ("qswo", "hickory", "white oak")
    ):
        return True
    if "iron forge" in t and "walnut" in l:
        return True
    for i in range(data_start_i, data_end_i):
        r1 = i + 1
        if r1 in hidden:
            continue
        drow = rows[i]
        for col in range(start, end + 1):
            col1 = col + 1
            if _held_cell(r1, col1):
                continue
            c = drow[col] if col < len(drow) else None
            if _is_number(c):
                return False
    return True


def _short_wood_label(text: str) -> str:
    t = re.sub(r"\s+", " ", text).strip()
    t = t.replace("\n", " / ")
    if len(t) > 42:
        t = t[:40].rstrip() + "…"
    return t


def _drop_contact_or_terms(text: str) -> bool:
    low = text.casefold()
    if re.search(r"finished\s+wholesale\s+price\s+list", low):
        return True
    if re.search(r"\bnet\s*30\b|payment\s+terms|finance\s+charge", low):
        return True
    if re.search(r"call\s+or\s+visit|heartland\s+fabrics", low) and re.search(
        r"phone|website|www\.|http|call", low
    ):
        return True
    if re.search(r"vendor\s+(?:address|payment|phone)|removed\s*-\s*not\s+for\s+retail", low):
        return True
    if re.search(
        r"\b(?:phone|fax|email|e-mail)\b|\bwww\.|\.com\b|\d{3}[-.\s]?\d{3}[-.\s]?\d{4}",
        low,
    ) and not re.search(r"seat|fabric|leather|option", low):
        return True
    return False


def _sanitize_note(text: str, multiplier: float, *, allow_dollar_multiply: bool) -> str:
    if not text:
        return ""
    if _drop_contact_or_terms(text):
        return ""
    # Percent / priced-as notes: keep wording, do not invent dollar amounts
    if re.search(r"\d+\s*%|same\s+price\s+as|priced\s+as|deduct\s+10", text, re.I):
        t = sanitize_retail_text(text, multiplier, allow_percent=True)
        # Strip any accidental $ transforms on percent lines — keep as note text
        return t or ""
    if allow_dollar_multiply:
        return sanitize_retail_text(text, multiplier) or ""
    # Notes that may contain $ add-ons (Live edge Add $15) — multiply those
    return sanitize_retail_text(text, multiplier) or ""


def _load_sheet(path: Path) -> tuple[list[list[Any]], set[int]]:
    wb = openpyxl.load_workbook(path, data_only=True)
    if SHEET_NAME in wb.sheetnames:
        ws = wb[SHEET_NAME]
    else:
        ws = wb[[n for n in wb.sheetnames if "finished" in n.casefold()][0]]
    hidden: set[int] = set()
    for r in range(1, ws.max_row + 1):
        if ws.row_dimensions[r].hidden:
            hidden.add(r)
    rows: list[list[Any]] = []
    for r in range(1, ws.max_row + 1):
        row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        rows.append(row)
    wb.close()
    return rows, hidden


def _held_cell(row_1based: int, col_1based: int) -> bool:
    if row_1based in HELD_ALBANY_RUSTIC_ROWS and col_1based == 7:
        return True
    if row_1based in HELD_SINCLAIR_BENCH_ROWS and 2 <= col_1based <= 7:
        return True
    if row_1based in HELD_TIFTON_LE_ROWS and 2 <= col_1based <= 4:
        return True
    r0, r1, c0, c1 = HELD_NORTH_STAR_QSWO
    if r0 <= row_1based <= r1 and c0 <= col_1based <= c1:
        return True
    r0, r1, c0, c1 = HELD_IRON_FORGE_WALNUT
    if r0 <= row_1based <= r1 and c0 <= col_1based <= c1:
        return True
    if (row_1based, col_1based) == HELD_CRAYTON_OAK_1LEAF:
        return True
    return False


def _held_row_entirely(row_1based: int) -> bool:
    return row_1based in HELD_SINCLAIR_BENCH_ROWS or row_1based in HELD_TIFTON_LE_ROWS


def _count_held_wholesale_cells(rows: list[list], hidden: set[int]) -> dict[str, int]:
    """Count priced source cells omitted because they are on the held list."""
    counts = {"chairs": 0, "tables": 0, "servers": 0, "options": 0}
    for i, row in enumerate(rows):
        r1 = i + 1
        if r1 in hidden or r1 in HELD_SINCLAIR_BENCH_ROWS or r1 in HELD_TIFTON_LE_ROWS:
            continue
        if 28 <= r1 <= 717:
            zone = "chairs"
        elif 763 <= r1 <= 2249:
            zone = "tables"
        elif 2279 <= r1 <= 2353:
            zone = "servers"
        elif 718 <= r1 <= 2372:
            zone = "options"
        else:
            continue
        for j, c in enumerate(row):
            if _is_number(c) and _held_cell(r1, j + 1):
                counts[zone] += 1
    return counts


def process_rh_yoder_workbook(
    path: Path, *, multiplier: float = MULTIPLIER_DEFAULT
) -> RhYoderResult:
    rows, hidden = _load_sheet(path)
    out: list[ProcessedRow] = []
    stats = {
        "chairs": ZoneStats(),
        "tables": ZoneStats(),
        "servers": ZoneStats(),
        "options": ZoneStats(),
    }

    out.append(ProcessedRow(RowKind.TITLE, [DOC_TITLE]))
    out.append(ProcessedRow(RowKind.BLANK, []))

    # --- Zone 1: Chairs / barstools / benches ---
    out.append(ProcessedRow(RowKind.SECTION, ["Chairs, Barstools & Benches"]))
    _process_chairs(rows, hidden, out, stats["chairs"], multiplier)

    # --- Zone 1b + 2b + 4 options gathered later; chair options first ---
    out.append(ProcessedRow(RowKind.BLANK, []))
    out.append(ProcessedRow(RowKind.SECTION, ["Tables"]))
    _process_tables(rows, hidden, out, stats["tables"], multiplier)

    out.append(ProcessedRow(RowKind.BLANK, []))
    out.append(ProcessedRow(RowKind.SECTION, ["Servers"]))
    _process_servers(rows, hidden, out, stats["servers"], multiplier)

    out.append(ProcessedRow(RowKind.BLANK, []))
    out.append(ProcessedRow(RowKind.SECTION, ["Options & Finishes"]))
    _process_chair_options(rows, hidden, out, stats["options"], multiplier)
    _process_table_options(rows, hidden, out, stats["options"], multiplier)
    _process_specialty_finishes(rows, hidden, out, stats["options"], multiplier)

    cleaned = collapse_blank_rows(out)
    held_audit = _count_held_wholesale_cells(rows, hidden)
    held_whole_bench = sum(
        1 for r in HELD_SINCLAIR_BENCH_ROWS.union(HELD_TIFTON_LE_ROWS)
        if r not in hidden
        for j, c in enumerate(rows[r - 1])
        if 2 <= j + 1 <= 7 and _is_number(c)
    )
    held_audit["chairs"] += held_whole_bench
    summary = {k: v.as_dict() for k, v in stats.items()}
    for zone, n in held_audit.items():
        summary[zone]["held_skipped"] = n
    summary["totals"] = {
        "wood_prices": sum(v.wood_prices for v in stats.values()),
        "seat_addons": sum(v.seat_addons for v in stats.values()),
        "dollar_addons": sum(v.dollar_addons for v in stats.values()),
        "held_skipped": sum(held_audit.values()),
        "item_rows": sum(v.item_rows for v in stats.values()),
        "priced_cells": sum(
            v.wood_prices + v.seat_addons + v.dollar_addons for v in stats.values()
        ),
    }
    return RhYoderResult(rows=cleaned, stats=summary)


def _process_chairs(
    rows: list[list],
    hidden: set[int],
    out: list[ProcessedRow],
    st: ZoneStats,
    multiplier: float,
) -> None:
    series: str | None = None
    skip_series = False
    for i, row in enumerate(rows):
        r1 = i + 1
        if r1 < 28 or r1 > 717:
            continue
        if r1 in hidden:
            continue
        a = _cell_str(row[0] if row else "")
        if not a:
            continue
        if _drop_contact_or_terms(a):
            continue

        # Series header: name in A, wood labels or bare title with following priced rows
        b = row[1] if len(row) > 1 else None
        is_item_desc = bool(
            re.search(
                r"^(?:side chair|arm chair|\d+\"?\s*(?:swivel|stationary)|"
                r"desk chair|\d+\"\s*x\s*\d+|com needs|live edge on top|"
                r"tacks are|stretcher)",
                a,
                re.I,
            )
        )
        next_has_prices = r1 < len(rows) and any(
            _is_number(rows[r1][j]) if j < len(rows[r1]) else False
            for j in range(1, 7)
        )
        bare_series = (
            not is_item_desc
            and not _is_number(b)
            and not _is_dash(b)
            and not _is_included(b)
            and not any(_is_number(c) for c in row[1:11])
            and not re.match(r"^\d", a)
            and next_has_prices
        )

        # Prefer explicit wood-header series rows
        if isinstance(b, str) and _is_wood_label(b):
            series = a
            skip_series = any(
                d in _norm_series(series) for d in DISCONTINUED_SERIES
            ) or bool(DISCONTINUED_RE.search(series))
            if skip_series:
                continue
            out.append(
                ProcessedRow(
                    RowKind.SUBHEAD,
                    [sanitize_retail_text(series, multiplier) or series],
                )
            )
            _append_wide_header(
                out, ["Item", *CHAIR_WOOD_LABELS, *CHAIR_SEAT_LABELS]
            )
            continue

        # Bare series title (Albany, Tifton Bench, …)
        if bare_series:
            series = a
            skip_series = any(d in _norm_series(series) for d in DISCONTINUED_SERIES)
            if skip_series or DISCONTINUED_RE.search(series):
                continue
            lab = sanitize_retail_text(series, multiplier) or series
            out.append(ProcessedRow(RowKind.SUBHEAD, [lab]))
            _append_wide_header(
                out, ["Item", *CHAIR_WOOD_LABELS, *CHAIR_SEAT_LABELS]
            )
            continue

        if skip_series:
            continue

        # Note-only rows
        if not any(_is_number(c) for c in row[1:11]):
            note = _sanitize_note(a, multiplier, allow_dollar_multiply=True)
            # also mid-row notes
            extras = [
                _sanitize_note(_cell_str(c), multiplier, allow_dollar_multiply=True)
                for c in row[1:]
                if isinstance(c, str) and _cell_str(c)
            ]
            parts = [p for p in [note, *extras] if p]
            if parts:
                out.append(ProcessedRow(RowKind.PROSE, parts[:3]))
            continue

        if _held_row_entirely(r1):
            st.held_skipped += sum(
                1 for j in range(1, 7) if j < len(row) and _is_number(row[j])
            )
            continue

        # Item price row
        cells: list[str] = [sanitize_retail_text(a, multiplier) or a]
        got = False
        for j in range(1, 7):  # B–G woods (0-based 1-6)
            c = row[j] if j < len(row) else None
            col1 = j + 1
            if _held_cell(r1, col1):
                st.held_skipped += 1
                cells.append("")
                continue
            if _is_dash(c) or c is None or c == "":
                cells.append("")
                continue
            if _is_number(c):
                cells.append(retail_from_raw(c, multiplier))
                st.wood_prices += 1
                got = True
            else:
                cells.append("")
        for j in range(7, 10):  # H–J seat add-ons
            c = row[j] if j < len(row) else None
            if _is_included(c):
                cells.append("Included")
                got = True
            elif _is_dash(c) or c is None or c == "":
                cells.append("")
            elif _is_number(c):
                cells.append(retail_from_raw(c, multiplier))
                st.seat_addons += 1
                got = True
            else:
                cells.append("")
        if got:
            # Trim trailing empties but keep alignment for first woods
            while cells and cells[-1] == "":
                cells.pop()
            out.append(ProcessedRow(RowKind.WIDE, cells))
            st.item_rows += 1


def _parse_table_groups(
    title_row: list, size_row: list
) -> list[tuple[str, int, int]]:
    """Return (wood_label, start_col_0based, end_col_0based) from Solid spans."""
    solid_starts = [
        j
        for j, c in enumerate(size_row)
        if isinstance(c, str) and SOLID_RE.match(c.strip())
    ]
    if not solid_starts:
        # Single Solid Top style
        for j, c in enumerate(size_row):
            if isinstance(c, str) and re.search(r"solid", c, re.I):
                solid_starts.append(j)
                break
    labels = [
        (j, _short_wood_label(c))
        for j, c in enumerate(title_row)
        if isinstance(c, str) and _is_wood_label(c)
    ]
    groups: list[tuple[str, int, int]] = []
    if not solid_starts:
        return groups
    for i, start in enumerate(solid_starts):
        end = solid_starts[i + 1] - 1 if i + 1 < len(solid_starts) else max(
            len(size_row), len(title_row)
        ) - 1
        # Find last leaf/solid header within span
        last = start
        for j in range(start, end + 1):
            c = size_row[j] if j < len(size_row) else None
            if isinstance(c, str) and (SOLID_RE.match(c.strip()) or LEAF_RE.match(c.strip())):
                last = j
        label = labels[i][1] if i < len(labels) else f"Wood {i + 1}"
        groups.append((label, start, last))
    return groups


def _peek_table_data_end(
    rows: list[list], hidden: set[int], size_i: int, title: str
) -> int:
    j = size_i + 1
    while j < min(len(rows), 2249):
        r1d = j + 1
        if r1d in hidden:
            j += 1
            continue
        drow = rows[j]
        da = drow[0] if drow else None
        if isinstance(da, str) and da.strip():
            if _is_wood_label(str(drow[1] if len(drow) > 1 else "")) or (
                re.search(r"\btable\b|reclaimed barn wood top", da, re.I)
                and not any(_is_number(c) for c in drow[1:6])
                and not re.match(r"^\d", da.strip())
                and not re.search(r"x", da, re.I)
            ):
                break
            if SIZE_LABEL_RE.match(da.strip()) and any(
                isinstance(c, str) and SOLID_RE.match(c.strip()) for c in drow
            ):
                break
        j += 1
    return j


def _table_block_rustic_deduct_note(drow: list, multiplier: float) -> str:
    mcol = drow[12] if len(drow) > 12 else None
    if not isinstance(mcol, str) or not TABLE_RUSTIC_DEDUCT_RE.search(mcol):
        return ""
    if re.search(r"any\s+rustic", mcol, re.I):
        return ""
    return _sanitize_note(mcol, multiplier, allow_dollar_multiply=False)


def _emit_table_rustic_note(
    out: list[ProcessedRow],
    drow: list,
    multiplier: float,
    *,
    emitted: set[int],
    r1: int,
) -> None:
    if r1 in emitted:
        return
    rustic = _table_block_rustic_deduct_note(drow, multiplier)
    if rustic:
        emitted.add(r1)
        out.append(ProcessedRow(RowKind.PROSE, [rustic]))


def _process_tables(
    rows: list[list],
    hidden: set[int],
    out: list[ProcessedRow],
    st: ZoneStats,
    multiplier: float,
) -> None:
    rustic_note_rows: set[int] = set()
    i = 762  # 0-based → Excel row 763
    while i < min(len(rows), 2249):
        r1 = i + 1
        if r1 in hidden:
            i += 1
            continue
        row = rows[i]
        a = _cell_str(row[0] if row else "")
        if not a:
            i += 1
            continue
        if a.casefold() in {"table pricing", "buffet/server pricing"}:
            i += 1
            continue
        if _drop_contact_or_terms(a):
            i += 1
            continue

        # Detect table title: A text + wood labels in row, or known table name
        wood_labs = [
            j for j, c in enumerate(row) if isinstance(c, str) and _is_wood_label(c)
        ]
        is_title = bool(wood_labs) or (
            re.search(r"\btable\b|\bbench\b|reclaimed barn wood top", a, re.I)
            and not any(_is_number(c) for c in row[1:])
            and not SIZE_LABEL_RE.match(a)
        )
        if not is_title:
            # Footnote / deduct note
            if re.search(r"deduct\s+10|same\s+price|add\s+\d+%", a, re.I) or (
                isinstance(row[12] if len(row) > 12 else None, str)
                and re.search(r"deduct\s+10", str(row[12]), re.I)
            ):
                note = _sanitize_note(a, multiplier, allow_dollar_multiply=False)
                for c in row[1:]:
                    if isinstance(c, str) and re.search(r"deduct|same price|add \d+%", c, re.I):
                        n2 = _sanitize_note(c, multiplier, allow_dollar_multiply=False)
                        if n2:
                            note = f"{note}; {n2}" if note else n2
                if note:
                    out.append(ProcessedRow(RowKind.PROSE, [note]))
            i += 1
            continue

        title = a
        if any(d in _norm_series(title) for d in DISCONTINUED_SERIES) or DISCONTINUED_RE.search(
            title
        ):
            # Skip until next title
            i += 1
            while i < 2249:
                r1b = i + 1
                if r1b in hidden:
                    i += 1
                    continue
                aa = _cell_str(rows[i][0] if rows[i] else "")
                labs = [
                    j
                    for j, c in enumerate(rows[i])
                    if isinstance(c, str) and _is_wood_label(c)
                ]
                if labs and aa and aa.casefold() != title.casefold():
                    break
                if (
                    aa
                    and re.search(r"\btable\b", aa, re.I)
                    and aa.casefold() != title.casefold()
                    and not any(_is_number(c) for c in rows[i][1:])
                ):
                    break
                i += 1
            continue

        # Size header row
        size_i = i + 1
        while size_i < len(rows) and (size_i + 1) in hidden:
            size_i += 1
        if size_i >= len(rows):
            i += 1
            continue
        size_row = rows[size_i]
        size_a = _cell_str(size_row[0] if size_row else "")
        if not SIZE_LABEL_RE.match(size_a) and not any(
            isinstance(c, str) and SOLID_RE.match(c.strip()) for c in size_row
        ):
            # Reclaimed-style: title then Size
            if not wood_labs and "reclaimed" in title.casefold():
                groups = [("Reclaimed Barn Wood", 1, 1)]
            else:
                i += 1
                continue
        else:
            groups = _parse_table_groups(row, size_row)
            if not groups and "reclaimed" in title.casefold():
                groups = [("Reclaimed Barn Wood", 1, 1)]

        if not groups:
            i += 1
            continue

        data_end = _peek_table_data_end(rows, hidden, size_i, title)
        groups = [
            g
            for g in groups
            if not _drop_empty_wood_group(
                title, g[0], rows, hidden, size_i + 1, data_end, g[1], g[2]
            )
        ]
        if not groups:
            for ti in range(size_i + 1, data_end):
                if ti + 1 in hidden:
                    continue
                _emit_table_rustic_note(
                    out, rows[ti], multiplier, emitted=rustic_note_rows, r1=ti + 1
                )
            i = data_end
            continue

        out.append(
            ProcessedRow(
                RowKind.SUBHEAD,
                [sanitize_retail_text(title, multiplier) or title],
            )
        )
        _append_wide_header(out, _table_column_headers(groups, size_row))

        # Data rows
        j = size_i + 1
        while j < min(len(rows), 2249):
            r1d = j + 1
            if r1d in hidden:
                j += 1
                continue
            drow = rows[j]
            da = drow[0] if drow else None
            # Next title?
            if isinstance(da, str) and da.strip():
                if _is_wood_label(str(drow[1] if len(drow) > 1 else "")) or (
                    re.search(r"\btable\b|reclaimed barn wood top", da, re.I)
                    and not any(_is_number(c) for c in drow[1:6])
                    and not re.match(r"^\d", da.strip())
                    and not re.search(r"x", da, re.I)
                ):
                    break
                if SIZE_LABEL_RE.match(da.strip()) and any(
                    isinstance(c, str) and SOLID_RE.match(c.strip()) for c in drow
                ):
                    # Continuation block same series with new woods — handle as new title
                    break
            _emit_table_rustic_note(
                out, drow, multiplier, emitted=rustic_note_rows, r1=r1d
            )
            if (
                r1d in rustic_note_rows
                and not any(_is_number(c) for c in drow[1:])
                and not (_cell_str(da) and _cell_str(da).strip())
            ):
                j += 1
                continue

            # Footnotes
            if isinstance(da, str) and da.strip() and not any(
                _is_number(c) for c in drow[1:]
            ) and not re.match(r"^\d", da.strip()) and "x" not in da.casefold():
                # deduct note in far column
                note = _sanitize_note(da, multiplier, allow_dollar_multiply=True)
                for c in drow[1:]:
                    if isinstance(c, str) and re.search(
                        r"deduct|same price|add \d+%|reclaimed", c, re.I
                    ):
                        n2 = _sanitize_note(c, multiplier, allow_dollar_multiply=False)
                        if n2:
                            note = f"{note}; {n2}" if note else n2
                if note:
                    out.append(ProcessedRow(RowKind.PROSE, [note]))
                _emit_table_rustic_note(
                    out, drow, multiplier, emitted=rustic_note_rows, r1=r1d
                )
                j += 1
                # If this looks like end-of-block prose, keep going until blank/title
                continue

            size_label = _cell_str(da) if da is not None else ""
            if not size_label and not any(_is_number(c) for c in drow[1:]):
                _emit_table_rustic_note(
                    out, drow, multiplier, emitted=rustic_note_rows, r1=r1d
                )
                j += 1
                continue
            if isinstance(da, (int, float)) and not isinstance(da, bool):
                size_label = str(int(da)) if float(da) == int(da) else str(da)

            cells = [sanitize_retail_text(size_label, multiplier) or size_label]
            # For multi-leaf groups: emit one retail per wood using Solid column only
            # in the header row display — actually emit all leaf prices as separate
            # wide values under each wood? Spot check wants woods per row.
            # Structure: each wood group has Solid, 1 Leaf, 2 Leaf… — show Solid price
            # as the wood column and put leaf prices as additional $ in the row, OR
            # show one row per size with one $ per wood (Solid) and note leaf cols.
            #
            # User: "list by series name, description and size only" + woods per row.
            # Emit: Size | oak_solid | oak_1leaf | … gets huge. Better: one $ per wood
            # group using each leaf column as parallel woods labeled in header.
            #
            # Practical approach matching Viztech tables: for each wood group, take
            # ALL leaf columns as sequential prices under that wood's span, flattened
            # into the row with header listing wood+leaf.
            #
            # Simpler approach used here: one cell per wood group = Solid price when
            # present, else first numeric; additional leaf prices appended with labels
            # in a prose trailer. Actually STRUCTURE counts 6,098 table wood prices —
            # every leaf cell is a priced cell. So emit every leaf cell as a $ token.
            #
            # Row format: [size, $… for each priced leaf cell in group order]
            # Header lists wood labels once; leaf config in prose above.
            for lab, start, end in groups:
                for col in range(start, end + 1):
                    c = drow[col] if col < len(drow) else None
                    col1 = col + 1
                    if _held_cell(r1d, col1):
                        st.held_skipped += 1
                        cells.append("")
                        continue
                    if _is_dash(c) or c is None or c == "":
                        cells.append("")
                        continue
                    if _is_number(c):
                        cells.append(retail_from_raw(c, multiplier))
                        st.wood_prices += 1
                    else:
                        cells.append("")
            if any(c.startswith("$") for c in cells[1:] if c):
                while cells and cells[-1] == "":
                    cells.pop()
                out.append(ProcessedRow(RowKind.WIDE, cells))
                st.item_rows += 1
            j += 1

        trail = j
        while trail < data_end:
            r1t = trail + 1
            if r1t in hidden:
                trail += 1
                continue
            _emit_table_rustic_note(
                out,
                rows[trail],
                multiplier,
                emitted=rustic_note_rows,
                r1=trail + 1,
            )
            trail += 1

        i = j

    for r1 in range(764, 2250):
        if r1 in hidden:
            continue
        _emit_table_rustic_note(
            out, rows[r1 - 1], multiplier, emitted=rustic_note_rows, r1=r1
        )


def _process_servers(
    rows: list[list],
    hidden: set[int],
    out: list[ProcessedRow],
    st: ZoneStats,
    multiplier: float,
) -> None:
    i = 2278  # Excel 2279
    series: str | None = None
    wood_cols: list[tuple[str, int]] = []
    while i < min(len(rows), 2353):
        r1 = i + 1
        if r1 in hidden:
            i += 1
            continue
        row = rows[i]
        a = _cell_str(row[0] if row else "")
        if a.casefold() == "buffet/server pricing":
            i += 1
            continue
        if _drop_contact_or_terms(a):
            i += 1
            continue
        if SIZE_LABEL_RE.match(a):
            i += 1
            continue
        deduct_only = [
            c
            for c in row[1:12]
            if isinstance(c, str) and re.search(r"deduct\s+10", c, re.I)
        ]
        if deduct_only and not any(_is_number(c) for c in row[1:12]):
            for c in deduct_only:
                note = _sanitize_note(c, multiplier, allow_dollar_multiply=False)
                if note:
                    out.append(ProcessedRow(RowKind.PROSE, [note]))
            if a and not SIZE_LABEL_RE.match(a):
                extra = _sanitize_note(a, multiplier, allow_dollar_multiply=True)
                if extra:
                    out.append(ProcessedRow(RowKind.PROSE, [extra]))
            i += 1
            continue
        wood_labs = [
            (j, _short_wood_label(c))
            for j, c in enumerate(row)
            if isinstance(c, str) and _is_wood_label(c)
        ]
        if a and wood_labs and not any(_is_number(c) for c in row[1:10]):
            series = a
            wood_cols = [(lab, j) for j, lab in wood_labs]
            barn_j = next(
                (
                    j
                    for j, c in enumerate(row)
                    if isinstance(c, str) and re.search(r"barn\s*wood", c, re.I)
                ),
                None,
            )
            if barn_j is not None:
                wood_cols = [(lab, colj) for lab, colj in wood_cols if colj != barn_j]
                wood_cols.append(("Barn Wood Top add", barn_j))
            out.append(
                ProcessedRow(
                    RowKind.SUBHEAD,
                    [sanitize_retail_text(series, multiplier) or series],
                )
            )
            _append_wide_header(out, ["Size", *[w[0] for w in wood_cols]])
            i += 1
            continue
        if a and not any(_is_number(c) for c in row[1:12]):
            note = _sanitize_note(a, multiplier, allow_dollar_multiply=True)
            if note and not SIZE_LABEL_RE.match(a):
                out.append(ProcessedRow(RowKind.PROSE, [note]))
            i += 1
            continue
        if a and wood_cols and any(_is_number(c) for c in row[1:12]):
            cells = [sanitize_retail_text(a, multiplier) or a]
            for lab, j in wood_cols:
                c = row[j] if j < len(row) else None
                col1 = j + 1
                if _held_cell(r1, col1):
                    st.held_skipped += 1
                    cells.append("")
                    continue
                if _is_dash(c) or c is None or c == "":
                    cells.append("")
                elif _is_number(c):
                    cells.append(retail_from_raw(c, multiplier))
                    if "add" in lab.casefold():
                        st.dollar_addons += 1
                    else:
                        st.wood_prices += 1
                else:
                    cells.append("")
            while cells and cells[-1] == "":
                cells.pop()
            out.append(ProcessedRow(RowKind.WIDE, cells))
            st.item_rows += 1
        i += 1


def _process_chair_options(
    rows: list[list],
    hidden: set[int],
    out: list[ProcessedRow],
    st: ZoneStats,
    multiplier: float,
) -> None:
    out.append(ProcessedRow(RowKind.SUBHEAD, ["Chair options"]))
    for i in range(717, min(len(rows), 738)):
        r1 = i + 1
        if r1 in hidden:
            continue
        row = rows[i]
        a = _cell_str(row[0] if row else "")
        if not a or a.casefold().startswith("special option"):
            continue
        if _drop_contact_or_terms(a):
            continue
        k = row[10] if len(row) > 10 else None  # column K
        if _is_number(k):
            retail = retail_from_raw(k, multiplier)
            st.dollar_addons += 1
            prose = f"{sanitize_retail_text(a, multiplier) or a}: {retail}"
            out.append(ProcessedRow(RowKind.PROSE, [prose]))
        elif isinstance(k, str) and re.search(r"per\s+set|\$\s*\d+", k, re.I):
            # "$40 per set" style in A or K
            text = f"{a} {k}".strip()
            note = _sanitize_note(text, multiplier, allow_dollar_multiply=True)
            # Rewrite "additional $40 per set" → retail
            if note:
                out.append(ProcessedRow(RowKind.PROSE, [note]))
                if "$" in note:
                    st.dollar_addons += 1
        elif isinstance(k, str) and re.search(r"deduct|same price", k, re.I):
            note = _sanitize_note(f"{a}: {k}", multiplier, allow_dollar_multiply=False)
            if note:
                out.append(ProcessedRow(RowKind.PROSE, [note]))
        else:
            # Embedded $40 in A732
            if "$" in a:
                note = _sanitize_note(a, multiplier, allow_dollar_multiply=True)
            else:
                note = _sanitize_note(a, multiplier, allow_dollar_multiply=False)
            if note:
                out.append(ProcessedRow(RowKind.PROSE, [note]))


def _process_table_options(
    rows: list[list],
    hidden: set[int],
    out: list[ProcessedRow],
    st: ZoneStats,
    multiplier: float,
) -> None:
    out.append(ProcessedRow(RowKind.SUBHEAD, ["Table options"]))
    for i in range(2249, min(len(rows), 2278)):
        r1 = i + 1
        if r1 in hidden:
            continue
        row = rows[i]
        a = _cell_str(row[0] if row else "")
        if not a:
            continue
        if a.casefold() in {"available edge options on tables", "table upgrade costs"}:
            out.append(ProcessedRow(RowKind.PROSE, [a]))
            continue
        k = row[10] if len(row) > 10 else None
        if _is_number(k):
            retail = retail_from_raw(k, multiplier)
            st.dollar_addons += 1
            label = sanitize_retail_text(a, multiplier) or a
            if re.search(r"per\s*leaf|cost/leaf", a, re.I):
                out.append(ProcessedRow(RowKind.PROSE, [f"{label}: {retail} per leaf"]))
            else:
                out.append(ProcessedRow(RowKind.PROSE, [f"{label}: {retail}"]))
        elif isinstance(k, str) and re.search(r"\$\s*\d+", k):
            # "$45 per leaf"
            printed_m = re.search(r"\$\s*([\d,]+(?:\.\d+)?)", k)
            if printed_m and re.search(r"per\s+leaf", k, re.I):
                raw = float(printed_m.group(1).replace(",", ""))
                retail = retail_from_raw(raw, multiplier)
                st.dollar_addons += 1
                label = sanitize_retail_text(a, multiplier) or a
                out.append(ProcessedRow(RowKind.PROSE, [f"{label}: {retail} per leaf"]))
            else:
                note = _sanitize_note(f"{a}: {k}", multiplier, allow_dollar_multiply=True)
                if note:
                    out.append(ProcessedRow(RowKind.PROSE, [note]))
        elif isinstance(k, str) and re.search(
            r"%|same price|no charge|add\s+\d+", k, re.I
        ):
            label = re.sub(r"\s*\(unfinished\s+price\)\s*", " ", a, flags=re.I)
            label = re.sub(r"\s+", " ", label).strip()
            note = _sanitize_note(f"{label}: {k}", multiplier, allow_dollar_multiply=False)
            if note:
                out.append(ProcessedRow(RowKind.PROSE, [note]))
        else:
            label = re.sub(r"\s*\(unfinished\s+price\)\s*", " ", a, flags=re.I)
            note = _sanitize_note(label, multiplier, allow_dollar_multiply=False)
            if note:
                out.append(ProcessedRow(RowKind.PROSE, [note]))


def _process_specialty_finishes(
    rows: list[list],
    hidden: set[int],
    out: list[ProcessedRow],
    st: ZoneStats,
    multiplier: float,
) -> None:
    out.append(ProcessedRow(RowKind.SUBHEAD, ["Specialty finishes"]))
    # Header labels from row 2362/2363
    finish_cols = [
        (1, "2-Tone"),
        (2, "Polyurethane"),
        (4, "Cat. 2 / Paint / Wire Brush / Distress"),
        (6, "Cat. 3 color"),
    ]
    _append_wide_header(out, ["Item", *[lab for _, lab in finish_cols]])
    for i in range(2363, min(len(rows), 2372)):
        r1 = i + 1
        if r1 in hidden:
            continue
        row = rows[i]
        a = _cell_str(row[0] if row else "")
        if not a:
            continue
        if re.search(r"sheen\s+top\s+coat", a, re.I):
            out.append(
                ProcessedRow(
                    RowKind.PROSE,
                    [sanitize_retail_text(a, multiplier) or a],
                )
            )
            continue
        cells = [sanitize_retail_text(a, multiplier) or a]
        got = False
        for j, _lab in finish_cols:
            c = row[j] if j < len(row) else None
            if _is_dash(c) or c is None or c == "":
                cells.append("")
            elif _is_number(c):
                cells.append(retail_from_raw(c, multiplier))
                st.dollar_addons += 1
                got = True
            else:
                cells.append("")
        if got:
            while cells and cells[-1] == "":
                cells.pop()
            out.append(ProcessedRow(RowKind.WIDE, cells))
            st.item_rows += 1


def rh_yoder_forbidden_text(processed: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for row in processed:
        for c in row.cells:
            s = str(c)
            low = s.casefold()
            if "wholesale" in low or "dealer" in low or re.search(r"\bmap\b", low):
                bad.append(s)
            if re.search(r"call\s+for|call\s+or\s+visit", low):
                bad.append(s)
            bad.extend(forbidden_phrases(s))
            if re.search(
                r"\d{3}[-.\s]?\d{3}[-.\s]?\d{4}|@|\bwww\.|password", low
            ):
                bad.append(s)
    # unique preserve order
    seen: set[str] = set()
    out: list[str] = []
    for b in bad:
        if b not in seen:
            seen.add(b)
            out.append(b)
    return out


def held_items_for_pr() -> list[str]:
    return [
        "Albany chairs Rustic column (G35–G41)",
        'Sinclair 96" and 108" benches (B607–G608)',
        'Tifton Bench with Live Edge 96" and 108" (B618–D619)',
        "North Star Table QSWO/Hickory/White Oak column (I1488–M1497)",
        "Iron Forge Live Edge Walnut (J1288–J1291)",
        'Crayton Table Oak 42x60 1-leaf (C954)',
    ]
