"""Country Comfort wholesale workbooks → finished-only retail PDF rows (×2.7)."""

from __future__ import annotations

import math
import re
from pathlib import Path
from typing import Any

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows
from scripts.pricing_lib import retail_whole_dollar
from scripts.retail_text import forbidden_phrases, sanitize_retail_text

BUILDER = "Country Comfort"
MULTIPLIER = 2.7
DOC_TITLE = "Country Comfort Retail Price List"

SKIP_MAIN = frozenset({"MARKUP", "MASTER", "old Fairmont"})
SKIP_LB = frozenset({"Markup", "Sheet2"})

ORDER_RE = re.compile(r"^order\s*#", re.I)
FIN_RE = re.compile(r"^fin\.?$|^finished$", re.I)
UNFIN_RE = re.compile(r"^unfin\.?$|^unfinished$", re.I)
ORDER_TOKEN_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-]*$")
SIZE_RE = re.compile(r"^\d+\s*[xX×]\s*\d+")
FLAT_ADD_RE = re.compile(
    r"(fabric\s*seat|leather|adena\s*cherry\s*seat|cherry\s*seat|elm\s*or\s*qswo|"
    r"walnut\s*seat|wormy\s*maple|footrest|foot\s*protector|live\s*edge|"
    r"reclaimed\s*seat|com\s*fabric|two\s*tone|paint|performance)",
    re.I,
)
PCT_RE = re.compile(r"\d+\s*%")


def _format_flat_add(label: str, amt: float, multiplier: float) -> tuple[str, str]:
    """Return (label, price cell). Every flat-dollar add uses retail multiplier."""
    text = sanitize_retail_text(_fix_typo(label), multiplier, allow_percent=True)
    if not text:
        return "", ""
    retail = retail_whole_dollar(abs(float(amt)), multiplier)
    if "subtract" in text.lower():
        return text, f"−${retail:,}"
    return text, f"+${retail:,}"


def _cell(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _is_num(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _retail(v: float, multiplier: float = MULTIPLIER) -> str:
    return f"${retail_whole_dollar(float(v), multiplier):,}"


def _fix_typo(text: str) -> str:
    return re.sub(r"\bPreformance\b", "Performance", text, flags=re.I)


def _sheet_rows(ws) -> list[list[Any]]:
    return [list(row) for row in ws.iter_rows(values_only=True)]


def _find_order_header(rows: list[list[Any]]) -> tuple[int, list[int]] | None:
    for i, row in enumerate(rows):
        idxs = [j for j, c in enumerate(row) if isinstance(c, str) and ORDER_RE.match(c.strip())]
        if idxs:
            return i, idxs
    return None


def _left_block_end(row: list[Any], first_order_col: int, second_order_col: int | None) -> int:
    if second_order_col is not None:
        return second_order_col
    return len(row)


def _fin_cols(header: list[Any], start: int, end: int) -> list[int]:
    cols: list[int] = []
    for j in range(start, min(end, len(header))):
        lab = _cell(header[j])
        if FIN_RE.match(lab):
            cols.append(j)
    if cols:
        return cols
    # Bayfield-style: Fin. labels only; take numeric-ready cols after DESCRIPTION
    # Fallback: every even offset after description if Unfin./Fin. pairs without labels on Bayfield left
    return cols


def _wood_label_for_fin_col(rows: list[list[Any]], header_i: int, fin_col: int) -> str:
    """Gather wood names from rows above the ORDER # header for this Fin. column."""
    parts: list[str] = []
    # Prefer label on Unfin column (fin_col-1) then fin_col itself
    for probe in (fin_col - 1, fin_col, fin_col - 2):
        if probe < 0:
            continue
        for r in range(max(0, header_i - 6), header_i):
            if probe >= len(rows[r]):
                continue
            t = _cell(rows[r][probe])
            if not t or UNFIN_RE.match(t) or FIN_RE.match(t):
                continue
            if ORDER_RE.match(t) or t.upper() in {"ORDER #", "DESCRIPTION"}:
                continue
            if PCT_RE.fullmatch(t.replace(" ", "")):
                continue
            if _is_num(rows[r][probe]):
                continue
            if t not in parts:
                parts.append(t)
        if parts:
            break
    if not parts:
        return f"Wood {fin_col}"
    # Prefer a short readable join
    return " / ".join(parts[:4])


def _is_order_id(v: Any) -> bool:
    if v is None:
        return False
    if _is_num(v) and not isinstance(v, bool):
        return True
    s = _cell(v)
    if not s or not ORDER_TOKEN_RE.match(s):
        return False
    if s.lower() in {"order", "description", "unfin.", "fin.", "unfinished", "finished"}:
        return False
    # order tokens usually contain a digit
    return any(ch.isdigit() for ch in s)


def _parse_chair_block(
    rows: list[list[Any]],
    *,
    section_name: str,
    multiplier: float,
    header_i: int,
    order_col: int,
    block_end: int,
) -> tuple[list[ProcessedRow], int]:
    header = rows[header_i]
    desc_col = order_col + 1
    fin_cols = _fin_cols(header, order_col + 2, block_end)
    if not fin_cols:
        # Fin-only sheets (Bayfield left): columns labeled Fin. starting after DESCRIPTION
        for j in range(order_col + 2, block_end):
            if j < len(header) and FIN_RE.match(_cell(header[j])):
                fin_cols.append(j)
    if not fin_cols:
        return [], 0

    wood_headers = [_wood_label_for_fin_col(rows, header_i, c) for c in fin_cols]
    out: list[ProcessedRow] = [ProcessedRow(RowKind.SECTION, [section_name])]
    out.append(
        ProcessedRow(
            RowKind.TABLE_HEADER,
            ["Item #", "Description", *wood_headers],
        )
    )
    count = 0
    for r in range(header_i + 1, len(rows)):
        row = rows[r]
        oid = row[order_col] if order_col < len(row) else None
        if not _is_order_id(oid):
            # option / note rows under the table
            label = _cell(oid) if oid is not None else ""
            if not label:
                # sometimes label only in col0 while parsing a mid-sheet block
                label = _cell(row[0]) if row else ""
            joined = " ".join(_cell(c) for c in row[:block_end] if _cell(c))
            if not joined:
                continue
            if re.search(r"not available", joined, re.I):
                note = sanitize_retail_text(_fix_typo(joined), multiplier)
                if note:
                    out.append(ProcessedRow(RowKind.PROSE, [note]))
                continue
            if FLAT_ADD_RE.search(joined) or "add" in joined.lower() or "subtract" in joined.lower():
                amt = None
                for c in row:
                    if _is_num(c) and float(c) != 0:
                        amt = float(c)
                        break
                text = _fix_typo(label or joined)
                if amt is not None and not PCT_RE.search(text):
                    lab, price = _format_flat_add(text, amt, multiplier)
                    if lab:
                        out.append(ProcessedRow(RowKind.WIDE, [lab, price]))
                else:
                    text = sanitize_retail_text(text, multiplier, allow_percent=True)
                    if text:
                        out.append(ProcessedRow(RowKind.PROSE, [text]))
                continue
            continue

        desc = _cell(row[desc_col]) if desc_col < len(row) else ""
        desc = _fix_typo(desc)
        prices: list[str] = []
        any_price = False
        for fc in fin_cols:
            v = row[fc] if fc < len(row) else None
            if _is_num(v) and float(v) > 0:
                prices.append(_retail(float(v), multiplier))
                any_price = True
            else:
                prices.append("—")
        if not any_price:
            continue
        item = _cell(oid)
        out.append(ProcessedRow(RowKind.WIDE, [item, desc, *prices]))
        count += 1
    return out, count


def _parse_standard_collection(
    rows: list[list[Any]], sheet_name: str, multiplier: float
) -> tuple[list[ProcessedRow], int]:
    found = _find_order_header(rows)
    if not found:
        return [], 0
    header_i, order_idxs = found
    order_col = order_idxs[0]
    second = order_idxs[1] if len(order_idxs) > 1 else None
    end = _left_block_end(rows[header_i], order_col, second)
    return _parse_chair_block(
        rows,
        section_name=sheet_name,
        multiplier=multiplier,
        header_i=header_i,
        order_col=order_col,
        block_end=end,
    )


def _parse_loraine(rows: list[list[Any]], multiplier: float) -> tuple[list[ProcessedRow], int]:
    out = [ProcessedRow(RowKind.SECTION, ["Loraine"])]
    out.append(
        ProcessedRow(RowKind.TABLE_HEADER, ["Item #", "Description", "Price"])
    )
    count = 0
    for row in rows:
        if not _is_order_id(row[0] if row else None):
            continue
        desc = _fix_typo(_cell(row[1] if len(row) > 1 else ""))
        price = None
        for c in row[2:]:
            if _is_num(c) and float(c) > 0:
                price = float(c)
                break
        if price is None:
            continue
        out.append(
            ProcessedRow(
                RowKind.WIDE,
                [_cell(row[0]), desc, _retail(price, multiplier)],
            )
        )
        count += 1
    note = "Black or Sungold Metal"
    for row in rows[:5]:
        j = " ".join(_cell(c) for c in row if _cell(c))
        if "metal" in j.lower():
            note = j
            break
    out.insert(1, ProcessedRow(RowKind.PROSE, [note]))
    return out, count


def _parse_ruff_sawn(rows: list[list[Any]], multiplier: float) -> tuple[list[ProcessedRow], int]:
    """Ruff Sawn sheet has Christy and Delaney subsections."""
    out: list[ProcessedRow] = []
    total = 0
    # Find subsection titles and ORDER # headers
    sections: list[tuple[str, int]] = []
    for i, row in enumerate(rows):
        c0 = _cell(row[0] if row else "")
        c1 = _cell(row[1] if len(row) > 1 else "")
        if c1 and not c0 and c1.lower() in {"christy", "delaney"}:
            sections.append((f"Ruff Sawn — {c1}", i))
        elif c0.lower() in {"christy", "delaney"}:
            sections.append((f"Ruff Sawn — {c0}", i))
    if not sections:
        return _parse_standard_collection(rows, "Ruff Sawn", multiplier)

    for si, (name, start) in enumerate(sections):
        end = sections[si + 1][1] if si + 1 < len(sections) else len(rows)
        block = rows[start:end]
        # re-index header search within block
        found = _find_order_header(block)
        if not found:
            continue
        hi, idxs = found
        chunk, n = _parse_chair_block(
            block,
            section_name=name,
            multiplier=multiplier,
            header_i=hi,
            order_col=idxs[0],
            block_end=len(block[hi]),
        )
        out.extend(chunk)
        total += n
    return out, total


def _table_wood_header(cell: Any) -> str:
    t = _cell(cell).replace("\n", ", ")
    t = re.sub(r"\s+", " ", t).strip(" ,")
    return t or "Finished"


def _parse_table_sheet(
    rows: list[list[Any]], sheet_name: str, multiplier: float
) -> tuple[list[ProcessedRow], int]:
    """Parse multi-block table size sheets into one or more collection sections."""
    out: list[ProcessedRow] = []
    total = 0
    # Detect titled blocks from row 0 / later title cells
    # Strategy: walk rows; when we see a size like 36x36 with a price column, emit under current titles.

    # Collect column groups: each group has title cell, size col, leaf desc col, overall col, price col(s)
    # Simpler approach: for each pair of left/right panels, parse independently.

    titles_row = rows[0] if rows else []
    # Find title columns (non-empty text in row 1 that aren't "Overall Size")
    title_cols: list[tuple[int, str]] = []
    for j, c in enumerate(titles_row):
        t = _cell(c)
        if not t:
            continue
        if "overall size" in t.lower():
            continue
        if "description" in t.lower():
            continue
        title_cols.append((j, re.sub(r"\s+", " ", t.replace("\n", " ")).strip()))

    # Also pick up mid-sheet titles (e.g. Chelsea double pedestal)
    for i, row in enumerate(rows[1:20], start=1):
        for j, c in enumerate(row):
            t = _cell(c)
            if not t or len(t) < 8:
                continue
            if "pedestal" in t.lower() or "trestle" in t.lower() or "table" in t.lower():
                if "description" in t.lower() or "overall" in t.lower():
                    continue
                label = re.sub(r"\s+", " ", t.replace("\n", " ")).strip()
                if all(label != x[1] for x in title_cols):
                    title_cols.append((j, label))

    # Parse each titled column as a vertical block: size at title_col, desc at +1, overall +2, price +3 (+ optional +4)
    seen_blocks: set[str] = set()
    for title_col, title in title_cols:
        key = f"{title_col}:{title}"
        if key in seen_blocks:
            continue
        seen_blocks.add(key)
        # Find wood header near title_col+3
        wood = "Finished"
        for r in range(0, min(12, len(rows))):
            for probe in (title_col + 3, title_col + 4, title_col + 2):
                if probe < len(rows[r]) and _cell(rows[r][probe]) and not _is_num(rows[r][probe]):
                    cand = _table_wood_header(rows[r][probe])
                    if "overall" in cand.lower() or "description" in cand.lower():
                        continue
                    if len(cand) > 3:
                        wood = cand
                        break
            else:
                continue
            break

        # Optional second wood package column
        wood2 = None
        wood2_col = title_col + 4
        for r in range(0, min(12, len(rows))):
            if wood2_col < len(rows[r]) and _cell(rows[r][wood2_col]) and not _is_num(rows[r][wood2_col]):
                cand = _table_wood_header(rows[r][wood2_col])
                if cand and cand != wood and "overall" not in cand.lower() and len(cand) > 8:
                    wood2 = cand
                    break

        headers = ["Size", "Leaf / Top", "Overall", wood]
        if wood2:
            headers.append(wood2)
        section_rows: list[ProcessedRow] = [
            ProcessedRow(RowKind.SECTION, [title]),
            ProcessedRow(RowKind.TABLE_HEADER, headers),
        ]
        n = 0
        for row in rows:
            size = _cell(row[title_col]) if title_col < len(row) else ""
            if not SIZE_RE.match(size.replace(" ", "")):
                # sometimes size is already without spaces
                if not re.match(r"^\d+x\d+", size.replace(" ", ""), re.I):
                    continue
            leaf = _cell(row[title_col + 1]) if title_col + 1 < len(row) else ""
            overall = row[title_col + 2] if title_col + 2 < len(row) else None
            price = row[title_col + 3] if title_col + 3 < len(row) else None
            price2 = row[title_col + 4] if wood2 and title_col + 4 < len(row) else None
            if not _is_num(price) or float(price) <= 0:
                continue
            overall_s = str(int(overall)) if _is_num(overall) else _cell(overall)
            cells = [
                size.replace(" ", "").lower() if "x" in size.lower() else size,
                leaf,
                overall_s,
                _retail(float(price), multiplier),
            ]
            if wood2:
                if _is_num(price2) and float(price2) > 0:
                    cells.append(_retail(float(price2), multiplier))
                else:
                    cells.append("—")
            section_rows.append(ProcessedRow(RowKind.WIDE, cells))
            n += 1
        if n:
            out.extend(section_rows)
            total += n
    return out, total


def _parse_options(rows: list[list[Any]], title: str) -> list[ProcessedRow]:
    out = [ProcessedRow(RowKind.SECTION, [title])]
    for row in rows:
        label = _fix_typo(_cell(row[0] if row else ""))
        if not label or label.upper() == "OPTIONS":
            continue
        if label.lower().startswith("free options") or label.lower() == "standard hardwoods":
            out.append(ProcessedRow(RowKind.SUBHEAD, [label]))
            continue
        if label.lower().startswith("add ") and PCT_RE.search(label):
            out.append(ProcessedRow(RowKind.SUBHEAD, [label]))
            continue
        if label.lower().startswith("additional table"):
            out.append(ProcessedRow(RowKind.SUBHEAD, [label]))
            continue
        # flat dollar in col 3 or embedded in label
        amt = None
        for c in row[1:]:
            if _is_num(c):
                amt = float(c)
                break
        m = re.search(r"\+\s*\$?\s*([\d,]+)", label)
        pct = re.search(r"\+\s*(\d+)\s*%", label)
        if pct:
            out.append(ProcessedRow(RowKind.WIDE, [re.sub(r"\s+", " ", label.replace("\n", " ")), f"+{pct.group(1)}%"]))
            continue
        if amt is not None:
            lab, price = _format_flat_add(label.strip(), amt, MULTIPLIER)
            if lab:
                out.append(ProcessedRow(RowKind.WIDE, [lab, price]))
            continue
        if "=" in label:
            # e.g. Distressed = Free or + $50 (flat dollars → × multiplier)
            clean = re.sub(r"\s+", " ", label.replace("\n", " ")).strip()
            clean = sanitize_retail_text(clean, MULTIPLIER, allow_percent=True) or clean
            out.append(ProcessedRow(RowKind.PROSE, [clean]))
            continue
        if PCT_RE.search(label) or label.lower() in {
            "superior hardwoods",
            "premium hardwoods",
            "walnut",
            "red oak",
            "rustic cherry",
            "sap cherry",
            "rustic hickory",
            "rustic qswo",
            "brown maple",
            "rustic brown maple",
            "cherry",
            "hard maple",
            "elm",
            "qswo",
            "hickory",
            "burnt hickory",
        }:
            out.append(ProcessedRow(RowKind.PROSE, [label]))
            continue
        out.append(ProcessedRow(RowKind.PROSE, [label]))
    return out


def _parse_laurel_bay_sheet1(rows: list[list[Any]], multiplier: float) -> tuple[list[ProcessedRow], dict[str, int]]:
    found = _find_order_header(rows)
    counts: dict[str, int] = {}
    out: list[ProcessedRow] = []
    if not found:
        return out, counts
    header_i, idxs = found
    # Laurel Bay = left block; Brantley = second ORDER # block
    left_end = idxs[1] if len(idxs) > 1 else len(rows[header_i])
    chunk, n = _parse_chair_block(
        rows,
        section_name="Laurel Bay",
        multiplier=multiplier,
        header_i=header_i,
        order_col=idxs[0],
        block_end=left_end,
    )
    out.extend(chunk)
    counts["Laurel Bay"] = n
    if len(idxs) > 1:
        chunk, n = _parse_chair_block(
            rows,
            section_name="Brantley",
            multiplier=multiplier,
            header_i=header_i,
            order_col=idxs[1],
            block_end=len(rows[header_i]),
        )
        out.extend(chunk)
        counts["Brantley"] = n
    return out, counts


def process_country_comfort(
    main_path: Path,
    laurel_path: Path | None = None,
    *,
    multiplier: float = MULTIPLIER,
) -> tuple[list[ProcessedRow], dict]:
    main_path = Path(main_path)
    wb = openpyxl.load_workbook(main_path, data_only=True)
    out: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [DOC_TITLE])]
    counts: dict[str, int] = {}

    table_sheets = {
        "Chelsea & Mayville",
        "Reno & Galveston",
        "Colver, Orchid, OW, Christy",
        "Cloverdale Bench",
    }

    for name in wb.sheetnames:
        if name in SKIP_MAIN:
            continue
        rows = _sheet_rows(wb[name])
        if name == "Options":
            out.extend(_parse_options(rows, "Options"))
            continue
        if name == "Options 2":
            out.extend(_parse_options(rows, "Wood Grades & Table Options"))
            continue
        if name == "Loraine":
            chunk, n = _parse_loraine(rows, multiplier)
            out.extend(chunk)
            counts[name] = n
            continue
        if name == "Ruff Sawn":
            chunk, n = _parse_ruff_sawn(rows, multiplier)
            out.extend(chunk)
            counts[name] = n
            continue
        if name in table_sheets:
            chunk, n = _parse_table_sheet(rows, name, multiplier)
            out.extend(chunk)
            counts[name] = n
            continue
        chunk, n = _parse_standard_collection(rows, name, multiplier)
        if n:
            out.extend(chunk)
            counts[name] = n

    wb.close()

    if laurel_path and Path(laurel_path).is_file():
        lb = openpyxl.load_workbook(laurel_path, data_only=True)
        for name in lb.sheetnames:
            if name in SKIP_LB:
                continue
            rows = _sheet_rows(lb[name])
            chunk, cdict = _parse_laurel_bay_sheet1(rows, multiplier)
            out.extend(chunk)
            counts.update(cdict)
        lb.close()

    out = collapse_blank_rows(out)
    stats = {
        "collections": counts,
        "items": sum(counts.values()),
        "multiplier": multiplier,
    }
    return out, stats


def country_comfort_forbidden(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            bad.extend(forbidden_phrases(str(c)))
            if re.search(r"\bunfin", str(c), re.I) and "finished" not in str(c).lower():
                # allow wood names that don't say unfinished
                if re.search(r"\bunfinished\b|\bunfin\b", str(c), re.I):
                    bad.append(str(c))
    return bad


def is_country_comfort_workbook(path: Path) -> bool:
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        return False
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = set(wb.sheetnames)
    wb.close()
    return "Adena" in names and "Christy" in names and "MARKUP" in names
