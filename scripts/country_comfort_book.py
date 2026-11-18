"""Country Comfort wholesale workbooks → finished-only retail PDF rows (×2.7)."""

from __future__ import annotations

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
SKIP_LB = frozenset({"Markup"})  # Sheet2 holds Laurel Bay tables — import them

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
ERROR_RE = re.compile(r"^#(VALUE|REF|DIV/0|N/A|NAME|NULL|NUM)!", re.I)


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


def _is_error(v: Any) -> bool:
    if v is None:
        return False
    if isinstance(v, float) and v != v:  # NaN
        return True
    return bool(ERROR_RE.match(_cell(v)))


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


def _find_all_order_headers(rows: list[list[Any]]) -> list[tuple[int, list[int]]]:
    found: list[tuple[int, list[int]]] = []
    for i, row in enumerate(rows):
        idxs = [j for j, c in enumerate(row) if isinstance(c, str) and ORDER_RE.match(c.strip())]
        if idxs:
            found.append((i, idxs))
    return found


def _is_order_id(v: Any) -> bool:
    if v is None:
        return False
    if _is_num(v) and not isinstance(v, bool):
        # Whole-number SKUs / codes; reject obvious tiny add-on amounts later via context
        return float(v) == int(float(v)) and float(v) > 0
    s = _cell(v)
    if not s or not ORDER_TOKEN_RE.match(s):
        return False
    if s.lower() in {"order", "description", "unfin.", "fin.", "unfinished", "finished"}:
        return False
    return any(ch.isdigit() for ch in s)


def _looks_like_description(v: Any) -> bool:
    if not isinstance(v, str):
        return False
    t = v.strip()
    if not t or _is_num(v) or FIN_RE.match(t) or UNFIN_RE.match(t) or ORDER_RE.match(t):
        return False
    if t.upper() == "DESCRIPTION":
        return False
    return len(t) >= 3 and not PRICE_RE_LOCAL.match(t)


PRICE_RE_LOCAL = re.compile(r"^\$[\d,]+$")


def _detect_block_end(
    rows: list[list[Any]],
    header_i: int,
    order_col: int,
    second_order_col: int | None,
) -> int:
    """Left-collection end: second ORDER #, orphan DESCRIPTION, or SKU-as-price column."""
    header = rows[header_i]
    if second_order_col is not None:
        return second_order_col

    # Vanessa-style: DESCRIPTION without a preceding ORDER # (order #s sit in a Fin. column)
    for j, c in enumerate(header):
        if j <= order_col + 1:
            continue
        if isinstance(c, str) and c.strip().upper() == "DESCRIPTION":
            prev = _cell(header[j - 1]) if j > 0 else ""
            if not ORDER_RE.match(prev):
                # If previous header is Fin./Unfin., that column holds the paired order #s
                if FIN_RE.match(prev) or UNFIN_RE.match(prev):
                    return j - 1
                return j

    end = len(header)
    # Data-driven: Fin./Unfin. column whose first values are order IDs next to a description
    for r in range(header_i + 1, min(header_i + 10, len(rows))):
        row = rows[r]
        left_oid = row[order_col] if order_col < len(row) else None
        if not _is_order_id(left_oid):
            continue
        for j in range(order_col + 2, min(end, len(row))):
            v = row[j]
            if not _is_order_id(v):
                continue
            nxt = row[j + 1] if j + 1 < len(row) else None
            if _looks_like_description(nxt):
                # Same SKU in a calc block to the right — still cut here
                return j
    return end


def _col_has_finished_prices(
    rows: list[list[Any]], header_i: int, col: int, order_col: int
) -> bool:
    hits = 0
    for r in range(header_i + 1, min(header_i + 12, len(rows))):
        row = rows[r]
        if not _is_order_id(row[order_col] if order_col < len(row) else None):
            continue
        v = row[col] if col < len(row) else None
        if _is_error(v):
            continue
        if _is_num(v) and float(v) > 0 and not (
            _is_order_id(v)
            and col + 1 < len(row)
            and _looks_like_description(row[col + 1])
        ):
            hits += 1
    return hits >= 1


def _fin_col_holds_order_numbers(
    rows: list[list[Any]], header_i: int, fin_col: int, order_col: int
) -> bool:
    for r in range(header_i + 1, min(header_i + 10, len(rows))):
        row = rows[r]
        if not _is_order_id(row[order_col] if order_col < len(row) else None):
            continue
        v = row[fin_col] if fin_col < len(row) else None
        if _is_order_id(v) and fin_col + 1 < len(row) and _looks_like_description(row[fin_col + 1]):
            return True
    return False


def _fin_cols(
    header: list[Any],
    start: int,
    end: int,
    rows: list[list[Any]] | None = None,
    header_i: int = 0,
    order_col: int = 0,
) -> list[int]:
    """Finished-price columns, including unlabeled Hard Maple-style pairs; never order-# cells."""
    labeled: list[int] = []
    for j in range(start, min(end, len(header))):
        if FIN_RE.match(_cell(header[j])):
            labeled.append(j)

    # Pair origin: first Unfin./Fin. after description
    origin = start
    for j in range(start, min(end, len(header))):
        lab = _cell(header[j])
        if UNFIN_RE.match(lab):
            origin = j
            break
        if FIN_RE.match(lab):
            origin = j - 1 if j > start else j
            break

    inferred: list[int] = []
    j = origin
    while j + 1 < end:
        fin_j = j + 1
        unfin_lab = _cell(header[j]) if j < len(header) else ""
        fin_lab = _cell(header[fin_j]) if fin_j < len(header) else ""
        if FIN_RE.match(fin_lab):
            inferred.append(fin_j)
        elif UNFIN_RE.match(unfin_lab):
            # Fin. label missing next to Unfin. — still a price pair
            inferred.append(fin_j)
        elif rows is not None:
            # Both blank (Adena Hard Maple): wood name above + numeric Fin cell
            wood = _wood_label_for_fin_col(rows, header_i, fin_j)
            if wood and not wood.startswith("Wood "):
                if _col_has_finished_prices(rows, header_i, fin_j, order_col):
                    inferred.append(fin_j)
            elif _col_has_finished_prices(rows, header_i, fin_j, order_col) and _col_has_finished_prices(
                rows, header_i, j, order_col
            ):
                inferred.append(fin_j)
        j += 2

    cols = sorted(set(labeled) | set(inferred))
    cols = [c for c in cols if c < end]
    if rows is not None:
        cols = [
            c
            for c in cols
            if not _fin_col_holds_order_numbers(rows, header_i, c, order_col)
        ]
    return cols


def _wood_label_for_fin_col(rows: list[list[Any]], header_i: int, fin_col: int) -> str:
    """Gather wood names from rows above the ORDER # header for this Fin. column."""
    parts: list[str] = []
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
            # Normalize whitespace / newlines into single tokens
            for piece in re.split(r"[\n/]+", t):
                piece = piece.strip(" ,")
                if piece and piece not in parts:
                    parts.append(piece)
        if parts:
            break
    if not parts:
        return f"Wood {fin_col}"
    # Stack for PDF Paragraph wrapping; keep short
    if len(parts) == 1:
        return parts[0]
    return "\n".join(parts[:4])


def _parse_chair_block(
    rows: list[list[Any]],
    *,
    section_name: str,
    multiplier: float,
    header_i: int,
    order_col: int,
    block_end: int,
    skipped: list[dict] | None = None,
) -> tuple[list[ProcessedRow], int]:
    header = rows[header_i]
    desc_col = order_col + 1
    # Mid-sheet blocks (Cambria 2510) put DESCRIPTION next to ORDER # even if ORDER is not col 0
    for j in range(order_col + 1, min(order_col + 3, len(header))):
        if _cell(header[j]).upper() == "DESCRIPTION":
            desc_col = j
            break

    fin_cols = _fin_cols(
        header,
        order_col + 2 if desc_col == order_col + 1 else desc_col + 1,
        block_end,
        rows=rows,
        header_i=header_i,
        order_col=order_col,
    )
    # Cambria mid-block: Fin. labels may start a few cols after DESCRIPTION
    if not fin_cols:
        for j in range(desc_col + 1, block_end):
            if j < len(header) and FIN_RE.match(_cell(header[j])):
                fin_cols.append(j)
    if not fin_cols:
        return [], 0

    wood_headers = [_wood_label_for_fin_col(rows, header_i, c) for c in fin_cols]
    out: list[ProcessedRow] = [ProcessedRow(RowKind.SECTION, [section_name])]
    # Header as WIDE so pdf_builder keeps it with price rows (TABLE_HEADER goes to SKU buffer)
    out.append(ProcessedRow(RowKind.WIDE, ["Item #", "Description", *wood_headers]))
    count = 0
    for r in range(header_i + 1, len(rows)):
        row = rows[r]
        # Stop at next ORDER # header row (multi-block sheets)
        if any(isinstance(c, str) and ORDER_RE.match(c.strip()) for c in row):
            break
        oid = row[order_col] if order_col < len(row) else None
        if not _is_order_id(oid):
            label = _cell(oid) if oid is not None else ""
            if not label:
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
        all_zero_or_err = True
        saw_error = False
        for fc in fin_cols:
            v = row[fc] if fc < len(row) else None
            if _is_error(v):
                saw_error = True
                prices.append("—")
                continue
            if _is_num(v) and float(v) > 0:
                # Guard: never treat an order-number cell as a price
                if _is_order_id(v) and fc + 1 < len(row) and _looks_like_description(row[fc + 1]):
                    prices.append("—")
                    continue
                prices.append(_retail(float(v), multiplier))
                any_price = True
                all_zero_or_err = False
            else:
                if _is_num(v) and float(v) == 0:
                    pass
                else:
                    all_zero_or_err = False if _cell(v) else all_zero_or_err
                prices.append("—")
        item = _cell(oid)
        priced_n = sum(1 for p in prices if p != "—")
        # Stray single-cell formulas on mostly-blank (removed) rows — e.g. Arabella 2230B
        sparse_removed = priced_n > 0 and priced_n < 2 and len(fin_cols) >= 4
        # Also treat as removed when most finished cells are zero/error
        zeroish = 0
        for fc in fin_cols:
            v = row[fc] if fc < len(row) else None
            if _is_error(v) or (_is_num(v) and float(v) == 0) or v is None or _cell(v) == "":
                zeroish += 1
        mostly_blank = len(fin_cols) >= 4 and zeroish >= len(fin_cols) - 1 and priced_n <= 1
        # Catch #VALUE! sitting in Unfin. columns on the same row
        if not saw_error:
            for c in row[order_col:block_end]:
                if _is_error(c):
                    saw_error = True
                    break
        if not any_price or sparse_removed or mostly_blank:
            reason = "#VALUE!" if saw_error else "$0 / removed"
            if skipped is not None:
                skipped.append(
                    {"sheet": section_name, "item": item, "description": desc, "reason": reason}
                )
            continue
        out.append(ProcessedRow(RowKind.WIDE, [item, desc, *prices]))
        count += 1
    return out, count


def _parse_standard_collection(
    rows: list[list[Any]],
    sheet_name: str,
    multiplier: float,
    skipped: list[dict] | None = None,
) -> tuple[list[ProcessedRow], int]:
    headers = _find_all_order_headers(rows)
    if not headers:
        return [], 0
    out: list[ProcessedRow] = []
    total = 0
    # First header: left block only (ignore right calc/ORDER duplicate)
    header_i, order_idxs = headers[0]
    order_col = order_idxs[0]
    second = order_idxs[1] if len(order_idxs) > 1 else None
    end = _detect_block_end(rows, header_i, order_col, second)
    chunk, n = _parse_chair_block(
        rows,
        section_name=sheet_name,
        multiplier=multiplier,
        header_i=header_i,
        order_col=order_col,
        block_end=end,
        skipped=skipped,
    )
    out.extend(chunk)
    total += n

    # Additional mid-sheet ORDER # blocks (Cambria 2510, etc.) — only if ORDER col shifts
    for hi, idxs in headers[1:]:
        oc = idxs[0]
        # Skip duplicate right-side calc headers on the same row style as first
        if oc == order_col and hi == header_i:
            continue
        # Skip if this header is the paired right ORDER # on a two-ORDER row already handled
        if hi == header_i and len(order_idxs) > 1 and oc in order_idxs[1:]:
            continue
        sec = second if len(idxs) > 1 else None
        # For mid-sheet blocks without a second ORDER, end at row width / next content
        bend = _detect_block_end(rows, hi, oc, idxs[1] if len(idxs) > 1 else None)
        # Avoid re-parsing a right calc block that starts with same SKUs as left
        sample = rows[hi + 1] if hi + 1 < len(rows) else []
        sample_oid = sample[oc] if oc < len(sample) else None
        if (
            sample_oid is not None
            and out
            and any(
                pr.kind == RowKind.WIDE
                and pr.cells
                and str(pr.cells[0]) == _cell(sample_oid)
                for pr in out
            )
            and oc > 2
        ):
            # Likely a calc duplicate of an already-emitted item
            continue
        sub_name = sheet_name
        chunk, n = _parse_chair_block(
            rows,
            section_name=sub_name,
            multiplier=multiplier,
            header_i=hi,
            order_col=oc,
            block_end=bend,
            skipped=skipped,
        )
        # Merge into same collection: drop duplicate SECTION if continuing
        if chunk and chunk[0].kind == RowKind.SECTION and out:
            chunk = chunk[1:]
        out.extend(chunk)
        total += n
    return out, total


def _parse_loraine(rows: list[list[Any]], multiplier: float) -> tuple[list[ProcessedRow], int]:
    out = [ProcessedRow(RowKind.SECTION, ["Loraine"])]
    out.append(ProcessedRow(RowKind.WIDE, ["Item #", "Description", "Price"]))
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


def _parse_ruff_sawn(
    rows: list[list[Any]], multiplier: float, skipped: list[dict] | None = None
) -> tuple[list[ProcessedRow], int]:
    """Ruff Sawn sheet has Christy and Delaney subsections."""
    out: list[ProcessedRow] = []
    total = 0
    sections: list[tuple[str, int]] = []
    for i, row in enumerate(rows):
        c0 = _cell(row[0] if row else "")
        c1 = _cell(row[1] if len(row) > 1 else "")
        if c1 and not c0 and c1.lower() in {"christy", "delaney"}:
            sections.append((f"Ruff Sawn — {c1}", i))
        elif c0.lower() in {"christy", "delaney"}:
            sections.append((f"Ruff Sawn — {c0}", i))
    if not sections:
        return _parse_standard_collection(rows, "Ruff Sawn", multiplier, skipped=skipped)

    for si, (name, start) in enumerate(sections):
        end = sections[si + 1][1] if si + 1 < len(sections) else len(rows)
        block = rows[start:end]
        found = _find_order_header(block)
        if not found:
            continue
        hi, idxs = found
        bend = _detect_block_end(block, hi, idxs[0], idxs[1] if len(idxs) > 1 else None)
        chunk, n = _parse_chair_block(
            block,
            section_name=name,
            multiplier=multiplier,
            header_i=hi,
            order_col=idxs[0],
            block_end=bend,
            skipped=skipped,
        )
        out.extend(chunk)
        total += n
    return out, total


def _table_wood_header(cell: Any) -> str:
    raw = _cell(cell)
    parts = [p.strip(" ,") for p in re.split(r"[\n]+", raw)]
    parts = [p for p in parts if p]
    t = ", ".join(parts)
    t = re.sub(r"\s*,\s*", ", ", t)
    t = re.sub(r",\s*,+", ",", t)
    return t.strip(" ,") or "Finished"


def _parse_table_sheet(
    rows: list[list[Any]], sheet_name: str, multiplier: float
) -> tuple[list[ProcessedRow], int]:
    """Parse multi-block table size sheets into one or more collection sections."""
    out: list[ProcessedRow] = []
    total = 0

    titles_row = rows[0] if rows else []
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

    seen_blocks: set[str] = set()
    for title_col, title in title_cols:
        key = f"{title_col}:{title}"
        if key in seen_blocks:
            continue
        seen_blocks.add(key)
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
            ProcessedRow(RowKind.WIDE, headers),
        ]
        n = 0
        for row in rows:
            size = _cell(row[title_col]) if title_col < len(row) else ""
            if not SIZE_RE.match(size.replace(" ", "")):
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
    grade_buf: list[str] = []
    current_grade: str | None = None

    def flush_grade() -> None:
        nonlocal grade_buf, current_grade
        if current_grade and grade_buf:
            out.append(
                ProcessedRow(
                    RowKind.PROSE,
                    [f"{current_grade}: " + ", ".join(grade_buf)],
                )
            )
        elif current_grade:
            out.append(ProcessedRow(RowKind.SUBHEAD, [current_grade]))
        grade_buf = []
        current_grade = None

    wood_names = {
        "superior hardwoods",
        "premium hardwoods",
        "standard hardwoods",
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
    }

    for row in rows:
        label = _fix_typo(_cell(row[0] if row else ""))
        if not label or label.upper() == "OPTIONS":
            continue
        low = label.lower()
        if low.startswith("free options"):
            flush_grade()
            out.append(ProcessedRow(RowKind.SUBHEAD, [label]))
            continue
        if low in {"standard hardwoods", "superior hardwoods", "premium hardwoods"}:
            flush_grade()
            current_grade = label
            continue
        if low.startswith("add ") and PCT_RE.search(label):
            flush_grade()
            out.append(ProcessedRow(RowKind.SUBHEAD, [label]))
            continue
        if low.startswith("additional table"):
            flush_grade()
            out.append(ProcessedRow(RowKind.SUBHEAD, [label]))
            continue
        if current_grade and low in wood_names:
            grade_buf.append(label)
            continue

        flush_grade()
        amt = None
        for c in row[1:]:
            if _is_num(c):
                amt = float(c)
                break
        pct = re.search(r"\+\s*(\d+)\s*%", label)
        if pct:
            # Single prose line — avoid duplicated floating "+25%" boxes
            clean = re.sub(r"\s+", " ", label.replace("\n", " ")).strip()
            clean = re.sub(r"\s*=\s*\+\s*\d+\s*%", "", clean).strip(" =")
            out.append(ProcessedRow(RowKind.PROSE, [f"{clean}: +{pct.group(1)}%"]))
            continue
        if amt is not None:
            lab, price = _format_flat_add(label.strip(), amt, MULTIPLIER)
            if lab:
                out.append(ProcessedRow(RowKind.WIDE, [lab, price]))
            continue
        if "=" in label:
            clean = re.sub(r"\s+", " ", label.replace("\n", " ")).strip()
            clean = sanitize_retail_text(clean, MULTIPLIER, allow_percent=True) or clean
            # Flat dollars embedded after = → multiply
            m = re.search(r"=\s*\+\s*\$?\s*([\d,]+)", clean)
            if m and not PCT_RE.search(clean):
                lab, price = _format_flat_add(clean.split("=")[0].strip(), float(m.group(1).replace(",", "")), MULTIPLIER)
                if lab:
                    out.append(ProcessedRow(RowKind.WIDE, [lab, price]))
                    continue
            out.append(ProcessedRow(RowKind.PROSE, [clean]))
            continue
        if PCT_RE.search(label) or low in wood_names:
            out.append(ProcessedRow(RowKind.PROSE, [label]))
            continue
        out.append(ProcessedRow(RowKind.PROSE, [label]))
    flush_grade()
    return out


def _parse_laurel_bay_sheet1(
    rows: list[list[Any]], multiplier: float, skipped: list[dict] | None = None
) -> tuple[list[ProcessedRow], dict[str, int]]:
    found = _find_order_header(rows)
    counts: dict[str, int] = {}
    out: list[ProcessedRow] = []
    if not found:
        return out, counts
    header_i, idxs = found
    left_end = _detect_block_end(
        rows, header_i, idxs[0], idxs[1] if len(idxs) > 1 else None
    )
    chunk, n = _parse_chair_block(
        rows,
        section_name="Laurel Bay",
        multiplier=multiplier,
        header_i=header_i,
        order_col=idxs[0],
        block_end=left_end,
        skipped=skipped,
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
            skipped=skipped,
        )
        out.extend(chunk)
        counts["Brantley"] = n
    return out, counts


def _parse_laurel_bay_tables(
    rows: list[list[Any]], multiplier: float
) -> tuple[list[ProcessedRow], int]:
    """Sheet2: Laurel Bay leg-table configurations (65 codes)."""
    wood_headers = ["Red Oak", "Rustic / Maple group", "Cherry / Maple / Elm / QSWO"]
    # Prefer labels from header row when present
    for row in rows[:5]:
        labs = [_table_wood_header(c) for c in row[6:9] if _cell(c)]
        if len(labs) >= 3:
            wood_headers = labs[:3]
            break

    out: list[ProcessedRow] = [
        ProcessedRow(RowKind.SECTION, ["Laurel Bay"]),
        ProcessedRow(
            RowKind.WIDE,
            ["Item #", "Size", "Leaf / Top", *wood_headers],
        ),
    ]
    count = 0
    for row in rows:
        code = _cell(row[0] if row else "")
        if not code or "LAB" not in code.upper():
            continue
        size = _cell(row[1] if len(row) > 1 else "")
        leaf = _cell(row[2] if len(row) > 2 else "")
        prices = []
        any_p = False
        for col in (6, 7, 8):
            v = row[col] if col < len(row) else None
            if _is_num(v) and float(v) > 0:
                prices.append(_retail(float(v), multiplier))
                any_p = True
            else:
                prices.append("—")
        if not any_p:
            continue
        out.append(ProcessedRow(RowKind.WIDE, [code, size, leaf, *prices]))
        count += 1
    if count == 0:
        return [], 0
    return out, count


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
    skipped: list[dict] = []

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
            chunk, n = _parse_ruff_sawn(rows, multiplier, skipped=skipped)
            out.extend(chunk)
            counts[name] = n
            continue
        if name in table_sheets:
            chunk, n = _parse_table_sheet(rows, name, multiplier)
            out.extend(chunk)
            counts[name] = n
            continue
        chunk, n = _parse_standard_collection(rows, name, multiplier, skipped=skipped)
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
            if name == "Sheet2":
                chunk, n = _parse_laurel_bay_tables(rows, multiplier)
                if n:
                    # Merge count into Laurel Bay
                    out.extend(chunk)
                    counts["Laurel Bay"] = counts.get("Laurel Bay", 0) + n
                    counts["Laurel Bay Tables"] = n
                continue
            chunk, cdict = _parse_laurel_bay_sheet1(rows, multiplier, skipped=skipped)
            out.extend(chunk)
            for k, v in cdict.items():
                counts[k] = counts.get(k, 0) + v
        lb.close()

    out = collapse_blank_rows(out)
    stats = {
        "collections": counts,
        "items": sum(v for k, v in counts.items() if k != "Laurel Bay Tables"),
        "multiplier": multiplier,
        "skipped": skipped,
        "table_configs": sum(
            counts.get(s, 0)
            for s in (
                "Chelsea & Mayville",
                "Reno & Galveston",
                "Colver, Orchid, OW, Christy",
                "Cloverdale Bench",
            )
        )
        + counts.get("Laurel Bay Tables", 0),
    }
    return out, stats


def country_comfort_forbidden(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            bad.extend(forbidden_phrases(str(c)))
            if re.search(r"\bunfin", str(c), re.I) and "finished" not in str(c).lower():
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
