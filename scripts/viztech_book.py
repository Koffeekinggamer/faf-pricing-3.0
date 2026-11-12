"""Process Viztech-exported dealer workbooks into retail ProcessedRows."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.dealer_terms import is_dealer_terms_text
from scripts.pdf_plain_text import clean_pdf_plain_text
from scripts.retail_text import forbidden_phrases, retail_price_cell, sanitize_retail_text

try:
    import xlrd
except ImportError:
    xlrd = None

MULTIPLIER_DEFAULT = 2.7

SKIP_SHEET_RE = re.compile(
    r"(?:^markup$|^mark-?up$|^cover$|wholesale\s*markup|finished\s*wholesale\s*markup|"
    r"^finished\s*wholesale\s*markup$|instructions|^admin$|controls|header\s*selection|"
    r"supplement\s*markup|detail\s*prices|customer\s*letter|cheat\s*sheet|"
    r"do\s*not\s*print|\(for\s*wholesale\)|internet\s*policy.*wholesale|^pricing$)",
    re.I,
)
UNFINISHED_SHEET_RE = re.compile(r"unfinished|retail\s*unf", re.I)
ITEM_CODE_RE = re.compile(
    r"^(?:[A-Z]{1,4}[-\d]|[\d]{3,5}[\s-][A-Z]{2}|#\d|M\d|MI\d|FV\d|[\d]+-[\dA-Z])",
    re.I,
)
MMYY_RE = re.compile(r"[_\s](\d{2})(\d{2})(?:-\d+)?\.(?:xlsx|xls)", re.I)


def normalize_builder_name(name: str) -> str:
    n = (name or "").strip().lstrip("@")
    return n.replace("\u2019", "'").replace("\u2018", "'").replace("\u201c", '"').replace("\u201d", '"')


def file_mmyy(path: Path) -> int:
    m = MMYY_RE.search(path.name)
    if not m:
        return 0
    return int(m.group(1)) * 100 + int(m.group(2))


def section_title_from_path(path: Path) -> str:
    n = path.stem
    for pat, label in (
        (r"Master", "Master"),
        (r"Quick_Ship", "Quick Ship"),
        (r"Urban_Essence", "Urban Essence"),
        (r"Ellison_Office", "Ellison Office"),
        (r"CYO_Program", "CYO Program"),
        (r"Spring_New_Introductions", "Spring New Introductions"),
        (r"Urban_Collection", "Urban Collection"),
        (r"Benning_Bookcase", "Benning Bookcase"),
        (r"Versa_Loft", "Versa Loft Beds"),
        (r"Luxhome", "LuxHome"),
        (r"Finished", "Finished"),
        (r"Pricelist", "Price list"),
    ):
        if re.search(pat, n, re.I):
            return label
    return "Price list"


def section_dedupe_key(path: Path) -> str:
    return section_title_from_path(path).casefold()


@dataclass
class SpotSample:
    item: str
    wholesale: float
    retail: str


@dataclass
class ViztechBuildResult:
    rows: list[ProcessedRow]
    spots: list[SpotSample] = field(default_factory=list)
    skipped_sheets: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _cell_str(c: Any) -> str:
    if c is None:
        return ""
    if isinstance(c, float) and c == int(c):
        return str(int(c))
    return str(c).strip()


def _is_wholesale_price(val: Any) -> bool:
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return False
    v = float(val)
    return 1 <= v <= 250_000


def _load_xlsx_sheets(path: Path) -> dict[str, list[list]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out: dict[str, list[list]] = {}
    for name in wb.sheetnames:
        ws = wb[name]
        out[name] = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return out


def _load_xls_sheets(path: Path) -> dict[str, list[list]]:
    if xlrd is None:
        raise RuntimeError("xlrd required for .xls")
    book = xlrd.open_workbook(str(path))
    out: dict[str, list[list]] = {}
    for i, name in enumerate(book.sheet_names()):
        sh = book.sheet_by_index(i)
        out[name] = [sh.row_values(r) for r in range(sh.nrows)]
    return out


def load_workbook_sheets(path: Path) -> dict[str, list[list]]:
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        return _load_xlsx_sheets(path)
    if path.suffix.lower() == ".xls":
        return _load_xls_sheets(path)
    raise ValueError(f"Unsupported workbook: {path}")


def _drop_row(row: list, joined: str) -> bool:
    if is_dealer_terms_text(joined):
        return True
    low = joined.lower()
    if any(
        x in low
        for x in (
            "viztech",
            "purchase order",
            "order confirmation",
            "by fax",
            "payment terms",
            "table of contents",
            "invoiced from",
            "std. terms",
            "std terms",
        )
    ):
        return True
    if "wholesale price list" in low or re.match(r"^finished\s+wholesale\s*$", low):
        return True
    if re.search(r"\bnet\s*30\b", low):
        return True
    if re.search(r"\b(fax|phone|e-?mail|@)\b", low) and "$" not in joined:
        if len(joined) < 120:
            return True
    return False


def _finished_price_cols(header_rows: list[list], width: int) -> set[int]:
    unfinished: set[int] = set()
    finished: set[int] = set()
    for h in header_rows:
        for j in range(min(width + 1, len(h))):
            c = h[j]
            if not isinstance(c, str):
                continue
            s = c.lower()
            if "unfinished" in s and "finished" not in s:
                unfinished.add(j)
            elif "finished" in s:
                finished.add(j)
    if finished:
        return finished
    # Barkman-style: skip col1 markup fractions; take numeric cols not in unfinished set
    return set()


def _row_item_label(row: list) -> tuple[str, int]:
    for j, c in enumerate(row[:6]):
        if isinstance(c, (int, float)) and not isinstance(c, bool):
            v = float(c)
            if 100 <= v <= 999999 and j == 0:
                return str(int(v)), j
        s = _cell_str(c)
        if not s:
            continue
        if ITEM_CODE_RE.match(s) or (len(s) > 8 and re.search(r"\d", s)):
            return s, j
        if j <= 1 and len(s) > 12 and not _is_wholesale_price(c):
            return s, j
    return "", -1


def _transform_row_prices(
    row: list,
    multiplier: float,
    price_cols: Iterable[int],
    *,
    spots: list[SpotSample],
    item_hint: str,
) -> list[str]:
    cells: list[str] = []
    label, _ = _row_item_label(row)
    if not label:
        label = item_hint
    if re.search(r"approximate\s+weight", label, re.I):
        return []
    lab = sanitize_retail_text(label, multiplier)
    if not lab:
        return []
    cells.append(lab)
    got_price = False
    for j in sorted(set(price_cols)):
        if j >= len(row):
            cells.append("")
            continue
        v = row[j]
        if _is_wholesale_price(v):
            wholesale = float(v)
            retail = retail_price_cell(v, multiplier)
            cells.append(retail)
            got_price = True
            if len(spots) < 3:
                spots.append(SpotSample(item=label[:80] or f"col{j}", wholesale=wholesale, retail=retail))
        else:
            t = sanitize_retail_text(_cell_str(v), multiplier)
            cells.append(t if t else "")
    while cells and not cells[-1]:
        cells.pop()
    if not got_price:
        return []
    return cells


def _infer_price_cols(rows: list[list], start: int) -> set[int]:
    cols: set[int] = set()
    header_block = rows[start : start + 12]
    width = max(len(r) for r in rows[start : start + 40]) if rows else 0
    labeled = _finished_price_cols(header_block, width)
    if labeled:
        return labeled
    for row in rows[start : start + 200]:
        for j, c in enumerate(row):
            if _is_wholesale_price(c):
                cols.add(j)
    # Bamm's: prefer right half if mirror empty left
    if cols and max(cols) >= 4:
        left = {j for j in cols if j < 4}
        right = {j for j in cols if j >= 4}
        if right and len(left) < len(right):
            cols = right
    if not cols:
        return set()
    # Drop unfinished-only columns when paired (odd/even): col with unfinished header
    bad: set[int] = set()
    for h in header_block:
        for j, c in enumerate(h):
            if isinstance(c, str) and c.strip().lower() == "unfinished":
                bad.add(j)
    return {j for j in cols if j not in bad}


def _process_sheet(
    name: str,
    rows: list[list],
    *,
    multiplier: float,
    spots: list[SpotSample],
) -> list[ProcessedRow]:
    out: list[ProcessedRow] = []
    clean_name = sanitize_retail_text(name.replace("_", " "), multiplier) or name
    if clean_name.lower() not in ("wholesale", "finished wholesale"):
        out.append(ProcessedRow(RowKind.SUBHEAD, [clean_name]))
    item_hint = ""
    i = 0
    while i < len(rows):
        row = rows[i]
        if row_is_blank(row):
            out.append(ProcessedRow(RowKind.BLANK, []))
            i += 1
            continue
        joined = clean_pdf_plain_text(" ".join(_cell_str(c) for c in row if _cell_str(c)))
        if _drop_row(row, joined):
            i += 1
            continue
        if re.search(r"\b(unfinished|deduct\s+\d+\s*%).*unfinished", joined, re.I):
            i += 1
            continue
        price_cols = _infer_price_cols(rows, i)
        cells = _transform_row_prices(row, multiplier, price_cols or range(len(row)), spots=spots, item_hint=item_hint)
        if cells:
            out.append(ProcessedRow(RowKind.WIDE, cells))
            i += 1
            continue
        label, _ = _row_item_label(row)
        if label and not price_cols:
            nums = [j for j, c in enumerate(row) if _is_wholesale_price(c)]
            if nums:
                cells = _transform_row_prices(row, multiplier, nums, spots=spots, item_hint=label)
                if cells:
                    out.append(ProcessedRow(RowKind.WIDE, cells))
                    item_hint = label
                    i += 1
                    continue
            item_hint = label
        prose_cells: list[str] = []
        for c in row:
            if _is_wholesale_price(c):
                continue
            raw = _cell_str(c)
            if re.search(r"approximate\s+weight", raw, re.I):
                continue
            t = sanitize_retail_text(raw, multiplier)
            if t and len(t) > 3:
                prose_cells.append(t)
        if prose_cells and len(" ".join(prose_cells)) > 40:
            out.append(ProcessedRow(RowKind.PROSE, prose_cells[:4]))
        i += 1
    return out


def luxhome_has_finished_products(path: Path) -> bool:
    sheets = load_workbook_sheets(path)
    for name, rows in sheets.items():
        if SKIP_SHEET_RE.search(name) or UNFINISHED_SHEET_RE.search(name):
            continue
        for row in rows[:200]:
            if any(_is_wholesale_price(c) for c in row):
                return True
    return False


def process_viztech_workbook(
    path: Path,
    *,
    multiplier: float = MULTIPLIER_DEFAULT,
    builder_name: str | None = None,
) -> ViztechBuildResult:
    sheets = load_workbook_sheets(path)
    spots: list[SpotSample] = []
    skipped: list[str] = []
    out: list[ProcessedRow] = []
    has_grid = False

    for name, rows in sheets.items():
        if SKIP_SHEET_RE.search(name):
            skipped.append(name)
            continue
        if UNFINISHED_SHEET_RE.search(name):
            skipped.append(name)
            continue
        if not rows or not any(any(c not in (None, "") for c in r) for r in rows):
            skipped.append(name)
            continue
        before = len(out)
        out.extend(_process_sheet(name, rows, multiplier=multiplier, spots=spots))
        if len(out) > before:
            has_grid = True

    if not has_grid:
        return ViztechBuildResult([], spots, skipped, [f"No price grid in {path.name}"])

    return ViztechBuildResult(collapse_blank_rows(out), spots, skipped, [])


def merge_builder_files(
    paths: list[Path],
    *,
    builder_name: str,
    multiplier: float,
) -> ViztechBuildResult:
    key = builder_name.casefold()
    if key == "aj's furniture":
        paths = [
            p
            for p in paths
            if "finished" in p.name.lower()
            or ("luxhome" in p.name.lower() and luxhome_has_finished_products(p))
        ]
        if not paths:
            return ViztechBuildResult([], [], [], ["No finished price file for AJ's Furniture"])
    # Keep newest MMYY per section key
    by_section: dict[str, Path] = {}
    for p in sorted(paths, key=file_mmyy):
        key = section_dedupe_key(p)
        prev = by_section.get(key)
        if prev is None or file_mmyy(p) >= file_mmyy(prev):
            by_section[key] = p

    title = f"{builder_name} Retail Price List"
    all_rows: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [title])]
    spots: list[SpotSample] = []
    skipped: list[str] = []
    warnings: list[str] = []

    order = sorted(by_section.items(), key=lambda kv: section_title_from_path(kv[1]))
    for _key, path in order:
        if "luxhome" in path.name.lower() and not luxhome_has_finished_products(path):
            warnings.append(f"Skipped LuxHome file (no finished prices): {path.name}")
            continue
        sec = section_title_from_path(path)
        all_rows.append(ProcessedRow(RowKind.SECTION, [sec]))
        res = process_viztech_workbook(path, multiplier=multiplier, builder_name=builder_name)
        if res.warnings:
            warnings.extend(res.warnings)
        if not res.rows:
            warnings.append(f"No rows from {path.name}")
            continue
        all_rows.extend(res.rows)
        skipped.extend(res.skipped_sheets)
        spots.extend(res.spots)

    if len(all_rows) <= 1:
        return ViztechBuildResult([], spots, skipped, warnings + ["No content after merge"])

    return ViztechBuildResult(collapse_blank_rows(all_rows), spots[:3], skipped, warnings)


def viztech_forbidden(rows: list[ProcessedRow]) -> list[str]:
    bad: list[str] = []
    for pr in rows:
        for c in pr.cells:
            bad.extend(forbidden_phrases(str(c)))
            if is_dealer_terms_text(str(c)):
                bad.append(str(c))
    return bad
