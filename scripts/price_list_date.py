"""Price list effective dates: parse, infer, and compare for builder uploads."""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

_MONTH = (
    r"January|February|March|April|May|June|July|August|September|"
    r"October|November|December"
)

EFFECTIVE_MONTH_DAY_YEAR = re.compile(
    rf"(?:Effective\s+)?({_MONTH})\s+(\d{{1,2}})(?:st|nd|rd|th)?,?\s+(\d{{4}})",
    re.I,
)
RETAIL_TITLE_SUFFIX = re.compile(r"Retail Price List\s*·\s*(.+)$", re.I)
FN_CHAIRS_TITLE = re.compile(
    rf"(\d{{4}})\s*·\s*Effective\s+({_MONTH})\s+(\d{{1,2}}),?\s+(\d{{4}})",
    re.I,
)
YEAR_ONLY = re.compile(r"^\s*(20\d{2})\s*$")
MMYY_IN_NAME = re.compile(r"[_\s](\d{2})(\d{2})(?:-\d+)?\.(?:xlsx|xls|pdf)?", re.I)


def _month_day_year(month: str, day: int, year: int) -> date | None:
    try:
        return datetime.strptime(f"{month} {day} {year}", "%B %d %Y").date()
    except ValueError:
        return None


def parse_price_list_phrase(text: str) -> date | None:
    """Parse a catalog effective / title date phrase into a calendar date."""
    if not text or not text.strip():
        return None
    t = text.strip()
    m = FN_CHAIRS_TITLE.search(t)
    if m:
        return _month_day_year(m.group(2), int(m.group(3)), int(m.group(4)))
    m = EFFECTIVE_MONTH_DAY_YEAR.search(t)
    if m:
        return _month_day_year(m.group(1), int(m.group(2)), int(m.group(3)))
    if YEAR_ONLY.fullmatch(t):
        return date(int(t), 1, 1)
    return None


def parse_iso_price_list_date(value: str) -> date | None:
    try:
        return date.fromisoformat(value)
    except (TypeError, ValueError):
        return None


def mmyy_from_path(path: Path) -> date | None:
    """Viztech-style version stamp in filenames (e.g. _0526 → May 1, 2026)."""
    m = MMYY_IN_NAME.search(path.name)
    if not m:
        return None
    mm, yy = int(m.group(1)), int(m.group(2))
    if not (1 <= mm <= 12):
        return None
    return date(2000 + yy, mm, 1)


def max_price_list_date(dates: list[date | None]) -> date | None:
    known = [d for d in dates if d is not None]
    return max(known) if known else None


def backfill_from_retail_pdf(pdf_path: Path, *, max_pages: int = 3) -> date | None:
    """Read printed retail PDF title / effective wording (catalog date, not 'Updated')."""
    try:
        import pdfplumber
    except ImportError:
        return None
    if not pdf_path.is_file():
        return None
    try:
        with pdfplumber.open(pdf_path) as pdf:
            chunks: list[str] = []
            for page in pdf.pages[:max_pages]:
                chunks.append(page.extract_text() or "")
        text = "\n".join(chunks)
    except Exception:
        return None

    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        m = RETAIL_TITLE_SUFFIX.search(line)
        if m:
            d = parse_price_list_phrase(m.group(1))
            if d:
                return d
        d = parse_price_list_phrase(line)
        if d:
            return d
    return None


def evaluate_price_list_update(
    existing: dict[str, Any] | None,
    incoming: date | None,
) -> tuple[bool, str]:
    """
  Decide whether an upload may replace an existing builder PDF.

  Returns (allow, human-readable reason).
  """
    if existing is None:
        return True, "new builder"
    if incoming is None:
        return False, "incoming price list date could not be determined; not replacing"
    stored = existing.get("price_list_date")
    if not stored:
        return (
            True,
            "existing builder has no stored price_list_date; accepting dated upload",
        )
    old = parse_iso_price_list_date(str(stored))
    if old is None:
        return (
            True,
            "stored price_list_date is invalid; accepting dated upload",
        )
    if incoming > old:
        return True, f"new list {incoming.isoformat()} is newer than stored {old.isoformat()}"
    if incoming == old:
        return False, f"same price list date as catalog ({old.isoformat()}); skipped"
    return False, (
        f"new list {incoming.isoformat()} is older than stored {old.isoformat()}; skipped"
    )


def infer_from_source(
    source: Path,
    builder_name: str,
    *,
    sheet: str | None = None,
) -> date | None:
    """Best-effort price list date from a dealer upload (before full transform)."""
    from scripts.country_view_oak_book import is_country_view_oak_workbook

    if is_country_view_oak_workbook(source):
        return date(2025, 4, 15)

    stamped = mmyy_from_path(source)
    if stamped is not None:
        return stamped

    suffix = source.suffix.lower()
    if suffix in {".xlsx", ".xls", ".csv"}:
        from scripts.read_dealer_file import DealerFileError, read_dealer_file

        try:
            grid = read_dealer_file(source, sheet_name=sheet)
        except DealerFileError:
            return None

        from scripts.artisan_book import extract_effective_label, is_artisan_style_grid
        from scripts.dealer_book import extract_catalog_date

        if is_artisan_style_grid(grid):
            label = extract_effective_label(grid)
            if label:
                return parse_price_list_phrase(label)

        catalog = extract_catalog_date(grid)
        if catalog:
            return parse_price_list_phrase(catalog)

        from scripts.forks_valley_book import is_forks_valley_bundle

        if is_forks_valley_bundle(source):
            return date(2026, 1, 1)

        from scripts.woodside_book import is_woodside_workbook, woodside_price_list_date

        if is_woodside_workbook(source):
            return woodside_price_list_date(source)

    return None


def infer_for_viztech_files(files: list[Path]) -> date | None:
    return max_price_list_date([mmyy_from_path(p) for p in files])
