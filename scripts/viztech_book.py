"""Process Viztech-exported dealer workbooks into retail ProcessedRows."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.dealer_terms import is_builder_contact_text, is_dealer_terms_text
from scripts.pdf_plain_text import clean_pdf_plain_text
from scripts.retail_text import (
    excel_percent_to_display,
    forbidden_phrases,
    is_percent_token,
    retail_price_cell,
    sanitize_retail_text,
)

try:
    import xlrd
except ImportError:
    xlrd = None

MULTIPLIER_DEFAULT = 2.7

SKIP_SHEET_RE = re.compile(
    r"(?:^markup$|^mark-?up$|^cover$|wholesale\s*markup|finished\s*wholesale\s*markup|"
    r"^finished\s*wholesale\s*markup$|instructions|^admin$|controls|header\s*selection|"
    r"supplement\s*markup|detail\s*prices|customer\s*letter|cheat\s*sheet|"
    r"do\s*not\s*print|\(for\s*wholesale\)|internet\s*policy.*wholesale|^pricing$|"
    r"price_multiplier|_calc|(?:^|_)calc$|for\s*print|markup\s*page|contact\s*and\s*info|"
    r"^fabric\s*list$|back\s*cover|retail\s*cover|retail\s*options|retail\s*furniture|"
    r"^2026\s*retail$|retail\s*price\s*list\s*markup|^increase\s*entry$)",
    re.I,
)
UNFINISHED_SHEET_RE = re.compile(r"unfinished|retail\s*unf", re.I)
ITEM_CODE_RE = re.compile(
    r"^(?:[A-Z]{1,4}[-\d]|[\d]{3,5}[\s-][A-Z]{2}|#\d|M\d|MI\d|FV\d|[\d]+-[\dA-Z])",
    re.I,
)
MMYY_RE = re.compile(r"[_\s](\d{2})(\d{2})(?:-\d+)?\.(?:xlsx|xls)", re.I)
FINISH_COST_RE = re.compile(r"cost\s+to\s+(?:stain|paint)\s+finish", re.I)
RETAIL_MARKUP_COL_RE = re.compile(
    r"sug\.?\s*retail|suggested\s*retail|\bmap\b|\bmsrp\b|"
    r"internet\s*prices?|pallet\s*qtys?|pallet\s*price",
    re.I,
)
SEALED_PAINTED_RE = re.compile(r"sealed\s*/\s*painted|sealed\s+and\s+painted", re.I)
# Option / note lines: drop contact-us / call-for-pricing clauses; keep the rest.
CONTACT_PRICING_PHRASE_RE = re.compile(
    r"(?:^|[;\s]+)"
    r"(?:"
    r"contact\s+us\s+for\s+pricing(?:\s+on\s+[^.;]*)?|"
    r"call\s+(?:us\s+)?for\s+pricing(?:\s+on\s+[^.;]*)?"
    r")"
    r"\.?",
    re.I,
)


def _strip_contact_pricing_phrases(text: str) -> str:
    """Remove 'contact us' / 'call for pricing' clauses from notes; keep other wording."""
    if not text or not str(text).strip():
        return ""
    t = CONTACT_PRICING_PHRASE_RE.sub(" ", str(text))
    t = re.sub(r"\s{2,}", " ", t).strip()
    t = re.sub(r"^[;\s,.-]+|[;\s,.-]+$", "", t).strip()
    # Restore a trailing period when the kept clause still looks like a sentence.
    if t and not t.endswith((".", "!", "?", '"', "'")) and re.search(
        r"\b(?:width|size|sizes|only|offered)\b", t, re.I
    ):
        t = t + "."
    return t


def normalize_builder_name(name: str) -> str:
    n = (name or "").strip().lstrip("@")
    return n.replace("\u2019", "'").replace("\u2018", "'").replace("\u201c", '"').replace("\u201d", '"')


def file_mmyy(path: Path) -> int:
    m = MMYY_RE.search(path.name)
    if not m:
        return 0
    return int(m.group(1)) * 100 + int(m.group(2))


def _section_label_from_upload_name(path: Path) -> str | None:
    """Viztech upload names: builder-slug__CODE_2026_Section_Pricelist_…"""
    m = re.search(r"__\w+_\d{4}_(.+?)_Pricelist", path.name, re.I)
    if not m:
        return None
    raw = m.group(1).replace("_", " ")
    raw = re.sub(r"\bevry\b", "Every", raw, flags=re.I)
    raw = re.sub(r"([a-z])([A-Z])", r"\1 \2", raw)
    words: list[str] = []
    for w in raw.split():
        if w.lower() in {"for", "and", "of", "the"}:
            words.append(w.lower())
        else:
            words.append(w[:1].upper() + w[1:] if w else w)
    return " ".join(words).strip() or None


def section_title_from_path(path: Path) -> str:
    fname = path.name
    from_upload = _section_label_from_upload_name(path)
    if from_upload:
        return from_upload
    if re.search(r"buckeye-rockers__AL_", fname, re.I):
        return "Authentex"
    if re.search(r"buckeye-rockers__BR_", fname, re.I):
        return "Standard"
    if re.search(r"cupola-furniture__.*LuxWood", fname, re.I):
        return "LuxWood"
    if re.search(r"cupola-furniture__", fname, re.I):
        return "Standard"
    if re.search(r"Bath_Cabinets", fname, re.I):
        return "Bath Cabinets"
    if re.search(r"Bedroom_Tables", fname, re.I):
        return "Bedroom, Tables & Barstools"
    if re.search(r"Work_Stations", fname, re.I):
        return "Work Stations"
    if re.search(r"_Slabs?_", fname, re.I):
        return "Slabs"
    if re.search(r"_Barrel_", fname, re.I):
        return "Barrel"
    if re.search(r"_Saw_Marked_", fname, re.I):
        return "Saw Marked"
    if re.search(r"_Dining_", fname, re.I):
        return "Dining"
    if re.search(r"Cyp\s*and\s*Treated|Cypress.*Treated", fname, re.I):
        return "Cypress & Treated"
    if re.search(r"Poly[_\s-]*Aurora", fname, re.I):
        return "Poly Aurora"
    if re.search(r"Poly[_\s-]*Eclipse", fname, re.I):
        return "Poly Eclipse"
    if re.search(r"Internet[_\s-]*pricing|HW_Internet", fname, re.I):
        return "Internet"
    if re.search(r"New_Products_Pricelist", fname, re.I):
        return "New Products"
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
        (r"LuxWood|Luxwood", "LuxWood"),
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
    # Avoid dumping Excel datetime as "2026-03-24 00:00:00" into retail PDFs
    try:
        from datetime import date, datetime

        if isinstance(c, datetime):
            if c.hour == 0 and c.minute == 0 and c.second == 0:
                return f"{c.strftime('%b')} {c.day}, {c.year}"
            return c.isoformat(sep=" ")
        if isinstance(c, date):
            return f"{c.strftime('%b')} {c.day}, {c.year}"
    except Exception:
        pass
    return str(c).strip()


def _is_wholesale_price(val: Any) -> bool:
    if isinstance(val, str) and val.strip().upper() in {"#N/A", "N/A", "NA", "#NA"}:
        return False
    if is_percent_token(val):
        return False
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return False
    v = float(val)
    return 1 <= v <= 250_000


def _is_na_cell(val: Any) -> bool:
    return isinstance(val, str) and val.strip().upper() in {"#N/A", "N/A", "NA", "#NA"}


def _is_percent_token(val: Any) -> bool:
    return is_percent_token(val)


def _excel_format_is_percent(number_format: str | None) -> bool:
    if not number_format:
        return False
    # Ignore date/time formats that can contain stray %
    fmt = str(number_format)
    if re.search(r"[ymdhs]", fmt, re.I) and "%" not in fmt.replace("%%", ""):
        return False
    return "%" in fmt.replace("%%", "")


def _load_xlsx_sheets(path: Path) -> dict[str, list[list]]:
    """Load values; convert percent-formatted numerics to 'N%' display tokens."""
    wb_val = openpyxl.load_workbook(path, data_only=True, read_only=True)
    wb_fmt = openpyxl.load_workbook(path, data_only=False, read_only=True)
    out: dict[str, list[list]] = {}
    for name in wb_val.sheetnames:
        ws_v = wb_val[name]
        ws_f = wb_fmt[name] if name in wb_fmt.sheetnames else None
        rows: list[list] = []
        fmt_rows = ws_f.iter_rows() if ws_f is not None else None
        for r_v in ws_v.iter_rows():
            fmt_row = next(fmt_rows) if fmt_rows is not None else None
            row: list = []
            for idx, c_v in enumerate(r_v):
                v = c_v.value
                fmt = ""
                if fmt_row is not None and idx < len(fmt_row):
                    fmt = fmt_row[idx].number_format or ""
                if (
                    isinstance(v, (int, float))
                    and not isinstance(v, bool)
                    and _excel_format_is_percent(fmt)
                ):
                    row.append(excel_percent_to_display(float(v)))
                else:
                    row.append(v)
            rows.append(row)
        out[name] = rows
    wb_val.close()
    wb_fmt.close()
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
    # Ignore raw wholesale numbers when classifying dealer/contact prose
    text_only = clean_pdf_plain_text(
        " ".join(
            _cell_str(c)
            for c in row
            if _cell_str(c) and not _is_wholesale_price(c)
        )
    )
    probe = text_only or joined
    if is_dealer_terms_text(probe):
        return True
    low = probe.lower()
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
            "internet stores",
            "internet price",
            "visa, mastercard",
            "showroom floor",
            "password",
            "viztechfurniture",
            "my account",
            "register",
            "lagwana.cloud",
        )
    ):
        return True
    if re.search(
        r"\bdiscount\b.*\b(?:payment|received|days)\b|\bpicked\s+up\b.*\d+\s+days\b|"
        r"\bconsidered\s+stock\b|instructions\s*-\s*please\s+read|\d+\s*%\s*discount\b|"
        r"\bfinshing\b|\bfinishing\b.*\bx\s*200\s*%|"
        r"approximate\s+finish(?:ing)?\s+costs?|"
        r"internet\s+policy",
        probe,
        re.I,
    ):
        return True
    if "wholesale price list" in low or re.match(r"^finished\s+wholesale\s*$", low):
        return True
    if re.search(r"\bnet\s*30\b", low):
        return True
    if is_builder_contact_text(probe):
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
            s = c.lower().strip()
            if RETAIL_MARKUP_COL_RE.search(s):
                continue
            if FINISH_COST_RE.search(s):
                # Heritage-style finish add-on, not a finished product price
                continue
            if re.fullmatch(r"unf\.?", s) or (
                re.search(r"\bunfinished\d*\b", s) and not re.search(r"\bfinished\b", s)
            ):
                unfinished.add(j)
            elif re.fullmatch(r"fin\.?", s) or re.search(r"\bfinished\b", s):
                finished.add(j)
            elif SEALED_PAINTED_RE.search(s):
                finished.add(j)
            elif re.search(r"\b(?:stain(?:ed)?|glaz(?:e|ing)|2[\s-]?tone|paint)\b", s):
                # Fishers / Hilltop: Stain / Stained / 2-Tone & Glazing are finished
                finished.add(j)
    # Hoosier-style "FINISHED PRICES" band: expand through following wood headers
    band_starts = [
        j
        for h in header_rows
        for j, c in enumerate(h)
        if isinstance(c, str) and re.search(r"(?<!un)finished\s*prices?", c, re.I)
    ]
    if band_starts:
        start = min(band_starts)
        expanded: set[int] = set()
        for h in header_rows:
            for j in range(start, min(width + 1, len(h))):
                c = h[j]
                if isinstance(c, str) and c.strip():
                    if re.search(r"\bunfinished\b", c, re.I) and not re.search(
                        r"(?<!un)finished", c, re.I
                    ):
                        break
                    expanded.add(j)
                elif c not in (None, "") and _is_wholesale_price(c):
                    expanded.add(j)
        if expanded:
            return expanded - unfinished
    if finished:
        return finished - unfinished
    # Barkman-style: skip col1 markup fractions; take numeric cols not in unfinished set
    return set()


def _excluded_price_header_cols(header_rows: list[list], width: int) -> set[int]:
    """Columns that look numeric but are retail markups, pallet qty, or unfinished bases."""
    bad: set[int] = set()
    sealed: set[int] = set()
    for h in header_rows:
        for j in range(min(width + 1, len(h))):
            c = h[j]
            if not isinstance(c, str):
                continue
            s = c.lower().strip()
            if RETAIL_MARKUP_COL_RE.search(s):
                bad.add(j)
            if SEALED_PAINTED_RE.search(s):
                sealed.add(j)
            if re.fullmatch(r"price\s*ea\.?", s) or s in {"price ea", "price ea."}:
                bad.add(j)
            if FINISH_COST_RE.search(s):
                bad.add(j)
    # When Sealed/Painted exists, drop other Hershy wholesale tiers (Price Ea already bad)
    if sealed:
        for h in header_rows:
            for j in range(min(width + 1, len(h))):
                c = h[j]
                if not isinstance(c, str):
                    continue
                s = c.lower().strip()
                if j in sealed:
                    continue
                if re.search(r"price\s*ea|pallet|sug|retail|wholesale\s*price", s):
                    bad.add(j)
    return bad


@dataclass
class FinishCostLayout:
    """Unfinished wood bases + finish add-on → composed finished wholesale."""

    wood_cols: list[int]
    stain_col: int | None
    paint_col: int | None
    top_col: int | None


def _detect_finish_cost_layout(
    rows: list[list], start: int
) -> FinishCostLayout | None:
    """Heritage-style: unfinished woods + COST TO STAIN FINISH mirrored into trailing bases."""
    header_block, width = _header_block_near(rows, start)
    unfinished_hdr: set[int] = set()
    stain_hdr: int | None = None
    paint_hdr: int | None = None
    top_hdr: int | None = None
    for h in header_block:
        for j in range(min(width + 1, len(h))):
            c = h[j]
            if not isinstance(c, str):
                continue
            s = c.lower()
            if FINISH_COST_RE.search(s) and "paint" in s:
                paint_hdr = j if paint_hdr is None else paint_hdr
            elif FINISH_COST_RE.search(s):
                stain_hdr = j if stain_hdr is None else stain_hdr
            elif re.search(r"island\s+tops?\s+are\s+finished|finished\s+top", s):
                top_hdr = j if top_hdr is None else top_hdr
            elif re.search(r"\bunfinished\b", s) and not re.search(r"\bfinished\b", s):
                unfinished_hdr.add(j)
    if stain_hdr is None and paint_hdr is None:
        return None
    # Wood header cols: labeled unfinished group, or columns left of finish cost under wood names
    finish_hdrs = {j for j in (stain_hdr, paint_hdr, top_hdr) if j is not None}
    wood_hdrs = sorted(j for j in unfinished_hdr if j not in finish_hdrs)
    if not wood_hdrs:
        # Occasionals: wood specie names sit under UNFINISHED without the word unfinished
        first_finish = min(finish_hdrs)
        for h in header_block:
            # Skip note / prose rows that happen to mention a wood species
            joined = " ".join(
                c.strip() for c in h if isinstance(c, str) and c.strip()
            ).lower()
            if re.search(r"\bnote\b|priced with|call us|see page|optional", joined):
                continue
            for j in range(min(first_finish, len(h))):
                c = h[j]
                if not isinstance(c, str):
                    continue
                s = c.lower().strip()
                if not s or len(s) > 80:
                    continue
                if s in {"item #", "item#", "description", "markup", "mark up"}:
                    continue
                if re.search(
                    r"column\s*#|unfinished|finish|email|cost\s+to|priced with|bar stools",
                    s,
                ):
                    continue
                if re.search(
                    r"\b(?:oak|maple|hickory|cherry|elm|qswo|walnut|pine)\b", s
                ):
                    wood_hdrs.append(j)
        wood_hdrs = sorted(set(wood_hdrs))
    if len(wood_hdrs) < 2:
        return None
    # Header cols may have gaps (Kitchen island top at col 8 after empty col 7).
    # Map by sorted header order onto the trailing numeric block.
    ordered_hdrs = sorted(set(wood_hdrs) | finish_hdrs)
    for row in rows[start : start + 80]:
        nums = [j for j, c in enumerate(row) if _is_wholesale_price(c)]
        trailing = [j for j in nums if j > max(ordered_hdrs)]
        if len(trailing) < len(wood_hdrs):
            continue
        # Pair as many trailing bases as we have header slots (top may be absent)
        nmap = min(len(trailing), len(ordered_hdrs))
        mapped = {ordered_hdrs[i]: trailing[i] for i in range(nmap)}
        wood_cols = [mapped[j] for j in wood_hdrs if j in mapped]
        if len(wood_cols) < len(wood_hdrs):
            continue
        if not all(j < len(row) and _is_wholesale_price(row[j]) for j in wood_cols):
            continue
        stain_col = mapped.get(stain_hdr) if stain_hdr is not None else None
        paint_col = mapped.get(paint_hdr) if paint_hdr is not None else None
        top_col = mapped.get(top_hdr) if top_hdr is not None else None
        return FinishCostLayout(
            wood_cols=wood_cols,
            stain_col=stain_col,
            paint_col=paint_col,
            top_col=top_col,
        )
    return None


def _sealed_or_base_headers(
    header_rows: list[list], width: int
) -> tuple[int | None, int | None]:
    """Hershy Cyp/Treated: Sealed/Painted preferred, else Price Ea."""
    sealed: int | None = None
    base: int | None = None
    for h in header_rows:
        for j in range(min(width + 1, len(h))):
            c = h[j]
            if not isinstance(c, str):
                continue
            s = c.lower().strip()
            if SEALED_PAINTED_RE.search(s):
                sealed = j
            elif re.fullmatch(r"price\s*ea\.?", s):
                base = j
    return sealed, base


_HEADER_HINT_RE = re.compile(
    r"^(?:item\s*#|item\s*number|model\s*#|model\s*no\.?|description|"
    r"wholesale\s*price|sealed\s*/\s*painted|price\s*ea\.?|"
    r"sug\.?\s*retail|suggested\s*retail|unf\.?|fin\.?|stained)$|"
    r"\bunfinished\b|\bfinished\s*prices?|\bcost\s+to\s+(?:stain|paint)\s+finish|"
    r"\bwoodgrain|\bbrights\b|\bstained\b",
    re.I,
)


def _row_looks_like_column_header(row: list) -> bool:
    """True for grid header rows, not instructional prose that mentions 'finished'."""
    hits = 0
    for c in row:
        if not isinstance(c, str):
            continue
        s = c.strip()
        if not s or len(s) > 80:
            continue
        if _HEADER_HINT_RE.search(s):
            hits += 1
    return hits >= 2 or (
        hits >= 1
        and any(
            isinstance(c, str)
            and re.search(r"^(?:item\s*#|model\s*#|model\s*no\.?)$", c.strip(), re.I)
            for c in row
        )
    )


def _header_block_near(rows: list[list], start: int) -> tuple[list[list], int]:
    header_start = None
    # Prefer a real header above the current row (skip the current data line).
    for back in range(1, min(start, 40) + 1):
        idx = start - back
        row = rows[idx] if idx < len(rows) else []
        if _row_looks_like_column_header(row):
            header_start = idx
            break
    if header_start is None:
        # Fall back: UNFINISHED / FINISH band rows (Heritage Occasionals)
        for back in range(1, min(start, 40) + 1):
            idx = start - back
            row = rows[idx] if idx < len(rows) else []
            joined = " ".join(
                c.strip() for c in row if isinstance(c, str) and c.strip()
            ).lower()
            if "unfinished" in joined and "finish" in joined:
                header_start = idx
                break
    if header_start is None:
        header_start = max(0, start - 1)
    header_block = rows[header_start : header_start + 12]
    width = max((len(r) for r in rows[header_start : header_start + 40]), default=0)
    return header_block, width


def _transform_sealed_or_base_row(
    row: list,
    sealed_col: int | None,
    base_col: int | None,
    multiplier: float,
    *,
    spots: list[SpotSample],
    item_hint: str,
) -> list[str]:
    label, _ = _row_item_label(row)
    if not label:
        label = item_hint
    if not label or re.search(r"\bapprox", label, re.I):
        return []
    if label.strip().startswith("•"):
        return []
    lab = sanitize_retail_text(label, multiplier)
    if not lab:
        return []
    wholesale = None
    if sealed_col is not None and sealed_col < len(row) and _is_wholesale_price(row[sealed_col]):
        wholesale = float(row[sealed_col])
    elif base_col is not None and base_col < len(row) and _is_wholesale_price(row[base_col]):
        wholesale = float(row[base_col])
    if wholesale is None:
        return []
    retail = retail_price_cell(wholesale, multiplier)
    if len(spots) < 3:
        spots.append(SpotSample(item=label[:80], wholesale=wholesale, retail=retail))
    return [lab, retail]


def _is_option_note_row(row: list, label: str) -> bool:
    """Hermies bullet standards / add-on notes that sit beside price-grid columns."""
    if not label:
        return False
    if label.strip().startswith("•"):
        return True
    if re.match(r"^(?:standard|optional|for\s+\d)", label.strip(), re.I):
        return True
    joined = " ".join(_cell_str(c) for c in row if _cell_str(c) and not _is_wholesale_price(c))
    if re.search(r"\badd\b.*\$|\-\s*add\b", joined, re.I) and not ITEM_CODE_RE.match(label):
        return True
    return False


def _is_options_column_header_row(row: list) -> bool:
    """Hogback-style 'Standard Features / Options / Add' header with no values."""
    texts = [
        re.sub(r":\s*$", "", str(c).strip().lower())
        for c in row
        if isinstance(c, str) and str(c).strip()
    ]
    if not texts:
        return False
    if any(_is_wholesale_price(c) or _is_percent_token(c) for c in row):
        return False
    allowed = {"standard features", "options", "add", "features", "option"}
    return all(t in allowed for t in texts) and any(t == "add" for t in texts)


def _row_has_add_label(row: list) -> bool:
    return any(
        isinstance(c, str) and re.search(r"\badds?\b", c, re.I) for c in row if c not in (None, "")
    )


def _row_has_embedded_dollar_text(row: list) -> bool:
    """True when a string cell already carries $ amounts (Hillside seat notes)."""
    return any(isinstance(c, str) and "$" in c for c in row if c not in (None, ""))


def _add_label_texts(row: list) -> list[str]:
    return [
        c.strip()
        for c in row
        if isinstance(c, str) and c.strip() and re.search(r"\badds?\b", c, re.I)
    ]


def _should_use_add_upcharge_transform(row: list) -> bool:
    """Route only sparse percent/dollar Add upcharges — not prose notes or wood grids."""
    if not _row_has_add_label(row):
        return False
    # Multi-line notes already contain wholesale $ text; leave for sanitize/prose.
    if _row_has_embedded_dollar_text(row):
        return False
    has_pct = any(
        _is_percent_token(c)
        or (
            isinstance(c, (int, float))
            and not isinstance(c, bool)
            and 0 < float(c) < 1
        )
        for c in row
    )
    wholesales = [c for c in row if _is_wholesale_price(c)]
    if has_pct:
        return True
    if not wholesales:
        return False
    labels = _add_label_texts(row)
    # Hillside-style mirrored two-panel rows duplicate the same Add label.
    if len(labels) >= 2 and len({re.sub(r"\s+", " ", x) for x in labels}) == 1:
        return False
    # Wide wood-price grids (Hardwood XX's across species) keep normal column layout.
    if len(wholesales) >= 3:
        return False
    return True


def _transform_add_upcharge_row(
    row: list,
    multiplier: float,
    *,
    spots: list[SpotSample],
    item_hint: str,
) -> list[str]:
    """Keep values beside Add labels; percent tokens are never multiplied."""
    label_parts: list[str] = []
    values: list[str] = []
    seen_labels: set[str] = set()
    for c in row:
        if c in (None, ""):
            continue
        if _is_percent_token(c):
            values.append(str(c).strip())
            continue
        if isinstance(c, (int, float)) and not isinstance(c, bool):
            v = float(c)
            if 0 < v < 1:
                # Fraction without a captured format — still show as percent near Add
                values.append(excel_percent_to_display(v))
            elif _is_wholesale_price(c):
                retail = retail_price_cell(c, multiplier)
                values.append(retail)
                if len(spots) < 3:
                    spots.append(
                        SpotSample(
                            item=(item_hint or "add")[:80],
                            wholesale=v,
                            retail=retail,
                        )
                    )
            continue
        if isinstance(c, str):
            t = sanitize_retail_text(c, multiplier)
            if not t:
                continue
            key = re.sub(r"\s+", " ", t).casefold()
            if key in seen_labels:
                continue
            seen_labels.add(key)
            label_parts.append(t)
    if not label_parts or not values:
        return []
    label = clean_pdf_plain_text(" ".join(label_parts))
    label = _strip_contact_pricing_phrases(label)
    if not label:
        return []
    joined_vals = " ".join(values)
    if re.search(r"\badd\s*$", label, re.I):
        return [f"{label} {joined_vals}"]
    return [f"{label} Add {joined_vals}"]


def _transform_finish_cost_row(
    row: list,
    layout: FinishCostLayout,
    multiplier: float,
    *,
    spots: list[SpotSample],
    item_hint: str,
) -> list[str]:
    label, _ = _row_item_label(row)
    if not label:
        label = item_hint
    if not label or re.search(r"\bapprox", label, re.I):
        return []
    lab = sanitize_retail_text(label, multiplier)
    if not lab:
        return []
    finish_add = 0.0
    stain_v = None
    if layout.stain_col is not None and layout.stain_col < len(row):
        v = row[layout.stain_col]
        if _is_wholesale_price(v):
            stain_v = float(v)
            finish_add = stain_v
    if finish_add <= 0 and layout.paint_col is not None and layout.paint_col < len(row):
        v = row[layout.paint_col]
        if _is_wholesale_price(v):
            finish_add = float(v)
    top_add = 0.0
    if layout.top_col is not None and layout.top_col < len(row):
        v = row[layout.top_col]
        if _is_wholesale_price(v):
            top_add = float(v)
    cells = [lab]
    got = False
    for j in layout.wood_cols:
        if j >= len(row) or not _is_wholesale_price(row[j]):
            cells.append("")
            continue
        wholesale = float(row[j]) + finish_add + top_add
        retail = retail_price_cell(wholesale, multiplier)
        cells.append(retail)
        got = True
        if len(spots) < 3:
            spots.append(
                SpotSample(item=label[:80] or f"col{j}", wholesale=wholesale, retail=retail)
            )
    while cells and not cells[-1]:
        cells.pop()
    return cells if got else []


def _item_code_cols(header_rows: list[list], width: int) -> set[int]:
    cols: set[int] = set()
    for h in header_rows:
        for j in range(min(width + 1, len(h))):
            c = h[j]
            if not isinstance(c, str):
                continue
            s = c.lower().strip()
            if re.search(
                r"^(?:item\s*#|item\s*no\.?|item\s*number|model\s*#|model\s*no\.?|sku|order\s*#)$",
                s,
            ):
                cols.add(j)
    if not cols and width >= 1:
        # Default: first column is often the item code when a description follows
        cols.add(0)
    return cols

def _row_item_label(row: list) -> tuple[str, int]:
    for j, c in enumerate(row[:6]):
        if _is_wholesale_price(c) and j > 0:
            # Don't treat mid-row wholesale amounts as the item label
            continue
        if isinstance(c, (int, float)) and not isinstance(c, bool):
            v = float(c)
            if v == int(v) and 1 <= v <= 999999 and j == 0:
                return str(int(v)), j
        s = _cell_str(c)
        if not s:
            continue
        if re.fullmatch(r"\d+(?:\.\d+)?", s) and j > 0:
            continue
        if ITEM_CODE_RE.match(s) or (len(s) > 8 and re.search(r"\d", s)):
            return s, j
        if j <= 1 and len(s) > 12 and not _is_wholesale_price(c):
            return s, j
        # Hillside-style: "Side Chair" / "Arm Chair" in a mid label column
        if (
            j <= 3
            and len(s) >= 4
            and re.search(r"[A-Za-z]", s)
            and not re.fullmatch(r"unf\.?|fin\.?|stained", s, re.I)
            and not _is_wholesale_price(c)
        ):
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
    if re.search(r"\bapprox", label, re.I):
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
        if _is_na_cell(v):
            cells.append("")
            continue
        if isinstance(v, str) and v.strip().lower() in {"n/a", "na", "-", "—", "–"}:
            cells.append("")
            continue
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
    # Drop n/a placeholders left as blanks between size/wood prices
    return [cells[0]] + [c for c in cells[1:] if c]


def _labeled_cols_have_prices(
    rows: list[list], start: int, labeled: set[int]
) -> bool:
    """True when most priced rows actually store values in the labeled columns."""
    hits = misses = 0
    for row in rows[start : start + 120]:
        if not any(_is_wholesale_price(c) for c in row):
            continue
        if any(j < len(row) and _is_wholesale_price(row[j]) for j in labeled):
            hits += 1
        else:
            misses += 1
    return hits > 0 and hits >= misses


def _mirror_finished_cols(
    rows: list[list], start: int, header_block: list[list], width: int
) -> set[int]:
    """Map finished headers to trailing wholesale bases (Fighting Creek markup=0)."""
    unfinished: set[int] = set()
    finished: set[int] = set()
    for h in header_block:
        for j in range(min(width + 1, len(h))):
            c = h[j]
            if not isinstance(c, str):
                continue
            s = c.lower()
            if re.search(r"\bunfinished\d*\b", s) and not re.search(r"\bfinished\b", s):
                unfinished.add(j)
            elif re.search(r"\bfinished\b", s) or re.search(
                r"\b(?:stain|glaz(?:e|ing)|2[\s-]?tone|paint)\b", s
            ):
                finished.add(j)
    labeled = finished | unfinished
    if not labeled:
        return set()
    first_hdr = min(labeled)
    for row in rows[start : start + 80]:
        nums = [j for j, c in enumerate(row) if _is_wholesale_price(c)]
        if len(nums) < len(labeled):
            continue
        # Trailing block after the header-labeled region
        trailing = [j for j in nums if j > max(labeled)]
        if len(trailing) < len(labeled):
            continue
        offset = trailing[0] - first_hdr
        if offset <= 0:
            continue
        mirrored = {j + offset for j in finished}
        if any(
            j < len(row) and _is_wholesale_price(row[j])
            for j in mirrored
            for row in rows[start : start + 40]
        ):
            return mirrored
    return set()


def _infer_price_cols(rows: list[list], start: int) -> set[int]:
    cols: set[int] = set()
    # Prefer a nearby header row above the current line when present
    header_block, width = _header_block_near(rows, start)
    item_cols = _item_code_cols(header_block, width)
    excluded = _excluded_price_header_cols(header_block, width)
    labeled = _finished_price_cols(header_block, width) - excluded
    # Prefer finished/stained columns that actually hold values (Hilltop mirror left empty)
    if labeled:
        populated = {
            j
            for j in labeled
            if any(
                j < len(row) and _is_wholesale_price(row[j])
                for row in rows[start : start + 120]
            )
        }
        if populated and _labeled_cols_have_prices(rows, start, populated):
            return populated - item_cols - excluded
    if labeled and _labeled_cols_have_prices(rows, start, labeled):
        return labeled - item_cols - excluded
    if labeled:
        mirrored = _mirror_finished_cols(rows, start, header_block, width)
        if mirrored:
            return mirrored - item_cols - excluded
    # Hershy Aurora/Eclipse: prefer wholesale + woodgrain, drop suggested retail
    wholesale_hdr: set[int] = set()
    woodgrain_hdr: set[int] = set()
    for h in header_block:
        for j, c in enumerate(h):
            if not isinstance(c, str):
                continue
            s = c.lower().strip()
            if re.search(r"wholesale\s*price", s):
                wholesale_hdr.add(j)
            if re.search(r"woodgrain|brights", s):
                woodgrain_hdr.add(j)
    if wholesale_hdr:
        prefer = (wholesale_hdr | woodgrain_hdr) - excluded - item_cols
        if prefer and _labeled_cols_have_prices(rows, start, prefer):
            return prefer
    for row in rows[start : start + 200]:
        for j, c in enumerate(row):
            if j in item_cols or j in excluded:
                continue
            if _is_wholesale_price(c):
                cols.add(j)
    # Bamm's: prefer right half only when the left region has no prices on nearby rows
    # (do not use hits/misses — option-add columns elsewhere can skew that check).
    if cols and max(cols) >= 4:
        left = {j for j in cols if j < 4}
        right = {j for j in cols if j >= 4}
        left_has = any(
            j < len(row) and _is_wholesale_price(row[j])
            for row in rows[start : start + 40]
            for j in left
        )
        if right and left and not left_has:
            cols = right
        elif right and not left:
            cols = right
    if not cols:
        return set()
    # Drop unfinished-only columns when paired with an unfinished header
    bad: set[int] = set(item_cols) | excluded
    for h in header_block:
        for j, c in enumerate(h):
            if not isinstance(c, str):
                continue
            s = c.strip().lower()
            if s == "unfinished" or (
                re.search(r"\bunfinished\b", s) and not re.search(r"\bfinished\b", s)
            ):
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
    if clean_name.strip().lower() in ("terms", "markup", "cover", "backcover") or re.search(
        r"\bmark-?up\b", clean_name, re.I
    ):
        pass
    elif clean_name.strip().lower() not in ("wholesale", "finished wholesale"):
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
        if re.search(r"how\s+to\s+price\s+beverage|take\s+the\s+base\s+price\s+in\s+column", joined, re.I):
            i += 1
            continue
        if _is_options_column_header_row(row):
            i += 1
            continue
        if _should_use_add_upcharge_transform(row):
            cells = _transform_add_upcharge_row(
                row, multiplier, spots=spots, item_hint=item_hint
            )
            if cells:
                out.append(ProcessedRow(RowKind.PROSE, cells[:4]))
                i += 1
                continue
        # Pure Add label with no adjacent value and no embedded $ — omit so the
        # PDF never ends a line with a bare "Add".
        if (
            _row_has_add_label(row)
            and not _row_has_embedded_dollar_text(row)
            and not any(
                _is_percent_token(c) or _is_wholesale_price(c) for c in row
            )
            and not any(
                isinstance(c, (int, float))
                and not isinstance(c, bool)
                and 0 < float(c) < 1
                for c in row
            )
        ):
            i += 1
            continue
        label_early, _ = _row_item_label(row)
        # Option notes are prose with embedded $ text. Rows that also carry
        # numeric wholesale cells (Hardwood XX Add grids) must keep the price path.
        if _is_option_note_row(row, label_early) and not any(
            _is_wholesale_price(c) for c in row
        ):
            # Keep each note cell (Hillside two-panel books); do not join+truncate
            # into a single 200-char blob that drops seat/Double-X lines.
            note_cells: list[str] = []
            seen_notes: set[str] = set()
            for c in row:
                if not isinstance(c, str) or not c.strip():
                    continue
                prose = _strip_contact_pricing_phrases(
                    sanitize_retail_text(c, multiplier)
                )
                if not prose or len(prose) <= 12 or is_dealer_terms_text(prose):
                    continue
                key = re.sub(r"\s+", " ", prose).casefold()
                if key in seen_notes:
                    continue
                seen_notes.add(key)
                note_cells.append(prose)
            if note_cells:
                out.append(ProcessedRow(RowKind.PROSE, note_cells[:4]))
            i += 1
            continue
        finish_layout = _detect_finish_cost_layout(rows, i)
        if finish_layout and any(_is_wholesale_price(c) for c in row):
            cells = _transform_finish_cost_row(
                row, finish_layout, multiplier, spots=spots, item_hint=item_hint
            )
            if cells:
                out.append(ProcessedRow(RowKind.WIDE, cells))
                i += 1
                continue
        header_block, width = _header_block_near(rows, i)
        sealed_col, base_col = _sealed_or_base_headers(header_block, width)
        if sealed_col is not None:
            cells = _transform_sealed_or_base_row(
                row,
                sealed_col,
                base_col,
                multiplier,
                spots=spots,
                item_hint=item_hint,
            )
            if cells:
                out.append(ProcessedRow(RowKind.WIDE, cells))
                i += 1
                continue
        price_cols = _infer_price_cols(rows, i)
        cells = _transform_row_prices(
            row, multiplier, price_cols or range(len(row)), spots=spots, item_hint=item_hint
        )
        if not cells and any(_is_wholesale_price(c) for c in row):
            # Inferred wood cols can miss Add-upcharge rows (Hillside Hardwood XX
            # mirrored panels). Fall back to actual wholesale positions; prefer the
            # left panel when the same Add label is duplicated.
            nums = [j for j, c in enumerate(row) if _is_wholesale_price(c)]
            add_idxs = [
                j
                for j, c in enumerate(row)
                if isinstance(c, str) and re.search(r"\badds?\b", c, re.I)
            ]
            if len(add_idxs) >= 2:
                left = [j for j in nums if j < add_idxs[1]]
                if left:
                    nums = left
            if nums:
                cells = _transform_row_prices(
                    row, multiplier, nums, spots=spots, item_hint=item_hint
                )
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
        prose_before_len = 0
        for c in row:
            if _is_wholesale_price(c):
                continue
            raw = _cell_str(c)
            if re.search(r"\bapprox", raw, re.I):
                continue
            cleaned = sanitize_retail_text(raw, multiplier)
            if cleaned:
                prose_before_len += len(cleaned)
            t = _strip_contact_pricing_phrases(cleaned)
            if t and len(t) > 3:
                prose_cells.append(t)
        joined_prose = " ".join(prose_cells)
        # Keep normal long prose, or short leftovers after contact-phrase strip
        # (e.g. 'Only offered in 60" width.').
        keep_prose = len(joined_prose) > 40 or (
            prose_before_len > len(joined_prose) and len(joined_prose) > 12
        )
        if (
            prose_cells
            and keep_prose
            and not any(is_builder_contact_text(c) for c in prose_cells)
            and not any(is_dealer_terms_text(c) for c in prose_cells)
        ):
            out.append(ProcessedRow(RowKind.PROSE, prose_cells[:4]))
        i += 1
    return out


def _unfinished_only_sheet(rows: list[list]) -> bool:
    for row in rows[:12]:
        joined = " ".join(_cell_str(c) for c in row if _cell_str(c)).lower()
        if re.search(r"all\s+prices\s+(are\s+)?given\s+as\s+unfinished", joined):
            return True
        if re.search(r"unfinished\s+prices?\s+only", joined):
            return True
        # Heirloom Reflections: "Wholesale/Retail UNFINISHED Price List"
        if re.search(
            r"(?:wholesale|retail)\s+unfinished\s+price\s+list|"
            r"unfinished\s+price\s+list",
            joined,
        ):
            return True
    return False


def _countryside_ambiguous_finished(rows: list[list]) -> bool:
    """True when header grid mixes finished/unfinished without clear finished columns."""
    unf = fin = False
    for row in rows[:40]:
        for c in row:
            if not isinstance(c, str):
                continue
            s = c.lower()
            if re.search(r"\bunfinished\b", s):
                unf = True
            if re.search(r"\bfinished\b", s) and "unfinished" not in s:
                fin = True
    return unf and not fin


def process_countryside_rustic_workbook(
    path: Path,
    *,
    multiplier: float = MULTIPLIER_DEFAULT,
) -> ViztechBuildResult:
    sheets = load_workbook_sheets(path)
    spots: list[SpotSample] = []
    skipped: list[str] = []
    out: list[ProcessedRow] = []
    has_grid = False

    for name, rows in sheets.items():
        low = name.casefold()
        if "retail" in low:
            skipped.append(name)
            continue
        if "markup" in low and "wholesale" not in low:
            skipped.append(name)
            continue
        if not rows or not any(any(c not in (None, "") for c in r) for r in rows):
            skipped.append(name)
            continue
        if _countryside_ambiguous_finished(rows):
            return ViztechBuildResult(
                [],
                spots,
                skipped + [name],
                ["Finished vs unfinished not distinguishable"],
            )
        before = len(out)
        out.extend(_process_sheet(name, rows, multiplier=multiplier, spots=spots))
        if len(out) > before:
            has_grid = True

    if not has_grid:
        return ViztechBuildResult([], spots, skipped, [f"No wholesale grid in {path.name}"])

    return ViztechBuildResult(collapse_blank_rows(out), spots, skipped, [])


def luxhome_has_finished_products(path: Path) -> bool:
    sheets = load_workbook_sheets(path)
    for name, rows in sheets.items():
        if SKIP_SHEET_RE.search(name) or UNFINISHED_SHEET_RE.search(name):
            continue
        for row in rows[:200]:
            if any(_is_wholesale_price(c) for c in row):
                return True
    return False


def _prefer_wholesale_sheet_names(
    sheets: dict[str, list[list]],
) -> set[str]:
    """When both Wholesale and a markup Price List exist, keep wholesale only."""
    names = list(sheets.keys())
    lower = {n: n.casefold() for n in names}
    has_wholesale = any(re.search(r"\bwholesale\b", v) for v in lower.values())
    skip: set[str] = set()
    if has_wholesale:
        for n, v in lower.items():
            if re.search(r"\bwholesale\b", v):
                continue
            if re.search(r"price\s*list", v) and not re.search(r"\bwholesale\b", v):
                skip.add(n)
    # Slab workbooks: Prices is often formula-only (empty data_only); Master holds wholesale
    if any(v == "prices" for v in lower.values()) and any(v == "master" for v in lower.values()):
        prices_name = next(n for n, v in lower.items() if v == "prices")
        master_name = next(n for n, v in lower.items() if v == "master")
        prices_has = any(
            _is_wholesale_price(c) for row in sheets[prices_name] for c in row[1:]
        )
        master_has = any(
            _is_wholesale_price(c) for row in sheets[master_name] for c in row[1:]
        )
        if master_has and not prices_has:
            skip.add(prices_name)
        else:
            skip.add(master_name)
    # When Master holds a real priced grid, skip duplicate Sheet*/Pricing/Title Page.
    # Hillside-style books use Master as markup-only and keep Sheet3+ price grids.
    if any(v == "master" for v in lower.values()):
        master_name = next(n for n, v in lower.items() if v == "master")
        master_prices = sum(
            1
            for row in sheets[master_name]
            for c in row[1:]
            if _is_wholesale_price(c)
        )
        if master_prices >= 10:
            for n, v in lower.items():
                if v == "master" or v == "markup":
                    continue
                if (
                    v.startswith("sheet")
                    or v == "title page"
                    or "product size" in v
                    or re.search(r"^t\s*price|^cy\s*price", v)
                ):
                    skip.add(n)
                if v == "pricing":
                    pricing_has = any(
                        _is_wholesale_price(c)
                        for row in sheets[n]
                        for c in row[1:]
                    )
                    if not pricing_has:
                        skip.add(n)
        else:
            # Markup-only Master — do not treat Sheet* as duplicates
            skip.add(master_name)
    # Drop Internet / Drop Ship shipping charts
    for n, v in lower.items():
        if re.search(r"internet\s*pricing|drop\s*ship", v):
            skip.add(n)
    return skip


def process_viztech_workbook(
    path: Path,
    *,
    multiplier: float = MULTIPLIER_DEFAULT,
    builder_name: str | None = None,
) -> ViztechBuildResult:
    if re.search(r"markup-calculator", path.name, re.I):
        return process_countryside_rustic_workbook(path, multiplier=multiplier)
    sheets = load_workbook_sheets(path)
    if any(_unfinished_only_sheet(rows) for rows in sheets.values()):
        return ViztechBuildResult(
            [],
            [],
            list(sheets.keys()),
            [f"Unfinished-only wholesale prices in {path.name}"],
        )
    spots: list[SpotSample] = []
    skipped: list[str] = []
    out: list[ProcessedRow] = []
    has_grid = False
    prefer_skip = _prefer_wholesale_sheet_names(sheets)

    for name, rows in sheets.items():
        if SKIP_SHEET_RE.search(name):
            skipped.append(name)
            continue
        if name in prefer_skip:
            skipped.append(name)
            continue
        if UNFINISHED_SHEET_RE.search(name):
            skipped.append(name)
            continue
        if re.search(r"\bretail\b", name, re.I) and not re.search(r"\bwholesale\b", name, re.I):
            skipped.append(name)
            continue
        if not rows or not any(any(c not in (None, "") for c in r) for r in rows):
            skipped.append(name)
            continue
        if _unfinished_only_sheet(rows):
            skipped.append(name)
            continue
        # Hillside: skip INDEX / approximate finishing-cost sheets (page #s / add-ons)
        head = " ".join(
            _cell_str(c)
            for row in rows[:6]
            for c in row
            if _cell_str(c)
        ).lower()
        if re.match(r"^\s*index\b", head) or head.strip().startswith("index"):
            skipped.append(name)
            continue
        if re.search(
            r"approximate\s+(?:unfinished\s+and\s+)?(?:finished\s+)?(?:prices|finishing\s+costs?)",
            head,
        ):
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
    warnings: list[str] = []
    # Hershy Way: Internet list overlaps Cyp/Eclipse and is older (page ~2026-02-16 vs
    # wholesale ~2026-04-15). Prefer the newer wholesale files when both are present.
    if "hershy" in key:
        wholesale = [
            p
            for p in paths
            if not re.search(r"internet", p.name, re.I)
        ]
        internet = [p for p in paths if re.search(r"internet", p.name, re.I)]
        if wholesale and internet:
            for p in internet:
                warnings.append(
                    f"Skipped older overlapping Internet list ({p.name}); "
                    "using newer wholesale lists"
                )
            paths = wholesale
    # Keep newest MMYY per section key
    by_section: dict[str, Path] = {}
    for p in sorted(paths, key=file_mmyy):
        sec_key = section_dedupe_key(p)
        prev = by_section.get(sec_key)
        if prev is None or file_mmyy(p) >= file_mmyy(prev):
            by_section[sec_key] = p

    title = f"{builder_name} Retail Price List"
    all_rows: list[ProcessedRow] = [ProcessedRow(RowKind.TITLE, [title])]
    spots: list[SpotSample] = []
    skipped: list[str] = []

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

    data_rows = [r for r in all_rows if r.kind not in (RowKind.TITLE, RowKind.SECTION)]
    if len(all_rows) <= 1 or not data_rows:
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
