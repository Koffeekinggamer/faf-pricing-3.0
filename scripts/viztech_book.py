"""Process Viztech-exported dealer workbooks into retail ProcessedRows."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import openpyxl

from scripts.dealer_book import ProcessedRow, RowKind, collapse_blank_rows, row_is_blank
from scripts.dealer_terms import (
    PHONE_NUMBER_RE,
    is_builder_contact_text,
    is_dealer_terms_text,
)
from scripts.pdf_plain_text import clean_pdf_plain_text
from scripts.retail_text import (
    excel_percent_to_display,
    forbidden_phrases,
    is_call_for_pricing_cell,
    is_percent_token,
    retail_price_cell,
    sanitize_retail_text,
    strip_call_for_contact_phrases,
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
    r"^2026\s*retail$|retail\s*price\s*list\s*markup|^increase\s*entry$|"
    r"^personalization$|^contents$|^about\s*us$|^info$|dealer\s*notes|"
    r"front\s*page|^\s*percentage\s*$)",
    re.I,
)
# Finish upcharge columns (Jericho "Finished Prices Add", Kohler "Add for 2-Tone").
# Do NOT match a bare "Add" label — Hillside Options rows use that and sit inside
# the header look-ahead window, which would drop the first wood/Add price column.
ADD_ON_PRICE_COL_RE = re.compile(
    r"finished\s*prices?\s*add|"
    r"add\s+for\b|"
    r"\badds?\s+for\b",
    re.I,
)
UNFINISHED_SHEET_RE = re.compile(r"unfinished|retail\s*unf", re.I)
ITEM_CODE_RE = re.compile(
    # Require a digit after "N-" so prose like "2-tone Stain" is not an item code.
    r"^(?:[A-Z]{1,4}[-\d]|[\d]{3,5}[\s-][A-Z]{2}|#\d|M\d|MI\d|FV\d|[\d]+-\d)",
    re.I,
)
MMYY_RE = re.compile(r"[_\s](\d{2})(\d{2})(?:-\d+)?\.(?:xlsx|xls)", re.I)
FINISH_COST_RE = re.compile(
    r"cost\s+to\s+(?:stain|paint)\s+finish|finishing\s+costs?",
    re.I,
)
RETAIL_MARKUP_COL_RE = re.compile(
    r"sug\.?\s*retail|suggested\s*retail|\bmap\b|\bmsrp\b|"
    r"internet\s*prices?|pallet\s*qtys?|pallet\s*price",
    re.I,
)
SEALED_PAINTED_RE = re.compile(r"sealed\s*/\s*painted|sealed\s+and\s+painted", re.I)
def _strip_contact_pricing_phrases(text: str) -> str:
    """Remove 'contact us' / 'call for pricing' clauses from notes; keep other wording."""
    return strip_call_for_contact_phrases(text)


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


def _section_label_from_slug_suffix(path: Path) -> str | None:
    """builder-slug__Collection_Name.xlsx → 'Collection Name' (Kidron-style)."""
    m = re.match(r"^[^_]+(?:-[^_]+)*__(.+)\.(?:xlsx|xls)$", path.name, re.I)
    if not m:
        return None
    raw = m.group(1).strip()
    # Drop leading catalog codes like CODE_2026_ when present; keep plain collection names.
    coded = re.match(r"^[A-Z]{1,6}_\d{4}_(.+)$", raw, re.I)
    if coded:
        return None  # handled by _section_label_from_upload_name
    if re.search(r"_Pricelist|_Price_List|_Price-List", raw, re.I):
        return None
    label = raw.replace("_", " ").replace("-", " ")
    label = re.sub(r"\s+", " ", label).strip()
    if not label or label.casefold() in {"pricelist", "price list", "master"}:
        return None
    words: list[str] = []
    for w in label.split():
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
    # Kidron-style slug__Collection_Name.xlsx — after known catalog patterns only.
    from_suffix = _section_label_from_slug_suffix(path)
    if from_suffix:
        return from_suffix
    return "Price list"


def section_dedupe_key(path: Path) -> str:
    """One merge slot per workbook file (L&N main + bookcase share a section title)."""
    return path.stem.casefold()


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


def _is_wholesale_price(val: Any, *, allow_yearish: bool = False) -> bool:
    if isinstance(val, str) and val.strip().upper() in {"#N/A", "N/A", "NA", "#NA"}:
        return False
    if is_percent_token(val):
        return False
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return False
    v = float(val)
    # Calendar years in headers (e.g. INTEG "2026 Pricelist") are not prices.
    # Real wholesale amounts can fall in 2015–2035 (L&N desks, LA Patio, Comfort);
    # pass allow_yearish=True when the cell already sits in a finished price column.
    if (
        not allow_yearish
        and abs(v - round(v)) < 1e-9
        and 2015 <= int(round(v)) <= 2035
    ):
        return False
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


# Jericho Master: =ROUND(D6*$E$3,0) and = (ROUND(D111*$F$3,0))+10
_ROUND_PRODUCT_RE = re.compile(
    r"^=\s*\(?\s*ROUND\s*\(\s*"
    r"(?:(?P<s1>(?:'[^']+'|[^'!]+)!)?(?P<c1>\$?[A-Z]+\$?\d+))"
    r"\s*\*\s*"
    r"(?:(?P<s2>(?:'[^']+'|[^'!]+)!)?(?P<c2>\$?[A-Z]+\$?\d+))"
    r"\s*,\s*(?P<digits>\d+)\s*\)\s*\)?\s*"
    r"(?:(?P<op>[+-])\s*(?P<add>\d+(?:\.\d+)?))?\s*$",
    re.I,
)
# Prices sheet: =ROUND(Master!D111*Markup!$B$6,2)
_MASTER_REF_RE = re.compile(
    r"Master!\$?[A-Z]+\$?(\d+)",
    re.I,
)


def _col_row_from_a1(a1: str) -> tuple[int, int]:
    m = re.match(r"\$?([A-Z]+)\$?(\d+)$", a1.strip(), re.I)
    if not m:
        raise ValueError(a1)
    letters, row_s = m.group(1).upper(), m.group(2)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - 64)
    return col - 1, int(row_s) - 1


def _excel_round_half_up(value: Any, digits: int = 0) -> float:
    """Excel ROUND (half away from zero) — not Python banker's round()."""
    from decimal import Decimal, ROUND_HALF_UP

    quant = Decimal("1").scaleb(-digits) if digits else Decimal("1")
    dec = value if isinstance(value, Decimal) else Decimal(str(value))
    return float(dec.quantize(quant, rounding=ROUND_HALF_UP))


def _excel_round_product(a: float, b: float, digits: int = 0) -> float:
    """ROUND(a*b, digits) with Decimal multiplication (350×1.15 → 403, not 402)."""
    from decimal import Decimal

    return _excel_round_half_up(Decimal(str(a)) * Decimal(str(b)), digits)


def _numeric_cell_value(val: Any) -> float | None:
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val)
    if isinstance(val, str):
        s = val.strip()
        if s.endswith("%"):
            try:
                return float(s[:-1].strip()) / 100.0
            except ValueError:
                return None
        try:
            return float(s.replace(",", ""))
        except ValueError:
            return None
    return None


def _eval_round_product_formulas(
    rows: list[list],
    formulas: list[list],
    *,
    sheet_name: str,
    cross: dict[str, list[list]] | None = None,
) -> None:
    """Fill data_only gaps for Jericho ROUND(base*factor[,]+k) wood columns.

    Uses Excel half-up rounding on Decimal products so listed wholesale matches
    the workbook (e.g. ROUND(350*1.15,0)=403, ROUND(154*1.25,0)=193).
    """
    if not formulas:
        return
    cross = cross or {}

    def _lookup(sheet_pref: str | None, a1: str, local_rows: list[list]) -> float | None:
        sn = (sheet_pref or "").strip().strip("'") or sheet_name
        grid = local_rows if sn.casefold() == sheet_name.casefold() else None
        if grid is None:
            for k, g in cross.items():
                if k.casefold() == sn.casefold():
                    grid = g
                    break
        if grid is None:
            return None
        try:
            cc, rr = _col_row_from_a1(a1)
        except ValueError:
            return None
        if rr < 0 or rr >= len(grid) or cc < 0 or cc >= len(grid[rr]):
            return None
        return _numeric_cell_value(grid[rr][cc])

    changed = True
    passes = 0
    while changed and passes < 12:
        changed = False
        passes += 1
        for r, frow in enumerate(formulas):
            while len(rows) <= r:
                rows.append([])
            row = rows[r]
            for c, formula in enumerate(frow):
                if not isinstance(formula, str) or not formula.startswith("="):
                    continue
                if c < len(row) and row[c] not in (None, ""):
                    continue
                m = _ROUND_PRODUCT_RE.match(formula.strip())
                if not m:
                    continue
                a = _lookup(m.group("s1"), m.group("c1"), rows)
                b = _lookup(m.group("s2"), m.group("c2"), rows)
                if a is None or b is None:
                    continue
                digits = int(m.group("digits"))
                product = _excel_round_product(a, b, digits)
                if m.group("add") is not None:
                    addend = float(m.group("add"))
                    product = product + addend if m.group("op") == "+" else product - addend
                while len(row) <= c:
                    row.append(None)
                row[c] = product
                changed = True


def _join_prices_item_codes_onto_master(
    sheets: dict[str, list[list]],
    formula_map: dict[str, list[list]],
) -> None:
    """Copy Prices-sheet item codes (and descriptions) onto Master rows.

    Town Hill / Modern Shadow Boxes / Economy Bookcases store codes on Prices
    (TH36, MSB305, ECO-3, …) while Master leaves column A blank or only puts the
    code in the description column. Prices formulas reference Master!D{row}.
    """
    master = sheets.get("Master")
    prices = sheets.get("Prices")
    prices_forms = formula_map.get("Prices")
    if not master or not prices or not prices_forms:
        return
    for r, prow in enumerate(prices):
        if not prow:
            continue
        code = prow[0] if isinstance(prow[0], str) else ""
        code = code.strip()
        if not code or not ITEM_CODE_RE.match(code):
            continue
        frow = prices_forms[r] if r < len(prices_forms) else []
        master_row: int | None = None
        for f in frow:
            if not isinstance(f, str):
                continue
            m = _MASTER_REF_RE.search(f)
            if m:
                master_row = int(m.group(1)) - 1
                break
        if master_row is None or master_row < 0 or master_row >= len(master):
            continue
        mrow = master[master_row]
        while len(mrow) < 2:
            mrow.append(None)
        existing = mrow[0]
        if existing in (None, "") or (
            isinstance(existing, str) and not existing.strip()
        ):
            mrow[0] = code
        prices_desc = prow[1].strip() if len(prow) > 1 and isinstance(prow[1], str) else ""
        if prices_desc:
            mdesc = mrow[1] if isinstance(mrow[1], str) else ""
            mdesc_s = mdesc.strip()
            # Replace blank / code-only Master descriptions with Prices text
            # (ECO-3, Mission CD-Single Storage, …).
            if (
                not mdesc_s
                or mdesc_s.casefold() == code.casefold()
                or (
                    mdesc_s.casefold().startswith(code.casefold())
                    and len(mdesc_s) <= len(code) + 2
                )
                or (
                    # Prefer the Prices wording when it is clearly the fuller label
                    len(prices_desc) > len(mdesc_s)
                    or re.search(r"\bsingle\b", prices_desc, re.I)
                )
            ):
                mrow[1] = prices_desc
        # Prefer Prices size when present (ECO-3…6: 13.5" on Prices vs 14" on Master).
        prices_size = (
            prow[2].strip() if len(prow) > 2 and isinstance(prow[2], str) else ""
        )
        if prices_size and re.search(
            r"\d|\"|″|\'|′|[xX×]", prices_size
        ):
            while len(mrow) < 3:
                mrow.append(None)
            mrow[2] = prices_size


def _load_xlsx_sheets(path: Path) -> dict[str, list[list]]:
    """Load values; convert percent-formatted numerics to 'N%' display tokens."""
    wb_val = openpyxl.load_workbook(path, data_only=True, read_only=True)
    wb_fmt = openpyxl.load_workbook(path, data_only=False, read_only=True)
    out: dict[str, list[list]] = {}
    formula_map: dict[str, list[list]] = {}
    for name in wb_val.sheetnames:
        ws_v = wb_val[name]
        ws_f = wb_fmt[name] if name in wb_fmt.sheetnames else None
        rows: list[list] = []
        formulas: list[list] = []
        fmt_rows = ws_f.iter_rows() if ws_f is not None else None
        for r_v in ws_v.iter_rows():
            fmt_row = next(fmt_rows) if fmt_rows is not None else None
            row: list = []
            frow: list = []
            for idx, c_v in enumerate(r_v):
                v = c_v.value
                fmt = ""
                fval = None
                if fmt_row is not None and idx < len(fmt_row):
                    fmt = fmt_row[idx].number_format or ""
                    fval = fmt_row[idx].value
                if (
                    isinstance(v, (int, float))
                    and not isinstance(v, bool)
                    and _excel_format_is_percent(fmt)
                    # Jericho specie factors are 1.15+ with a mistaken 0% format;
                    # only true Excel percent fractions (0–1] become 'N%' tokens.
                    and 0 < float(v) <= 1
                ):
                    row.append(excel_percent_to_display(float(v)))
                else:
                    row.append(v)
                frow.append(fval if isinstance(fval, str) and fval.startswith("=") else None)
            rows.append(row)
            formulas.append(frow)
        out[name] = rows
        formula_map[name] = formulas
    wb_val.close()
    wb_fmt.close()
    # Evaluate ROUND(base*factor) woods (Excel half-up) when Excel left no cache.
    for name, rows in out.items():
        _eval_round_product_formulas(
            rows, formula_map.get(name, []), sheet_name=name, cross=out
        )
    _join_prices_item_codes_onto_master(out, formula_map)
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


def _is_priced_product_row(row: list) -> bool:
    """True when the row looks like a SKU line (item code + wholesale $)."""
    return any(_is_wholesale_price(c) for c in row) and bool(_row_item_label(row)[0])


def _priced_product_survives_contact_false_positive(row: list, probe: str) -> bool:
    """Keep SKU rows that trip contact/dealer heuristics via Phone Holder / PHONE|FAX."""
    if not _is_priced_product_row(row):
        return False
    productish = " ".join(
        _cell_str(c)
        for c in row
        if _cell_str(c)
        and not re.fullmatch(r"phone|fax|tel\.?", _cell_str(c), re.I)
    )
    if not is_builder_contact_text(productish):
        return True
    # Cup/Phone Holder etc. — "phone" in product name with prices, no phone#
    if not PHONE_NUMBER_RE.search(productish):
        if not re.search(r"\b(?:email|@|www\.|https?://)\b", productish, re.I):
            return True
    return False


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
    # Product rows with item code + prices survive contact/dealer false positives
    # (Phone Holder descriptions, PHONE|FAX sidebar cells on chimes sheets).
    if _priced_product_survives_contact_false_positive(row, probe):
        # Still drop real policy/boilerplate that happens to share a code cell
        low_probe = probe.lower()
        if re.search(r"prices?\s+effective\b", low_probe):
            return True
        if re.search(r"to\s+access\s+digital\s+files|go\s+to\s*:", low_probe):
            return True
        if re.search(r"\bdiscontinued\b", low_probe):
            return True
        if any(
            x in low_probe
            for x in (
                "password",
                "viztechfurniture",
                "lagwana.cloud",
                "internet stores",
                "internet price",
                "internet dealer",
            )
        ):
            return True
        return False
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
        r"internet\s+(?:dealer\s+)?policy|brick\s+and\s+mortar|"
        r"\bmark-?up\b.*\bestimated\s+finished\b|"
        r"we\s+will\s+no\s+longer\s+offer\s+finishing\s+pricing",
        probe,
        re.I,
    ):
        return True
    if "wholesale price list" in low or re.match(r"^finished\s+wholesale\s*$", low):
        return True
    if re.search(r"\bdiscontinued\b", low):
        return True
    if re.search(r"\bdealers?\b", low) and re.search(
        r"internet|brick|mortar|mark-?up|promote", low
    ):
        return True
    if re.search(r"prices?\s+effective\b", low):
        return True
    if re.search(r"to\s+access\s+digital\s+files|go\s+to\s*:", low):
        return True
    if re.search(r"\bnet\s*30\b", low):
        return True
    if is_builder_contact_text(probe):
        return True
    return False


def _is_finished_prices_band_label(cell: str) -> bool:
    """Column-band title (L&N ESTIMATED FINISHED PRICE), not KKT prose notes."""
    s = cell.strip()
    if not s or len(s) > 72:
        return False
    if re.search(r"add\s+\d+\s*%|[\u2022\u2023]|^\s*[\u2022\-]\s", s, re.I):
        return False
    if re.search(r"\bto\s+finished\s+price\b", s, re.I):
        return False
    if re.search(r"estimated\s+finish(?:ing)?(?:\s+per\s+piece|\s+prices?)", s, re.I):
        return False
    return bool(
        re.search(
            r"(?:^|\s)(?:estimated\s+)?(?<!un)finished\s+prices?\s*$",
            s,
            re.I,
        )
    )


def _estimated_finished_band_cols(header_rows: list[list], width: int) -> set[int]:
    """L&N-style estimated-finished wood columns (not the unfinished band)."""
    band_starts = [
        j
        for h in header_rows
        for j, c in enumerate(h)
        if isinstance(c, str) and _is_finished_prices_band_label(c)
    ]
    if not band_starts:
        return set()
    start = min(band_starts)
    unfinished: set[int] = set()
    for h in header_rows:
        for j, c in enumerate(h):
            if isinstance(c, str) and re.search(r"\bunfinished\b", c, re.I) and not re.search(
                r"(?<!un)finished", c, re.I
            ):
                if j < start:
                    unfinished.add(j)
    expanded: set[int] = set()
    finish_per_piece: set[int] = set()
    for h in header_rows:
        joined = " ".join(c.strip() for c in h if isinstance(c, str) and c.strip())
        if not joined:
            continue
        if not _HEADER_HINT_RE.search(joined) and not re.search(
            r"estimated\s+finished|unfinished|item\s*#|oak|cherry|maple",
            joined,
            re.I,
        ):
            continue
        if re.search(r"\bdoor\s+add\b|\badds?\s+for\b", joined, re.I):
            continue
        for j in range(start, min(width + 1, len(h))):
            c = h[j]
            if isinstance(c, str) and c.strip():
                if re.search(r"\bunfinished\b", c, re.I) and not re.search(
                    r"(?<!un)finished", c, re.I
                ):
                    break
                if re.search(
                    r"estimated\s+finish(?:ing)?(?:\s+per\s+piece|\s+prices?)",
                    c,
                    re.I,
                ):
                    finish_per_piece.add(j)
                    break
                if re.search(
                    r"^(?:item\s*(?:#|no|number)|description|\bsize\b)$",
                    c.strip(),
                    re.I,
                ):
                    break
                if _is_finished_prices_band_label(c) and not re.search(
                    r"\b(?:oak|cherry|maple|walnut|hickory|elm|qswo)\b", c, re.I
                ):
                    continue
                expanded.add(j)
    expanded -= finish_per_piece
    return expanded - unfinished if expanded else set()


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
            if FINISH_COST_RE.search(s) or ADD_ON_PRICE_COL_RE.search(s):
                # Heritage finish cost / Jericho-Kohler finish Add — not wood prices
                continue
            if re.search(r"finished\s*add\b", s):
                # KKT "Finished ADD" — composed separately, not a wood price
                continue
            if re.search(
                r"estimated\s+finish(?:ing)?(?:\s+per\s+piece|\s+prices?)", s
            ):
                # L&N finish-cost helper column — not a wood price
                continue
            if re.fullmatch(r"unf(?:in)?\.?", s) or (
                re.search(r"\bunfinished\d*\b", s) and not re.search(r"\bfinished\b", s)
            ):
                unfinished.add(j)
            elif re.fullmatch(r"fin\.?", s) or (
                re.search(r"\bfinished\b", s)
                and not ADD_ON_PRICE_COL_RE.search(s)
                and not re.search(r"finished\s*add\b", s)
            ):
                finished.add(j)
            elif SEALED_PAINTED_RE.search(s):
                finished.add(j)
            elif re.search(r"\b(?:standard|wood\s*grain)\b", s) and re.search(
                r"\bpoly\b", s
            ):
                # LA Patio: Standard / Woodgrain / WOOD GRAIN poly finished goods
                finished.add(j)
            elif len(s) <= 48 and re.search(
                r"\b(?:fabric|leather|cowhide)\b", s
            ):
                # Lambright Comfort: Standard Fabric / Ultra Leather grade columns
                finished.add(j)
            elif len(s) <= 48 and re.search(
                r"\b(?:stain(?:ed)?|glaz(?:e|ing)|2[\s-]?tone|paint)\b", s
            ):
                # Fishers / Hilltop / LA Patio column headers — not long stain notes
                finished.add(j)
    # Hoosier / L&N "FINISHED PRICES" / "ESTIMATED FINISHED PRICE" band: expand
    # through following wood headers (exclude Add / finish-per-piece helpers).
    band_starts = [
        j
        for h in header_rows
        for j, c in enumerate(h)
        if isinstance(c, str)
        and _is_finished_prices_band_label(c)
        and not ADD_ON_PRICE_COL_RE.search(c)
        and not re.search(r"finished\s*add\b", c, re.I)
        and not re.search(
            r"estimated\s+finish(?:ing)?(?:\s+per\s+piece|\s+prices?)", c, re.I
        )
    ]
    if band_starts:
        start = min(band_starts)
        expanded: set[int] = set()
        finish_per_piece: set[int] = set()
        for h in header_rows:
            # Only expand across real header rows — skip data/prose that can
            # re-introduce Estimated Finish Per Piece after a Door Add line.
            joined = " ".join(
                c.strip() for c in h if isinstance(c, str) and c.strip()
            )
            if not joined:
                continue
            if not _HEADER_HINT_RE.search(joined) and not re.search(
                r"estimated\s+finished|unfinished|item\s*#|oak|cherry|maple",
                joined,
                re.I,
            ):
                continue
            if re.search(r"\bdoor\s+add\b|\badds?\s+for\b", joined, re.I):
                continue
            for j in range(start, min(width + 1, len(h))):
                c = h[j]
                if isinstance(c, str) and c.strip():
                    if re.search(r"\bunfinished\b", c, re.I) and not re.search(
                        r"(?<!un)finished", c, re.I
                    ):
                        break
                    if re.search(
                        r"estimated\s+finish(?:ing)?(?:\s+per\s+piece|\s+prices?)",
                        c,
                        re.I,
                    ):
                        finish_per_piece.add(j)
                        break
                    if re.search(
                        r"^(?:item\s*(?:#|no|number)|description|\bsize\b)$",
                        c.strip(),
                        re.I,
                    ):
                        break
                    # Skip the band title cell; wood names under it stay eligible.
                    if _is_finished_prices_band_label(c) and not re.search(
                        r"\b(?:oak|cherry|maple|walnut|hickory|elm|qswo)\b", c, re.I
                    ):
                        continue
                    expanded.add(j)
        expanded -= finish_per_piece
        if expanded:
            return expanded - unfinished
    # Fin|Unf or Unf|Fin pairs: keep the Finished column of each pair.
    # Stop before a second Description panel or a gap (KKT mirrored reprint).
    def _is_fin_hdr(s: str) -> bool:
        s = s.lower().strip()
        return bool(
            re.fullmatch(r"fin\.?", s)
            or (
                re.search(r"\bfinished\b", s)
                and not re.search(r"finished\s*add\b", s)
                and not ADD_ON_PRICE_COL_RE.search(s)
            )
        )

    def _is_unf_hdr(s: str) -> bool:
        s = s.lower().strip()
        return bool(
            re.fullmatch(r"unf(?:in)?\.?", s)
            or (
                re.search(r"\bunfinished\b", s) and not re.search(r"\bfinished\b", s)
            )
        )

    pair_finished: set[int] = set()
    for h in header_rows:
        for j in range(min(width, len(h) - 1)):
            a, b = h[j], h[j + 1]
            if not isinstance(a, str) or not isinstance(b, str):
                continue
            if _is_fin_hdr(a) and _is_unf_hdr(b):
                pair_finished.add(j)
            elif _is_unf_hdr(a) and _is_fin_hdr(b):
                pair_finished.add(j + 1)
    if pair_finished:
        # Mirrored reprint panels (same wood labels + same prices) — keep left
        # only. Greenville-style dual panels with *different* wood headers keep
        # both sides (4 finished woods).
        desc_cols = sorted(
            {
                j
                for h in header_rows
                for j, c in enumerate(h)
                if isinstance(c, str)
                and re.fullmatch(r"description", c.strip(), re.I)
            }
        )
        if len(desc_cols) >= 2:
            cut = desc_cols[1]
            left_pf = {j for j in pair_finished if j < cut}
            right_pf = {j for j in pair_finished if j >= cut}
            left_woods = _wood_header_cells(header_rows, before=cut)
            right_woods = _wood_header_cells(header_rows, after=cut)
            # True reprint: identical wood-header cells, empty right labels, or a
            # full left wood set (≥3) with a layout twin (Raleigh). Greenville
            # keeps both sides when left has only 1–2 woods and species *groups*
            # differ (even if maple/cherry tokens appear on both).
            if (
                not right_woods
                or left_woods == right_woods
                or (len(left_pf) >= 3 and right_pf)
                or (len(left_pf) <= 2 and right_pf)
            ):
                pair_finished = left_pf
            # else keep both panels (≥3 distinct finished woods on left + new species)
        sorted_pf = sorted(pair_finished)
        for a, b in zip(sorted_pf, sorted_pf[1:]):
            if b - a > 2:
                left = {j for j in pair_finished if j < b}
                right = {j for j in pair_finished if j >= b}
                left_woods = _wood_header_cells(header_rows, before=b)
                right_woods = _wood_header_cells(header_rows, after=b - 1)
                # Quick Ship / full-set reprint: identical / empty right headers
                # or left already carries a complete wood band.
                if (
                    not right_woods
                    or left_woods == right_woods
                    or (len(left) >= 3 and right)
                ):
                    pair_finished = left
                break
        if pair_finished:
            return pair_finished - unfinished
    if finished:
        return finished - unfinished
    # Barkman-style: skip col1 markup fractions; take numeric cols not in unfinished set
    return set()


def _wood_header_cells(
    header_rows: list[list], *, before: int | None = None, after: int | None = None
) -> set[str]:
    """Normalized full wood-header cell texts in a column range (not token soup)."""
    out: set[str] = set()
    wood_tok = re.compile(
        r"\b(?:oak|maple|cherry|hickory|elm|qswo|qwso|walnut|ash|alder|"
        r"pine|rustic|wormy|brown|hard|sap|white)\b",
        re.I,
    )
    for h in header_rows:
        for j, c in enumerate(h):
            if before is not None and j >= before:
                continue
            if after is not None and j <= after:
                continue
            if not isinstance(c, str) or not c.strip():
                continue
            if not wood_tok.search(c):
                continue
            # Skip Fin/Unf / Description labels that may sit on the same row
            if re.fullmatch(
                r"(?:fin(?:ished)?|unf(?:inished)?|description|item\s*(?:#|no|number))",
                c.strip(),
                re.I,
            ):
                continue
            norm = re.sub(r"\s+", " ", c).strip().casefold()
            out.add(norm)
    return out


def _wood_header_labels(
    header_rows: list[list], *, before: int | None = None, after: int | None = None
) -> set[str]:
    """Normalized wood-species tokens from header cells in a column range."""
    out: set[str] = set()
    for h in header_rows:
        for j, c in enumerate(h):
            if before is not None and j >= before:
                continue
            if after is not None and j <= after:
                continue
            if not isinstance(c, str):
                continue
            for m in re.finditer(
                r"\b(?:oak|maple|cherry|hickory|elm|qswo|walnut|ash|alder|"
                r"pine|rustic|wormy|brown|hard|sap|white)\b",
                c,
                re.I,
            ):
                out.add(m.group(0).casefold())
    return out


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
            if FINISH_COST_RE.search(s) or ADD_ON_PRICE_COL_RE.search(s):
                bad.add(j)
            if re.search(r"finished\s*add\b", s):
                bad.add(j)
            if re.search(
                r"estimated\s+finish(?:ing)?(?:\s+per\s+piece|\s+prices?)", s
            ):
                bad.add(j)
            if re.search(r"\badj\.?\s*shelves?\b|\badjustable\s+shelves?\b", s):
                # L&N bookcase shelf-count column — not a wholesale price
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


@dataclass
class FinishedAddLayout:
    """KKT chairs: unfinished wood bases + shared Finished ADD → finished wholesale."""

    wood_cols: list[int]
    add_col: int


def _detect_finished_add_layout(
    rows: list[list], start: int
) -> FinishedAddLayout | None:
    """Detect unfinished woods + a single Finished ADD column (Kountry Knob chairs).

    Family sections often use a bare ``Item # | ALAMO`` title without repeating
    the wood / Finished ADD headers — walk further up the sheet for those labels.
    """
    family_row = _nearest_family_header_row(rows, start)
    if _family_is_finished_only(rows, family_row):
        return None
    header_block, width = _header_block_near(rows, start)
    item_cols = _item_code_cols(header_block, width)
    # Also scan rows above the local header for the last Finished ADD band.
    scan_rows = list(header_block)
    scan_floor = family_row if family_row is not None else 0
    for back in range(1, min(start - scan_floor, 80) + 1):
        idx = start - back
        if idx < scan_floor:
            break
        row = rows[idx]
        joined = " ".join(
            c.strip() for c in row if isinstance(c, str) and c.strip()
        ).lower()
        if re.search(r"finished\s*add\b", joined) or (
            "unfinished" in joined and re.search(r"\badd\b", joined)
        ):
            if row not in scan_rows:
                scan_rows.insert(0, row)
            # Include the wood-name row just above Unfinished/Add
            if idx > 0 and rows[idx - 1] not in scan_rows:
                scan_rows.insert(0, rows[idx - 1])
            break

    add_cols: list[int] = []
    finished_add_cols: list[int] = []
    unfinished_cols: list[int] = []
    has_finished_add_label = False
    for h in scan_rows:
        for j in range(min(width + 1, len(h))):
            c = h[j]
            if not isinstance(c, str):
                continue
            s = c.lower().strip()
            if re.search(r"finished\s*add\b", s):
                has_finished_add_label = True
                finished_add_cols.append(j)
                add_cols.append(j)
            elif re.fullmatch(r"add", s):
                add_cols.append(j)
            elif re.fullmatch(r"unf(?:in)?\.?", s) or (
                re.search(r"\bunfinished\b", s) and not re.search(r"\bfinished\b", s)
            ):
                unfinished_cols.append(j)
    # MIA-style: Unfinished… + Add + Finished without the words "Finished ADD"
    if not has_finished_add_label:
        if unfinished_cols and add_cols and len(unfinished_cols) >= 2:
            has_finished_add_label = True
        else:
            return None
    add_col = (
        min(finished_add_cols) if finished_add_cols else (min(add_cols) if add_cols else None)
    )
    if add_col is None:
        return None
    woods = sorted({j for j in unfinished_cols if j < add_col})
    if len(woods) < 2:
        # Fall back: wood specie headers left of Add (may have blank first wood)
        woods = []
        for h in scan_rows:
            for j in range(add_col):
                c = h[j] if j < len(h) else None
                if not isinstance(c, str):
                    continue
                if re.search(
                    r"\b(?:oak|maple|cherry|hickory|elm|qswo|walnut|ash)\b", c, re.I
                ):
                    woods.append(j)
        woods = sorted(set(woods))
    # Include blank leading wood slots when Unfinished labels span more cols
    if unfinished_cols:
        u_left = sorted(j for j in unfinished_cols if j < add_col)
        if u_left and (not woods or min(u_left) < min(woods)):
            woods = sorted(set(woods) | set(u_left))
    woods = sorted(j for j in woods if j not in item_cols)
    if len(woods) < 2:
        return None
    # Confirm ADD values and unfinished bases on nearby rows
    add_ok = False
    wood_ok = 0
    for row in rows[start : start + 40]:
        if add_col < len(row) and _is_wholesale_price(row[add_col]):
            add_ok = True
        wood_ok = max(
            wood_ok,
            sum(1 for j in woods if j < len(row) and _is_wholesale_price(row[j])),
        )
        if add_ok and wood_ok >= 2:
            return FinishedAddLayout(wood_cols=woods, add_col=add_col)
    return None


def _transform_finished_add_row(
    row: list,
    layout: FinishedAddLayout,
    multiplier: float,
    *,
    spots: list[SpotSample],
    item_hint: str,
) -> list[str]:
    label, label_j = _row_item_label(row)
    if not label:
        label = item_hint
    if not label or re.search(r"\bapprox", label, re.I):
        return []
    lab = sanitize_retail_text(label, multiplier)
    if not lab:
        return []
    if layout.add_col >= len(row) or not _is_wholesale_price(row[layout.add_col]):
        return []
    add_v = float(row[layout.add_col])
    cells: list[str] = [lab]
    # Keep short description after item code (Side Chair, Arm Chair, …)
    if label_j >= 0:
        for j in range(label_j + 1, min(label_j + 3, layout.wood_cols[0])):
            if j >= len(row) or j == layout.add_col:
                continue
            v = row[j]
            if _is_wholesale_price(v) or _is_na_cell(v):
                continue
            if isinstance(v, str):
                t = sanitize_retail_text(v, multiplier)
                if t and _keep_item_mid_text(t):
                    cells.append(t)
                    break
    got = False
    for j in layout.wood_cols:
        if j >= len(row) or not _is_wholesale_price(row[j]):
            # Blank first-wood slots (LEXI/MIA) — omit, do not emit a dash column
            continue
        wholesale = float(row[j]) + add_v
        retail = retail_price_cell(wholesale, multiplier)
        cells.append(retail)
        got = True
        if len(spots) < 3:
            spots.append(
                SpotSample(item=label[:80] or f"col{j}", wholesale=wholesale, retail=retail)
            )
    return cells if got else []


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
    if len(wood_hdrs) < 2:
        # Occasionals / island tops: wood specie names sit under a single
        # UNFINISHED band label (or none) without per-wood "unfinished" cells.
        first_finish = min(finish_hdrs)
        specie_woods: list[int] = []
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
                    specie_woods.append(j)
        if len(set(specie_woods)) >= 2:
            wood_hdrs = sorted(set(specie_woods))
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
    r"\bwood\s*grain|\bbrights\b|\bstained\b|"
    r"\b(?:fabric|leather|cowhide)\b",
    re.I,
)


def _row_looks_like_column_header(row: list) -> bool:
    """True for grid header rows, not instructional prose that mentions 'finished'."""
    joined = " ".join(
        c.strip() for c in row if isinstance(c, str) and c.strip()
    )
    if re.search(r"\btop\s*size\b", joined, re.I) and not any(
        isinstance(c, str)
        and re.search(r"^item\s*(?:#|number|no\.?)$", c.strip(), re.I)
        for c in row
    ):
        # L&N size-band labels (UNFINISHED / ESTIMATED FINISHED PRICE) — not the wood grid.
        return False
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


def _row_has_fin_unf_pair_labels(row: list) -> bool:
    texts = [
        c.strip().lower()
        for c in row
        if isinstance(c, str) and c.strip()
    ]
    has_fin = any(
        re.fullmatch(r"fin\.?", t)
        or (re.search(r"\bfinished\b", t) and not re.search(r"finished\s*add\b", t))
        for t in texts
    )
    has_unf = any(
        re.fullmatch(r"unf(?:in)?\.?", t)
        or (re.search(r"\bunfinished\b", t) and not re.search(r"\bfinished\b", t))
        for t in texts
    )
    return has_fin and has_unf


def _nearest_family_header_row(rows: list[list], start: int) -> int | None:
    """Row index of the nearest ``Item #`` family title above ``start``."""
    for back in range(1, min(start, 120) + 1):
        idx = start - back
        row = rows[idx]
        if not row:
            continue
        if isinstance(row[0], str) and re.fullmatch(
            r"item\s*(?:#|number|no\.?)", row[0].strip(), re.I
        ):
            return idx
    return None


def _family_is_finished_only(rows: list[list], family_row: int | None) -> bool:
    """KKT families priced finished-only (Walnut is a wood column, not Finished ADD)."""
    if family_row is None:
        return False
    for j in range(family_row, min(family_row + 4, len(rows))):
        joined = " ".join(
            c.strip() for c in rows[j] if isinstance(c, str) and c.strip()
        ).lower()
        if "only available in finished" in joined:
            return True
        if "not available unfinished" in joined:
            return True
        texts = [
            c.strip().lower()
            for c in rows[j]
            if isinstance(c, str) and c.strip()
        ]
        fin = sum(
            1
            for t in texts
            if re.fullmatch(r"fin\.?", t)
            or (
                re.search(r"\bfinished\b", t)
                and not re.search(r"finished\s*add\b", t)
            )
        )
        unf = sum(
            1
            for t in texts
            if re.fullmatch(r"unf(?:in)?\.?", t)
            or (
                re.search(r"\bunfinished\b", t) and not re.search(r"\bfinished\b", t)
            )
        )
        if fin >= 3 and unf == 0:
            return True
    return False


def _finished_only_wood_cols(rows: list[list], start: int) -> set[int]:
    """Finished-only KKT chair rows: first contiguous wood wholesale band (skip twin)."""
    family_row = _nearest_family_header_row(rows, start)
    if not _family_is_finished_only(rows, family_row):
        return set()
    row = rows[start]
    _, label_j = _row_item_label(row)
    min_j = (label_j + 1) if label_j >= 0 else 1
    nums = sorted(
        j
        for j in range(min_j, len(row))
        if _is_wholesale_price(row[j], allow_yearish=True)
    )
    if len(nums) < 2:
        return set()
    run: list[int] = [nums[0]]
    for j in nums[1:]:
        if j <= run[-1] + 2:
            run.append(j)
        else:
            break
    return set(run)


def _header_block_near(rows: list[list], start: int) -> tuple[list[list], int]:
    header_start = None
    # Prefer a real header above the current row (skip the current data line).
    header_candidates: list[int] = []
    for back in range(1, min(start, 60) + 1):
        idx = start - back
        row = rows[idx] if idx < len(rows) else []
        if _row_looks_like_column_header(row):
            header_candidates.append(idx)
    if header_candidates:
        for idx in header_candidates:
            row = rows[idx]
            if any(
                isinstance(c, str)
                and re.search(r"^item\s*(?:#|number|no\.?)$", c.strip(), re.I)
                for c in row
            ):
                header_start = idx
                break
        if header_start is None:
            header_start = header_candidates[0]
    if header_start is None:
        # Fall back: UNFINISHED / FINISH band rows (Heritage Occasionals)
        for back in range(1, min(start, 60) + 1):
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
    # Include rows above the Item/Description header so section labels like
    # L&N "UNFINISHED" / "ESTIMATED FINISHED PRICE" and KKT dual Description
    # panels participate in column detection.
    family_row = _nearest_family_header_row(rows, start)
    finished_only = _family_is_finished_only(rows, family_row)
    block_start = max(0, header_start - 4)
    if finished_only and family_row is not None:
        block_start = max(block_start, family_row)
    header_block = rows[block_start : header_start + 12]
    # Raleigh benches reprint Item/wood headers without the Fin|Unf sub-row —
    # pull the nearest Fin|Unf pair labels from further above into the block.
    if not any(_row_has_fin_unf_pair_labels(h) for h in header_block):
        for back in range(1, min(header_start, 80) + 1):
            idx = header_start - back
            if idx < 0:
                break
            if finished_only and family_row is not None and idx < family_row:
                break
            if _row_has_fin_unf_pair_labels(rows[idx]):
                header_block = [rows[idx]] + list(header_block)
                block_start = min(block_start, idx)
                break
    if finished_only and family_row is not None and family_row > block_start:
        header_block = rows[family_row : header_start + 12]
        block_start = family_row
    block_end = min(len(rows), header_start + 12)
    for j in range(header_start + 1, min(len(rows), header_start + 16)):
        if (
            j < len(rows)
            and rows[j]
            and isinstance(rows[j][0], str)
            and re.fullmatch(
                r"item\s*(?:#|number|no\.?)", rows[j][0].strip(), re.I
            )
        ):
            block_end = j
            break
    header_block = rows[block_start:block_end]
    width = max((len(r) for r in rows[block_start : header_start + 40]), default=0)
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


# Mid-columns that are type / option tokens, not product descriptions.
_MID_TEXT_DENY = frozenset(
    {
        "chair",
        "upholstered",
        "upholstery",
        "standard",
        "optional",
        "swivel",
        "yes",
        "no",
    }
)
_BED_SIZE_RE = re.compile(
    r"^(twin|full|queen|king|cal(?:ifornia)?\s*king|"
    r"full\s*/?\s*queen|king\s*/?\s*cal(?:ifornia)?\s*king)$",
    re.I,
)


def _keep_item_mid_text(text: str) -> bool:
    """True for descriptions / bed sizes / Add-$ notes; false for type tokens."""
    t = text.strip()
    if not t:
        return False
    if t.casefold() in _MID_TEXT_DENY:
        return False
    if re.search(r"add\s+\$", t, re.I):
        return True
    # Bed / mattress size labels (Kid's Comfort, Kohler, …)
    if _BED_SIZE_RE.match(t):
        return True
    # Size-like: 10" x 18" or 25" H
    if re.search(r"\d+\s*[\"″\'′]|(\d+\s*[xX×]\s*\d+)", t):
        return True
    # Multi-word or long descriptions (J&M "QUICK SHIP Storage Rails…")
    if " " in t and len(t) >= 8:
        return True
    if len(t) >= 16:
        return True
    # Single-word furniture names (Dresser, Armoire, Nightstand, Chifferobe, …)
    if len(t) >= 5 and re.fullmatch(r"[A-Za-z][A-Za-z'./-]*", t):
        return True
    return False


def _transform_row_prices(
    row: list,
    multiplier: float,
    price_cols: Iterable[int],
    *,
    spots: list[SpotSample],
    item_hint: str,
) -> list[str]:
    cells: list[str] = []
    label, label_j = _row_item_label(row)
    if not label:
        label = item_hint
        label_j = -1
    if re.search(r"\bapprox", label, re.I):
        return []
    lab = sanitize_retail_text(label, multiplier)
    if not lab:
        return []
    cells.append(lab)
    price_set = set(price_cols)
    first_price = min(price_set) if price_set else len(row)
    # Keep Description / Size / "Add $… for Toe Kick" notes after the item code.
    # Skip short single-token mid-labels (Hope Works "Chair" / "Upholstered").
    for j in range(max(label_j + 1, 0), min(label_j + 3, first_price)):
        if j in price_set or j >= len(row):
            continue
        v = row[j]
        if _is_wholesale_price(v) or _is_na_cell(v):
            continue
        if not isinstance(v, str):
            continue
        # Stop if this looks like another item code (start of a mirrored panel).
        if j > label_j + 1 and ITEM_CODE_RE.match(v.strip()):
            break
        if v.strip().casefold() == lab.casefold():
            continue
        t = sanitize_retail_text(v, multiplier)
        if not t or is_dealer_terms_text(t):
            continue
        if not _keep_item_mid_text(t):
            continue
        cells.append(t)
    got_price = False
    for j in sorted(price_set):
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
        if is_call_for_pricing_cell(v):
            # Leave blank — do not invent a retail number for call-for cells.
            cells.append("")
            continue
        # Price-column cells may legitimately be 2015–2035 wholesale amounts.
        if _is_wholesale_price(v, allow_yearish=True):
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
    """True when most priced product rows actually store values in the labeled columns.

    Ignore short option/finish-add lines (Quick Ship '2 Tone' / 'Paint' / 'Glaze')
    so they cannot outvote real chair/table SKUs.
    """
    hits = misses = 0
    skip_labs = {
        "2 tone",
        "2-tone",
        "paint",
        "glaze",
        "tables",
        "benches",
        "chairs",
        "options",
    }
    for row in rows[start : start + 120]:
        if not any(_is_wholesale_price(c) for c in row):
            continue
        lab0 = _cell_str(row[0]) if row else ""
        if lab0.casefold().strip() in skip_labs or (
            lab0 and len(lab0) <= 8 and not any(_is_wholesale_price(row[j]) for j in labeled if j < len(row))
        ):
            # Lone option surcharge rows — do not count as product misses
            if not any(j < len(row) and _is_wholesale_price(row[j]) for j in labeled):
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
            s = c.lower().strip()
            if re.fullmatch(r"unf(?:in)?\.?", s) or (
                re.search(r"\bunfinished\d*\b", s) and not re.search(r"\bfinished\b", s)
            ):
                unfinished.add(j)
            elif (
                re.search(r"\bfinished\b", s) and not ADD_ON_PRICE_COL_RE.search(s)
            ) or re.search(r"\b(?:stain|glaz(?:e|ing)|2[\s-]?tone|paint)\b", s):
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


def _addon_cols_above(rows: list[list], start: int, width: int) -> set[int]:
    """Find finish-Add / Add-for columns from any header row above the data line.

    Jericho puts 'Finished Prices Add' as the last headed price column; anything
    to its right is noise/flags. Kohler mirrors 'Add for 2-Tone' on each panel.
    """
    bad: set[int] = set()
    # Scan from the top — Kohler collection sheets put wood/Add headers on row 1
    # while priced SKUs can sit well past a short lookback window.
    finished_add_sentinels: list[int] = []
    for row in rows[: max(start, 1)]:
        for j in range(min(width + 1, len(row))):
            c = row[j]
            if not isinstance(c, str):
                continue
            if ADD_ON_PRICE_COL_RE.search(c):
                bad.add(j)
                if re.search(r"finished\s*prices?\s*add", c, re.I):
                    finished_add_sentinels.append(j)
    if finished_add_sentinels:
        # Dual-panel books (Fenton): exclude from the *leftmost* Finished Prices Add
        # through the end of that panel, and the rightmost Add itself — wood cols on
        # the right panel stay eligible via _wood_cols_before_finished_add.
        for edge in finished_add_sentinels:
            bad.add(edge)
        # Trailing noise to the right of the last Finished Prices Add
        last = max(finished_add_sentinels)
        for j in range(last, width + 1):
            bad.add(j)
    return bad


def _wood_cols_before_finished_add(rows: list[list], width: int) -> set[int]:
    """Jericho Master/Fenton: wood price columns sit left of 'Finished Prices Add'.

    When two Add headers exist (mirrored Fenton panels), keep only the right panel
    so markup twins are not duplicated.
    """
    add_cols: list[int] = []
    for row in rows[:20]:
        for j in range(min(width + 1, len(row))):
            c = row[j]
            if isinstance(c, str) and re.search(r"finished\s*prices?\s*add", c, re.I):
                if j not in add_cols:
                    add_cols.append(j)
    if not add_cols:
        return set()
    add_cols.sort()
    # Use the rightmost panel when dual-panel; otherwise the only Add.
    add_j = add_cols[-1]
    prev_add = add_cols[-2] if len(add_cols) >= 2 else -1
    woods: set[int] = set()
    # Walk left from Add over wood-header / blank band; stop at Item/Description/Size.
    stop_re = re.compile(r"item\s*(?:no|#)|description|\bsize\b", re.I)
    for row in rows[:12]:
        for j in range(add_j - 1, prev_add, -1):
            if j < 0 or j >= len(row):
                continue
            c = row[j]
            if isinstance(c, str) and stop_re.search(c.strip()):
                break
            if isinstance(c, str) and re.search(
                r"oak|maple|cherry|hickory|alder|elm|ash|qswo|walnut|rqswo",
                c,
                re.I,
            ):
                woods.add(j)
            elif isinstance(c, str) and c.strip() and not ADD_ON_PRICE_COL_RE.search(c):
                # Non-wood label (e.g. blank Size on second panel) — keep scanning
                if stop_re.search(c):
                    break
    if woods:
        # Fill contiguous band from min wood through add_j-1 (covers empty header cells)
        return set(range(min(woods), add_j))
    # Fallback: contiguous numerics commonly occupy the 6 cols left of Add
    return set(range(max(prev_add + 1, add_j - 6), add_j))


def _has_fabric_or_option_headers(header_block: list[list], before_col: int) -> bool:
    """Comfort-style surcharges only; skip L&N unfinished wood bands."""
    for h in header_block:
        for j, c in enumerate(h):
            if j >= before_col or not isinstance(c, str):
                continue
            s = c.casefold()
            if any(
                k in s
                for k in (
                    "fabric",
                    "leather",
                    "swivel",
                    "lay flat",
                    "usb",
                    "head tilt",
                    "electric",
                )
            ):
                return True
    return False


def _leading_option_price_cols(
    rows: list[list], start: int, before_col: int
) -> set[int]:
    """Lambright Comfort: electric/swivel/USB surcharges before fabric grade columns."""
    row = rows[start]
    _, label_j = _row_item_label(row)
    lo = max((label_j + 1) if label_j >= 0 else 1, 1)
    hi = max(lo, before_col)
    return {
        j
        for j in range(lo, hi)
        if j < len(row) and _is_wholesale_price(row[j], allow_yearish=True)
    }


def _infer_price_cols(rows: list[list], start: int) -> set[int]:
    cols: set[int] = set()
    # Prefer a nearby header row above the current line when present
    header_block, width = _header_block_near(rows, start)
    item_cols = _item_code_cols(header_block, width)
    excluded_early = _excluded_price_header_cols(header_block, width)
    est_band = _estimated_finished_band_cols(header_block, width)
    if est_band and _labeled_cols_have_prices(rows, start, est_band):
        return est_band - item_cols - excluded_early
    only_finished = _finished_only_wood_cols(rows, start)
    if only_finished and _labeled_cols_have_prices(rows, start, only_finished):
        return only_finished - item_cols - excluded_early
    width = max(
        width,
        max((len(r) for r in rows[: start + 1]), default=0),
        max((len(r) for r in rows[max(0, start - 60) : start + 1]), default=0),
    )
    excluded = _excluded_price_header_cols(header_block, width) | _addon_cols_above(
        rows, start, width
    )
    wood_band = _wood_cols_before_finished_add(rows, width) - excluded - item_cols
    if wood_band and _labeled_cols_have_prices(rows, start, wood_band):
        return wood_band
    labeled = _finished_price_cols(header_block, width) - excluded
    # Prefer finished/stained columns that actually hold values (Hilltop mirror left empty)
    if labeled:
        populated = {
            j
            for j in labeled
            if any(
                j < len(row) and _is_wholesale_price(row[j], allow_yearish=True)
                for row in rows[start : start + 120]
            )
        }
        if populated and _labeled_cols_have_prices(rows, start, populated):
            chosen = populated - item_cols - excluded
            if chosen and _has_fabric_or_option_headers(header_block, min(chosen)):
                chosen |= _leading_option_price_cols(rows, start, min(chosen))
            return chosen
        # Fin|Unf / fabric labels with real values win over sparse option-line noise
        # (Quick Ship chairs after Oregon Arm) even when hits < misses.
        if populated and len(populated) >= 1:
            chosen = populated - item_cols - excluded
            if chosen and _has_fabric_or_option_headers(header_block, min(chosen)):
                chosen |= _leading_option_price_cols(rows, start, min(chosen))
            return chosen
    if labeled and _labeled_cols_have_prices(rows, start, labeled):
        return labeled - item_cols - excluded
    if labeled:
        mirrored = _mirror_finished_cols(rows, start, header_block, width)
        if mirrored:
            return mirrored - item_cols - excluded
        # J Miller: "Finished" header sits one column right of the value cells.
        # Skip when Finished is paired with Unfinished on the right (KKT Quick Ship).
        pair_unf_right = False
        for h in header_block:
            for j in labeled:
                if j + 1 < len(h) and isinstance(h[j + 1], str):
                    s = h[j + 1].lower().strip()
                    if re.fullmatch(r"unf(?:in)?\.?", s) or (
                        re.search(r"\bunfinished\b", s)
                        and not re.search(r"\bfinished\b", s)
                    ):
                        pair_unf_right = True
                        break
            if pair_unf_right:
                break
        if not pair_unf_right:
            left_of_finished = {
                j - 1
                for j in labeled
                if j > 0
                and any(
                    j - 1 < len(row) and _is_wholesale_price(row[j - 1])
                    for row in rows[start : start + 80]
                )
            }
            if left_of_finished and _labeled_cols_have_prices(
                rows, start, left_of_finished
            ):
                return left_of_finished - item_cols - excluded
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
            if re.search(r"wood\s*grain|brights", s):
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
            if (
                s in {"unfinished", "unfin", "unf", "unf."}
                or re.fullmatch(r"unf(?:in)?\.?", s)
                or (
                    re.search(r"\bunfinished\b", s)
                    and not re.search(r"\bfinished\b", s)
                )
            ):
                bad.add(j)
    return {j for j in cols if j not in bad}

def _row_unf_fin_kind(row: list) -> str | None:
    """INTEG-style rows labeled 'unfinished' / 'finished' in an early status column.

    Ignore trailing legend cells (L&N 'Finished' reprint labels) so we do not
    treat every priced row as an INTEG finished twin.
    """
    for c in row[:4]:
        if not isinstance(c, str):
            continue
        s = c.strip().lower()
        if s == "unfinished":
            return "unfinished"
        if s == "finished":
            return "finished"
    return None


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
    last_finished_add: FinishedAddLayout | None = None
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
        # INTEG: unfinished row + finished row pair → keep finished prices only
        pair_kind = _row_unf_fin_kind(row)
        if pair_kind == "unfinished":
            label_u, _ = _row_item_label(row)
            if not label_u:
                # Size / opening text often sits beside the unfinished tag
                for c in row:
                    if isinstance(c, str) and c.strip().lower() not in {
                        "unfinished",
                        "finished",
                    }:
                        t = sanitize_retail_text(c, multiplier)
                        if t and len(t) >= 3:
                            label_u = t
                            break
            if label_u:
                item_hint = label_u
            nxt = rows[i + 1] if i + 1 < len(rows) else None
            if nxt is not None and _row_unf_fin_kind(nxt) == "finished":
                price_cols = [
                    j for j, c in enumerate(nxt) if _is_wholesale_price(c)
                ]
                # Drop leading finishing-delta cells (col 0/1) when wood prices sit further right
                if len(price_cols) >= 3:
                    woods = [j for j in price_cols if j >= 4]
                    if woods:
                        price_cols = woods
                if price_cols:
                    cells = _transform_row_prices(
                        nxt,
                        multiplier,
                        price_cols,
                        spots=spots,
                        item_hint=item_hint or label_u or "",
                    )
                    # Prefer unfinished-row label when finished row has none
                    if cells and (item_hint or label_u):
                        lab = sanitize_retail_text(item_hint or label_u, multiplier)
                        if lab:
                            cells[0] = lab
                    if cells:
                        out.append(ProcessedRow(RowKind.WIDE, cells))
                i += 2
                continue
            # Unfinished without a finished twin — drop
            i += 1
            continue
        if pair_kind == "finished":
            # Orphan finished row (label on prior unfinished already consumed)
            price_cols = [j for j, c in enumerate(row) if _is_wholesale_price(c)]
            if price_cols:
                cells = _transform_row_prices(
                    row, multiplier, price_cols, spots=spots, item_hint=item_hint
                )
                if cells:
                    out.append(ProcessedRow(RowKind.WIDE, cells))
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
        # PDF never ends a line with a bare "Add". Keep percent-add prose
        # ("For 1¼\" tops add 15%") and section titles for the Options page.
        if (
            _row_has_add_label(row)
            and not _row_has_embedded_dollar_text(row)
            and not re.search(r"add\s+\d+\s*%", joined, re.I)
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
                if not prose or is_dealer_terms_text(prose):
                    continue
                # Keep short ALL-CAPS section titles (FOR TWO TONE) and percent adds
                if len(prose) <= 12 and not (
                    prose.isupper() or re.search(r"\d+\s*%", prose)
                ):
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
        # KKT Options: "2-tone Stain : Add $60" etc. are product-note prose even
        # when ITEM_CODE_RE historically matched the leading "2-t".
        if (
            label_early
            and re.search(r"\badd\b", joined, re.I)
            and (
                _row_has_embedded_dollar_text(row)
                or re.search(r"add\s+\d+\s*%", joined, re.I)
            )
            and not any(_is_wholesale_price(c) for c in row)
        ):
            prose = _strip_contact_pricing_phrases(
                sanitize_retail_text(joined, multiplier)
            )
            if prose and len(prose) > 8 and not is_dealer_terms_text(prose):
                out.append(ProcessedRow(RowKind.PROSE, [prose]))
                i += 1
                continue
        # Short ALL-CAPS option section titles (TOPS / FOR TWO TONE)
        if (
            label_early
            and label_early.isupper()
            and 3 <= len(label_early) <= 40
            and not any(_is_wholesale_price(c) for c in row)
            and not re.search(r"\d", label_early)
        ):
            prose = sanitize_retail_text(label_early, multiplier)
            if prose:
                out.append(ProcessedRow(RowKind.SUBHEAD, [prose]))
                i += 1
                continue
        if (
            row
            and isinstance(row[0], str)
            and re.fullmatch(r"item\s*(?:#|number|no\.?)", row[0].strip(), re.I)
        ):
            if _family_is_finished_only(rows, i):
                last_finished_add = None
        finished_add = _detect_finished_add_layout(rows, i) or last_finished_add
        if finished_add and _family_is_finished_only(
            rows, _nearest_family_header_row(rows, i)
        ):
            finished_add = None
        if finished_add and any(_is_wholesale_price(c) for c in row):
            cells = _transform_finished_add_row(
                row, finished_add, multiplier, spots=spots, item_hint=item_hint
            )
            if cells:
                last_finished_add = finished_add
                out.append(ProcessedRow(RowKind.WIDE, cells))
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
        row_width = max(len(row), max((len(r) for r in rows[max(0, i - 60) : i + 1]), default=0))
        addon_skip = _addon_cols_above(rows, i, row_width)
        excluded_data = (
            _excluded_price_header_cols(header_block, row_width) | addon_skip
        )
        if price_cols:
            use_cols = [
                j for j in sorted(price_cols) if j not in addon_skip and j not in excluded_data
            ]
        else:
            use_cols = [
                j
                for j in range(len(row))
                if j not in addon_skip and j not in excluded_data
            ]
        cells = _transform_row_prices(
            row, multiplier, use_cols, spots=spots, item_hint=item_hint
        )
        if not cells and any(_is_wholesale_price(c) for c in row):
            # Inferred wood cols can miss Add-upcharge rows (Hillside Hardwood XX
            # mirrored panels). Fall back to actual wholesale positions; prefer the
            # left panel when the same Add label is duplicated.
            nums = [
                j
                for j, c in enumerate(row)
                if _is_wholesale_price(c) and j not in excluded_data
            ]
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
            nums = [
                j
                for j, c in enumerate(row)
                if _is_wholesale_price(c) and j not in excluded_data
            ]
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
        # (e.g. 'Only offered in 60" width.'), or section labels like
        # 'Rectangle Glass Tops:' after a call-for parenthetical is removed.
        keep_prose = len(joined_prose) > 40 or (
            prose_before_len > len(joined_prose) and len(joined_prose) > 12
        ) or any(
            len(c) > 8 and c.rstrip().endswith(":") for c in prose_cells
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


def _markup_calculator_without_finished(rows: list[list]) -> bool:
    """True for HW Chair-style markup calculators that expose unfinished bases only."""
    head = " ".join(
        _cell_str(c) for row in rows[:6] for c in row if _cell_str(c)
    ).lower()
    if not re.search(r"markup\s*calculator", head):
        return False
    # Optional finishing adder ≠ finished price columns
    has_finished_col = False
    for row in rows[:8]:
        for c in row:
            if not isinstance(c, str):
                continue
            s = c.lower().strip()
            if re.search(r"\bunfinished\b", s):
                continue
            if re.search(r"cost\s+of\s+finishing|nat\.?\s*finish", s):
                continue
            if re.search(r"\bfinished\b|\bfin\.?\b", s):
                has_finished_col = True
                break
        if has_finished_col:
            break
    return not has_finished_col


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
    # Hoosier Poly / Interior Hardwoods: keep plain price list, drop markup twin
    plain_prices = [
        n
        for n, v in lower.items()
        if re.search(r"\bprices?\b|\bprice\s*list\b", v) and not re.search(r"markup", v)
    ]
    if plain_prices:
        for n, v in lower.items():
            if re.search(r"prices?\s+with\s+markup|\bwith\s+markup\b|price\s*list\s*markup", v):
                skip.add(n)
            elif re.search(r"\bmarkup\b", v) and re.search(r"price\s*list", v):
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
    # When Master holds a real priced grid, skip duplicate Sheet*/Pricing/Products/Title.
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
                    or v == "products"
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
    markup_calc_only = False
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
        if _markup_calculator_without_finished(rows):
            # HW Chair Catalog: unfinished bases + optional finish adder, no finished grid
            skipped.append(name)
            markup_calc_only = True
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
        reason = f"No price grid in {path.name}"
        if markup_calc_only:
            reason = (
                f"No finished prices (markup calculator / unfinished bases only) "
                f"in {path.name}"
            )
        return ViztechBuildResult([], spots, skipped, [reason])

    return ViztechBuildResult(collapse_blank_rows(out), spots, skipped, [])


def merge_builder_files(
    paths: list[Path],
    *,
    builder_name: str,
    multiplier: float,
) -> ViztechBuildResult:
    key = builder_name.casefold()
    if key == "kountry knob tables":
        mains = [p for p in paths if re.search(r"KKT_Pricelist", p.name, re.I)]
        if mains:
            paths = mains
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
