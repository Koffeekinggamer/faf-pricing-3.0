"""Shared retail text/price helpers for multi-sheet workbooks."""

from __future__ import annotations

import re
from typing import Any

from scripts.dealer_terms import is_dealer_terms_text
from scripts.finished_only import UNFINISHED_NOTE_RE, drop_unfinished_note_text
from scripts.pdf_plain_text import clean_pdf_plain_text
from scripts.pricing_lib import assert_no_approx, retail_whole_dollar, transform_text_prices

WHOLESALE_RE = re.compile(r"\bwholesale\b", re.I)
WHOLESALE_PRICELIST_RE = re.compile(r"wholesale\s+pricelist", re.I)
DEALER_RE = re.compile(r"\bdealer\b", re.I)
PREFERRED_RE = re.compile(r"preferred\s+dealer", re.I)
FOB_RE = re.compile(r"f\.?\s*o\.?\s*b", re.I)
PERCENT_RE = re.compile(r"\d+(?:\.\d+)?\s*%")

NOT_AVAILABLE = "—"

# Bare Excel percent fractions written into prose (e.g. "add 0.1 of regular price").
ADD_FRACTION_OF_PRICE_RE = re.compile(
    r"(?i)\b(add)\s+(0\.\d+)\s+(of|to)\s+(regular\s+price)"
)


def is_percent_token(value: Any) -> bool:
    """True for display tokens like '4%' produced from Excel percent formats."""
    return isinstance(value, str) and bool(re.fullmatch(r"\d+(?:\.\d+)?%", value.strip()))


def excel_percent_to_display(value: float) -> str:
    """Render an Excel percent fraction (0.04) as '4%' without applying a multiplier."""
    pct = float(value) * 100.0
    if abs(pct - round(pct)) < 1e-6:
        return f"{int(round(pct))}%"
    text = f"{pct:.4f}".rstrip("0").rstrip(".")
    return f"{text}%"


def rewrite_add_fraction_as_percent(text: str) -> str:
    """Turn 'add 0.1 of regular price' into 'add 10% of regular price'."""

    def _repl(m: re.Match[str]) -> str:
        frac = float(m.group(2))
        return f"{m.group(1)} {excel_percent_to_display(frac)} {m.group(3)} {m.group(4)}"

    return ADD_FRACTION_OF_PRICE_RE.sub(_repl, text)


def retail_price_cell(value: Any, multiplier: float) -> str:
    if value is None or value == "":
        return NOT_AVAILABLE
    # Percent display tokens must never be multiplied as wholesale dollars.
    if is_percent_token(value):
        return str(value).strip()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value == 0:
            return NOT_AVAILABLE
        return f"${retail_whole_dollar(float(value), multiplier):,}"
    text = str(value).strip()
    if not text:
        return NOT_AVAILABLE
    if text == "0":
        return NOT_AVAILABLE
    if is_percent_token(text):
        return text
    if "$" in text:
        return transform_text_prices(text, multiplier)
    try:
        num = float(text.replace(",", "").replace("$", ""))
        if num == 0:
            return NOT_AVAILABLE
        return f"${retail_whole_dollar(num, multiplier):,}"
    except ValueError:
        return text


def sanitize_retail_text(text: str, multiplier: float, *, allow_percent: bool = True) -> str:
    if not text or not isinstance(text, str):
        return ""
    text = clean_pdf_plain_text(text)
    if is_dealer_terms_text(text):
        return ""
    if UNFINISHED_NOTE_RE.search(text):
        return ""
    cleaned = drop_unfinished_note_text(text)
    if not cleaned:
        return ""
    text = cleaned
    text = rewrite_add_fraction_as_percent(text)

    def _scrub_approx(s: str) -> str:
        s = re.sub(r"\bapproximately\b", "about", s, flags=re.I)
        s = re.sub(r"\bapproximate\b", "estimated", s, flags=re.I)
        s = re.sub(r"\bapprox\.?\b", "", s, flags=re.I)
        return re.sub(r"\s{2,}", " ", s).strip()

    if allow_percent and PERCENT_RE.search(text) and "$" not in text:
        t = WHOLESALE_RE.sub("", text)
        t = DEALER_RE.sub("", t)
        t = _scrub_approx(t)
        assert_no_approx(t)
        return t
    if FOB_RE.search(text):
        return ""
    t = WHOLESALE_RE.sub("", text)
    t = PREFERRED_RE.sub("", t)
    t = DEALER_RE.sub("", t)
    # Customer PDFs must not contain "approx"; rewrite dimension hedges.
    t = _scrub_approx(t)
    if "$" in t:
        t = transform_text_prices(t, multiplier)
    assert_no_approx(t)
    return t


def forbidden_phrases(*parts: str) -> list[str]:
    bad = []
    for p in parts:
        s = (p or "").lower()
        if "wholesale" in s:
            bad.append(p)
        if "preferred dealer" in s:
            bad.append(p)
        if FOB_RE.search(s):
            bad.append(p)
    return bad
