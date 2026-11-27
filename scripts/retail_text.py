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

# Customer PDFs must not invite calling/contacting the builder for prices.
CALL_FOR_PAREN_RE = re.compile(r"\(\s*call\s+for[^)]*\)", re.I)
CALL_FOR_CLAUSE_RE = re.compile(
    r"(?:^|[;\s,]+)"
    r"(?:"
    r"contact\s+us\s+for\s+[^.!;]*|"
    r"(?:please\s+)?call\s+(?:us\s+)?for\s+[^.!;]*"
    r")"
    r"\.?",
    re.I,
)
CALL_FOR_STANDALONE_RE = re.compile(
    r"(?i)^\s*(?:please\s+)?call\s+for\s+(?:pricing|a\s+quote|quote|details)\s*\.?$"
)
SPECIAL_ORDER_CALL_SENTENCE_RE = re.compile(
    r"(?:^|[.!?]\s*)[^.!?]*\bspecial\s+order\s+options\s+are\s+available\b[^.!?]*[.!]?"
    r"|"
    r"(?:^|[.!?]\s*)[^.!?]*\bspecial\s+order\b[^.!?]*\b(?:call|contact)\b[^.!?]*[.!]?"
    r"|"
    r"(?:^|[.!?]\s*)[^.!?]*\bplease\s+call\s+for\b[^.!?]*[.!]?",
    re.I,
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


def is_call_for_pricing_cell(value: Any) -> bool:
    """True for price-grid cells that only say call/contact for pricing (leave blank)."""
    if not isinstance(value, str):
        return False
    t = value.strip()
    if not t:
        return False
    if CALL_FOR_STANDALONE_RE.fullmatch(t):
        return True
    return bool(
        re.fullmatch(
            r"(?i)\s*(?:contact\s+us\s+for\s+pricing|call\s+for\s+pricing)\s*\.?\s*",
            t,
        )
    )


def strip_call_for_contact_phrases(text: str) -> str:
    """Remove call-for / contact-us pricing invites; keep product wording when present."""
    if not text or not str(text).strip():
        return ""
    raw = str(text)
    if is_call_for_pricing_cell(raw):
        return ""
    # Drop whole special-order / please-call CTA sentences.
    t = SPECIAL_ORDER_CALL_SENTENCE_RE.sub(" ", raw)
    t = CALL_FOR_PAREN_RE.sub("", t)
    t = CALL_FOR_CLAUSE_RE.sub(" ", t)
    # Any remaining sentence that still asks the reader to call/contact for a price.
    parts = re.split(r"(?<=[.!?])\s+", t)
    kept: list[str] = []
    for part in parts:
        p = part.strip()
        if not p:
            continue
        if re.search(r"(?i)\b(?:call\s+for|contact\s+us\b|contact\s+\w+\s+for\b)", p):
            continue
        kept.append(p)
    t = " ".join(kept)
    t = re.sub(r"\s{2,}", " ", t).strip()
    # Trim separators but keep a trailing label colon (e.g. "Rectangle Glass Tops:").
    t = re.sub(r"^[;\s,.-]+", "", t).strip()
    t = re.sub(r"[;\s,.-]+$", "", t).strip()
    if t and not t.endswith((".", "!", "?", '"', "'", ":")) and re.search(
        r"\b(?:width|size|sizes|only|offered)\b", t, re.I
    ):
        t = t + "."
    return t


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
    text = strip_call_for_contact_phrases(text)
    if not text:
        return ""

    def _scrub_approx(s: str) -> str:
        s = re.sub(r"\bapproximately\b", "about", s, flags=re.I)
        s = re.sub(r"\bapproximate\b", "estimated", s, flags=re.I)
        s = re.sub(r"\bapprox\.?\b", "", s, flags=re.I)
        return re.sub(r"\s{2,}", " ", s).strip()

    if allow_percent and PERCENT_RE.search(text) and "$" not in text:
        t = WHOLESALE_RE.sub("", text)
        t = DEALER_RE.sub("", t)
        t = strip_call_for_contact_phrases(t)
        t = _scrub_approx(t)
        assert_no_approx(t)
        return t
    # Freight "FOB destination" prose — not Freemont item codes like FOB-3636-O.
    if FOB_RE.search(text) and not re.match(r"^FOB[-_\d]", text.strip(), re.I):
        return ""
    t = WHOLESALE_RE.sub("", text)
    t = PREFERRED_RE.sub("", t)
    t = DEALER_RE.sub("", t)
    t = strip_call_for_contact_phrases(t)
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
        # Freight FOB prose only — keep Freemont item codes like FOB-3636-O.
        if FOB_RE.search(s) and not re.match(r"^fob[-_\d]", s.strip(), re.I):
            bad.append(p)
    return bad
