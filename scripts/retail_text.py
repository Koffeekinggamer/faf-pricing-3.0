"""Shared retail text/price helpers for multi-sheet workbooks."""

from __future__ import annotations

import re
from typing import Any

from scripts.pricing_lib import assert_no_approx, retail_whole_dollar, transform_text_prices

WHOLESALE_RE = re.compile(r"\bwholesale\b", re.I)
WHOLESALE_PRICELIST_RE = re.compile(r"wholesale\s+pricelist", re.I)
DEALER_RE = re.compile(r"\bdealer\b", re.I)
PREFERRED_RE = re.compile(r"preferred\s+dealer", re.I)
FOB_RE = re.compile(r"f\.?\s*o\.?\s*b", re.I)
PERCENT_RE = re.compile(r"\d+(?:\.\d+)?\s*%")

NOT_AVAILABLE = "—"


def retail_price_cell(value: Any, multiplier: float) -> str:
    if value is None or value == "":
        return NOT_AVAILABLE
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value == 0:
            return NOT_AVAILABLE
        return f"${retail_whole_dollar(float(value), multiplier):,}"
    text = str(value).strip()
    if not text:
        return NOT_AVAILABLE
    if text == "0":
        return NOT_AVAILABLE
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
    if allow_percent and PERCENT_RE.search(text) and "$" not in text:
        t = WHOLESALE_RE.sub("", text)
        t = DEALER_RE.sub("", t)
        t = re.sub(r"\s{2,}", " ", t).strip()
        assert_no_approx(t)
        return t
    if FOB_RE.search(text):
        return ""
    t = WHOLESALE_RE.sub("", text)
    t = PREFERRED_RE.sub("", t)
    t = DEALER_RE.sub("", t)
    t = re.sub(r"\s{2,}", " ", t).strip()
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
