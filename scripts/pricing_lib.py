"""Retail pricing transforms for dealer wholesale sheets."""

from __future__ import annotations

import math
import re
from decimal import Decimal, ROUND_UP
from typing import Callable

# Whole-dollar amounts: $1,234 or $1234.00
DOLLAR_RE = re.compile(
    r"(?<!\w)"
    r"\$\s*"
    r"(\d{1,3}(?:,\d{3})*|\d+)"
    r"(?:\.(\d{1,2}))?"
    r"(?!\w)"
)

# Bare decimals that look like wholesale prices (two decimal places, not dimensions)
DECIMAL_PRICE_RE = re.compile(
    r"(?<![\d/xX×%\$])"
    r"(\d{1,3}(?:,\d{3})*|\d+)"
    r"\.(\d{2})"
    r"(?![\d/xX×%])"
)

DIMENSION_RE = re.compile(r"\d+\s*[xX×]\s*\d+")
PERCENT_RE = re.compile(r"\d+(?:\.\d+)?\s*%")

NON_PRICE_HEADERS = frozenset(
    {
        "item",
        "item#",
        "item #",
        "sku",
        "qty",
        "quantity",
        "size",
        "dimensions",
        "dim",
        "width",
        "height",
        "depth",
        "length",
        "model",
        "description",
        "desc",
    }
)


def retail_whole_dollar(wholesale: float, multiplier: float) -> int:
    """Retail = wholesale × multiplier, rounded up to the next whole dollar."""
    product = Decimal(str(wholesale)) * Decimal(str(multiplier))
    return int(product.to_integral_value(rounding=ROUND_UP))


def _parse_dollars(amount: str, cents: str | None) -> float | None:
    whole = amount.replace(",", "")
    if cents is None:
        return float(whole)
    if len(cents) == 1:
        cents = cents + "0"
    return float(f"{whole}.{cents}")


def transform_dollar_match(
    match: re.Match[str], multiplier: float
) -> str:
    wholesale = _parse_dollars(match.group(1), match.group(2))
    if wholesale is None or wholesale <= 0:
        return match.group(0)
    retail = retail_whole_dollar(wholesale, multiplier)
    return f"${retail:,}"


def transform_text_prices(text: str, multiplier: float) -> str:
    """Replace wholesale dollar amounts in a string with retail whole dollars."""
    if not text or not isinstance(text, str):
        return text
    if PERCENT_RE.search(text) and "$" not in text:
        return text
    if DIMENSION_RE.search(text) and "$" not in text:
        return text

    def sub_dollar(m: re.Match[str]) -> str:
        return transform_dollar_match(m, multiplier)

    return DOLLAR_RE.sub(sub_dollar, text)


def is_price_column(header: str | None) -> bool:
    if not header:
        return False
    h = str(header).strip().lower()
    if h in NON_PRICE_HEADERS:
        return False
    return any(
        kw in h
        for kw in ("price", "wholesale", "cost", "msrp", "retail", "amount", "$")
    )


def transform_cell_value(
    value,
    multiplier: float,
    column_header: str | None,
    *,
    force_price: bool = False,
) -> str | int | float | None:
    if value is None or value == "":
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        if is_price_column(column_header) or force_price:
            if value <= 0:
                return value
            return retail_whole_dollar(float(value), multiplier)
        return value

    text = str(value).strip()
    if PERCENT_RE.fullmatch(text) or text.endswith("%"):
        return text
    if DIMENSION_RE.search(text):
        return transform_text_prices(text, multiplier) if "$" in text else text

    if "$" in text:
        return transform_text_prices(text, multiplier)
    if is_price_column(column_header):
        try:
            num = float(text.replace(",", "").replace("$", ""))
            if num > 0:
                return retail_whole_dollar(num, multiplier)
        except ValueError:
            pass
    return transform_text_prices(text, multiplier) if "$" in text else text


def assert_no_approx(*parts: str) -> None:
    combined = " ".join(p for p in parts if p).lower()
    if "approx" in combined:
        raise ValueError('Output must not contain the word "approx".')
