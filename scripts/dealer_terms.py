"""Drop dealer-only policy / billing blocks from retail PDFs."""

from __future__ import annotations

import re

# Whole-row / block drops (information sheets, TOC pointers, billing policy).
DEALER_BLOCK_RE = re.compile(
    r"(?:"
    r"invoic(?:e|ed)\s+from\s+us|"
    r"purchase\s+orders?|"
    r"order\s+confirmation|"
    r"warranty.*repair.*invoic|"
    r"repair.*invoic|"
    r"\bby\s+fax\b|"
    r"replacement\s+parts|"
    r"payment\s+terms|"
    r"std\.?\s*terms|"
    r"net\s+\d+\s+days|"
    r"table\s+of\s+contents|"
    r"three\s+weeks\s+after\s+delivery|"
    r"3\s+weeks\s+after\s+delivery"
    r")",
    re.I | re.DOTALL,
)

# CI guard phrases (must not appear in shipped PDF text).
FORBIDDEN_DEALER_PHRASES: tuple[str, ...] = (
    "purchase order",
    "order confirmation",
    "invoiced from us",
    "by fax",
    "payment terms",
    "std. terms",
    "net 30",
    "table of contents",
    "replacement parts",
    "warranty or repair invoicing",
)

_FORBIDDEN_RES = [re.compile(re.escape(p), re.I) for p in FORBIDDEN_DEALER_PHRASES]
_INVOICE_PHRASE = re.compile(r"\binvoice\b", re.I)


def is_dealer_terms_text(text: str) -> bool:
    if not text or not str(text).strip():
        return False
    t = str(text).strip()
    if DEALER_BLOCK_RE.search(t):
        return True
    low = t.lower()
    if _INVOICE_PHRASE.search(t) and (
        "delivery" in low or "week" in low or "repair" in low or "warranty" in low
    ):
        return True
    return False


def filter_prose_cells(cells: list) -> list[str]:
    out: list[str] = []
    for c in cells:
        t = str(c).strip()
        if not t or is_dealer_terms_text(t):
            continue
        out.append(t)
    return out


def find_forbidden_dealer_phrases(text: str) -> list[str]:
    hits: list[str] = []
    for pat in _FORBIDDEN_RES:
        if pat.search(text):
            hits.append(pat.pattern)
    if _INVOICE_PHRASE.search(text) and (
        "invoiced from us" in text.lower() or "after delivery" in text.lower()
    ):
        hits.append("invoice (billing policy)")
    return hits
