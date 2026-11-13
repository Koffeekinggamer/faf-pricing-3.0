"""Drop dealer-only policy, billing, and builder contact blocks from retail PDFs."""

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
    r"3\s+weeks\s+after\s+delivery|"
    r"finance\s+charg|"
    r"past\s+due|"
    r"more\s+than\s+\d+\s+days\s+past\s+due|"
    r"\d+\s*%\s+finance\s+charg|"
    r"credit\s+terms|"
    r"late\s+fee"
    r")",
    re.I | re.DOTALL,
)

BUILDER_CONTACT_RE = re.compile(
    r"(?:"
    r"\b(?:phone|ph\.?\s*:?|tel\.?\s*:?|e-?mail|email|vm\.?)\b|"
    r"\bfax\b|"
    r"fax\s+service|"
    r"next\s+day\s+fax|"
    r"@[a-z0-9][a-z0-9._-]*\.[a-z]{2,}|"
    r"\bwww\.[a-z0-9]|"
    r"\bhttps?://|"
    r"\bpo\s+box\b|"
    r"\btownship\s+road\b|"
    r"\b\d+\s+cr\s+\d+\b"
    r")",
    re.I,
)

# Physical address heuristics (avoid SKU suffixes like "143 DR" for dresser).
STREET_LINE_RE = re.compile(
    r"\b\d{1,5}\s+(?:"
    r"(?:n\.?\s*|s\.?\s*|e\.?\s*|w\.?\s*)?"
    r"(?:\w+\s+){0,4}"
    r"(?:road|rd\.|street|st\.|avenue|ave\.|drive|lane|ln\.|way|highway|hwy|boulevard|blvd\.)"
    r")\b",
    re.I,
)

AMISH_COUNTY_ROAD_RE = re.compile(r"\b\d+\s+cr\s+\d+\b", re.I)

PHONE_NUMBER_RE = re.compile(
    r"(?:\(\d{3}\)|\d{3})[\s.\-]?\d{3}[\s.\-]?\d{4}"
)

ZIP_STATE_RE = re.compile(
    r"\b(?:al|ak|az|ar|ca|co|ct|de|fl|ga|hi|ia|id|il|in|ks|ky|la|ma|md|me|mi|mn|mo|ms|mt|nc|nd|ne|nh|nj|nm|nv|ny|oh|ok|or|pa|ri|sc|sd|tn|tx|ut|va|vt|wa|wi|wv|wy)\s+\d{5}(?:-\d{4})?\b",
    re.I,
)

FINANCE_BILLING_RE = re.compile(
    r"(?:"
    r"finance\s+charg|"
    r"past\s+due|"
    r"(?:past\s+due|more\s+than)\s+\d+\s+days?|"
    r"\binvoices?\b.*(?:past\s+due|\d+\s+days?)|"
    r"\b\d+\s*%\s+finance"
    r")",
    re.I,
)

# CI guard phrases (legacy dealer-only wording).
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
    "finance charge",
    "past due",
)

_FORBIDDEN_RES = [re.compile(re.escape(p), re.I) for p in FORBIDDEN_DEALER_PHRASES]
_INVOICE_PHRASE = re.compile(r"\binvoice\b", re.I)

# Customer PDF must not expose builder contact / billing (CI extraction checks).
PHONE_LABEL_RE = re.compile(r"\b(?:phone|ph\.?\s*:?|tel\.?\s*:?)\b", re.I)
FAX_LABEL_RE = re.compile(r"\bfax\b", re.I)
EMAIL_RE = re.compile(r"@[a-z0-9][a-z0-9._-]*\.[a-z]{2,}", re.I)

CUSTOMER_FORBIDDEN_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("phone", PHONE_LABEL_RE),
    ("fax", FAX_LABEL_RE),
    ("email", EMAIL_RE),
    ("finance charge", re.compile(r"\bfinance\s+charge\b", re.I)),
    ("past due", re.compile(r"\bpast\s+due\b", re.I)),
    ("invoice", _INVOICE_PHRASE),
)


def is_builder_contact_text(text: str) -> bool:
    if not text or not str(text).strip():
        return False
    t = str(text).strip()
    if BUILDER_CONTACT_RE.search(t):
        return True
    if _physical_address_line(t):
        return True
    if PHONE_NUMBER_RE.search(t):
        if re.search(r"\b(?:phone|ph|fax|tel|call|vm)\b", t, re.I):
            return True
        if "@" in t or _physical_address_line(t) or ZIP_STATE_RE.search(t):
            return True
        if "$" not in t and len(t) < 120:
            return True
    if ZIP_STATE_RE.search(t) and (
        STREET_LINE_RE.search(t)
        or PHONE_NUMBER_RE.search(t)
        or "@" in t
        or AMISH_COUNTY_ROAD_RE.search(t)
    ):
        return True
    return False


def _physical_address_line(text: str) -> bool:
    t = text.strip()
    if re.search(r"\bpo\s+box\b", t, re.I):
        return True
    if re.search(r"\btownship\s+road\b", t, re.I):
        return True
    if AMISH_COUNTY_ROAD_RE.search(t) and (
        PHONE_NUMBER_RE.search(t) or "@" in t or ZIP_STATE_RE.search(t)
    ):
        return True
    if STREET_LINE_RE.search(t) and (
        ZIP_STATE_RE.search(t) or PHONE_NUMBER_RE.search(t) or "@" in t
    ):
        return True
    return False


def is_dealer_terms_text(text: str) -> bool:
    if not text or not str(text).strip():
        return False
    t = str(text).strip()
    if is_builder_contact_text(t):
        return True
    if DEALER_BLOCK_RE.search(t):
        return True
    if FINANCE_BILLING_RE.search(t):
        return True
    low = t.lower()
    if _INVOICE_PHRASE.search(t) and (
        "delivery" in low
        or "week" in low
        or "repair" in low
        or "warranty" in low
        or "past due" in low
        or "finance" in low
        or "charge" in low
        or "payment" in low
        or "credit" in low
        or re.search(r"\d+\s+days?", low)
    ):
        return True
    return False


def filter_prose_cells(cells: list) -> list[str]:
    out: list[str] = []
    for c in cells:
        t = str(c).strip()
        if not t or is_dealer_terms_text(t) or is_builder_contact_text(t):
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


def find_customer_forbidden_in_text(text: str) -> list[str]:
    hits: list[str] = []
    for label, pat in CUSTOMER_FORBIDDEN_PATTERNS:
        if pat.search(text):
            hits.append(label)
    if _INVOICE_PHRASE.search(text) and FINANCE_BILLING_RE.search(text):
        if "finance charge" not in hits:
            hits.append("invoice")
    # Physical address lines (contact blocks), not SKU abbreviations.
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if _physical_address_line(line):
            hits.append("street address")
            break
    return hits
