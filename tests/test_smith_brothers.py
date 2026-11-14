from pathlib import Path

import pdfplumber
import pytest

from scripts.pricing_lib import retail_whole_dollar
from scripts.smith_brothers_book import (
    DOC_TITLE,
    MULTIPLIER,
    SPOT_CHECKS,
    pdf_contains_prices,
    pdf_has_reversed_artifacts,
)

CATALOG = Path(__file__).resolve().parents[1] / "sheets" / "Smith_Brothers.pdf"
REVERSED_MARKERS = ("serutaeF", "snoisnemiD", "nommoC", "swolliP", "gnicirP")


def _catalog_text() -> str:
    with pdfplumber.open(CATALOG) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)


def test_passthrough_multiplier_identity():
    assert retail_whole_dollar(3943, MULTIPLIER) == 3943
    assert retail_whole_dollar(2325.01, MULTIPLIER) == 2326


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_spot_prices():
    needles = [
        "$3,943",
        "$3,843",
        "3,891",
        "2,325",
        "$313",
        DOC_TITLE.split("·", 1)[0].strip(),
    ]
    found = pdf_contains_prices(CATALOG, needles)
    assert found["$3,943"], "253-13 Large Sofa Grade 10"
    assert found["$3,843"], "253-20 Loveseat Grade 10"
    assert found["3,891"], "253-20 Loveseat Grade 11 (not 3,991)"
    assert found["2,325"], "253-30 Chair Grade 23"
    assert found["$313"], "253 Foam Back Option"


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_no_reversed_word_artifacts():
    assert pdf_has_reversed_artifacts(CATALOG) == []
    text = _catalog_text()
    for marker in REVERSED_MARKERS:
        assert marker not in text, f"reversed artifact {marker!r} in catalog"


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_source_price_spot_checks():
    """At least 10 book prices (grade or option) must appear verbatim in the retail PDF."""
    text = _catalog_text()
    checked = 0
    for item, grade, price in SPOT_CHECKS:
        price_key = price.lstrip("-+").replace("$", "").replace(",", "")
        assert price in text or price_key in text.replace(",", "").replace("$", ""), (
            f"{item} / {grade}: expected {price}"
        )
        if grade.lower().startswith("grade"):
            assert grade.split()[0] in text and grade.split()[1] in text
        checked += 1
    assert checked >= 10


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_structured_page_count():
    with pdfplumber.open(CATALOG) as pdf:
        assert 30 <= len(pdf.pages) <= 80, "structured catalog should be compact vs raw dump"