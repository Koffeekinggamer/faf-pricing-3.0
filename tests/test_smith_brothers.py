from pathlib import Path

import pytest

from scripts.pricing_lib import retail_whole_dollar
from scripts.smith_brothers_book import DOC_TITLE, MULTIPLIER, pdf_contains_prices

CATALOG = Path(__file__).resolve().parents[1] / "sheets" / "Smith_Brothers.pdf"


def test_passthrough_multiplier_identity():
    assert retail_whole_dollar(3943, MULTIPLIER) == 3943
    assert retail_whole_dollar(2325.01, MULTIPLIER) == 2326


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_spot_prices():
    needles = [
        "$3,943",
        "3,891",
        "2,325",
        "$313",
        DOC_TITLE.split("·", 1)[0].strip(),
    ]
    found = pdf_contains_prices(CATALOG, needles)
    assert found["$3,943"], "253-13 Large Sofa Grade 10"
    assert found["3,891"], "253-20 Loveseat Grade 11 (not 3,991)"
    assert found["2,325"], "253-30 Chair Grade 23"
    assert found["$313"], "253 Foam Back Option"
    # Grade 11 row lists large-sofa 3,991 then loveseat 3,891 (book column order).
    assert found["3,891"], "253-20 Loveseat Grade 11"


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_page_count():
    import pdfplumber

    with pdfplumber.open(CATALOG) as pdf:
        assert len(pdf.pages) >= 150