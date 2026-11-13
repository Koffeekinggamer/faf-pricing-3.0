from pathlib import Path

import pytest

from scripts.country_view_oak_book import (
    _transform_price_line,
    is_country_view_oak_pdf,
    process_country_view_oak_pdf,
)
from scripts.pricing_lib import retail_whole_dollar

PDF = Path("/tmp/Country_View_Oak_Finished_April2025.pdf")


def test_country_view_oak_line_retail_math():
    cells = _transform_price_line(
        '36" X 48" Solid 4 $1,007 $1,359 $1,309 $1,208 $1,460', 2.7
    )
    assert cells[1] == "$2,719"
    assert cells[-1] == "$3,942"
    assert retail_whole_dollar(71, 2.7) == 192


@pytest.mark.skipif(not PDF.is_file(), reason="Country View Oak upload PDF not on disk")
def test_country_view_oak_extraction_parity():
    rows, stats = process_country_view_oak_pdf(PDF, multiplier=2.7)
    assert stats["dollar_parity"]
    assert stats["expected_wholesale_dollars"] > 5000
    assert len(rows) > 100


@pytest.mark.skipif(not PDF.is_file(), reason="Country View Oak upload PDF not on disk")
def test_country_view_oak_spot_retail():
    assert retail_whole_dollar(1007, 2.7) == 2719
    assert retail_whole_dollar(1460, 2.7) == 3942
    rows, _ = process_country_view_oak_pdf(PDF, multiplier=2.7)
    wide = [r for r in rows if r.kind.value == "wide"]
    hit_oak = [
        r for r in wide if r.cells and "36" in str(r.cells[0]) and "$2,719" in r.cells
    ]
    hit_walnut = [
        r for r in wide if r.cells and "36" in str(r.cells[0]) and "$3,942" in str(r.cells)
    ]
    assert hit_oak, "Expected 36x48 Solid oak retail $2,719"
    assert hit_walnut, "Expected 36x48 Solid walnut retail $3,942"


def test_is_country_view_oak_pdf():
    if not PDF.is_file():
        pytest.skip("upload missing")
    assert is_country_view_oak_pdf(PDF)
