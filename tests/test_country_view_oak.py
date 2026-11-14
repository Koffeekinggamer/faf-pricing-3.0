import os
from pathlib import Path

import pytest

from scripts.country_view_oak_book import (
    BUILDER_MARKUP,
    MULTIPLIER_DEFAULT,
    is_country_view_oak_workbook,
    process_country_view_oak_workbook,
    retail_from_wholesale,
    spot_check_36x48_solid,
    wholesale_values_five_group,
)
from scripts.pricing_lib import retail_whole_dollar

CVO_XLSX = Path(
    os.environ.get(
        "CVO_WORKBOOK",
        "/home/ubuntu/.cursor/projects/workspace/uploads/"
        "Country_View_Oak_April2025_fb1d.xlsx",
    )
)


def test_oak_36x48_solid_site_price_from_wholesale():
    assert retail_whole_dollar(357, 2.7) == 964
    assert retail_from_wholesale(357) == "$964"


def test_wholesale_columns_not_retail():
    row = (
        '36" X 48" ',
        "Solid",
        4,
        357,
        1006.74,
        481.95,
        1359.09,
        464.1,
        1308.76,
        428.4,
        1208.09,
        517.65,
        1459.77,
    )
    vals = wholesale_values_five_group(row)
    assert vals == [357, 481.95, 464.1, 428.4, 517.65]
    assert retail_from_wholesale(vals[0]) == "$964"
    assert retail_from_wholesale(vals[-1]) == "$1,398"


def test_country_view_oak_published_pdf_36x48_oak_964():
    pdf = Path(__file__).resolve().parents[1] / "sheets" / "Country_View_Oak.pdf"
    if not pdf.is_file():
        pytest.skip("catalog PDF not built")
    import pdfplumber

    found = False
    with pdfplumber.open(pdf) as doc:
        for page in doc.pages:
            for line in (page.extract_text() or "").split("\n"):
                if '36" X 48"' in line and "Solid" in line and "$964" in line:
                    found = True
                    break
    assert found, 'Expected 36" X 48" Solid oak $964 in sheets/Country_View_Oak.pdf'


def test_marked_retail_divided_would_be_wrong_site():
    """Builder retail col (×2.82) must not be used as wholesale."""
    marked = 1006.74
    implied_wholesale = marked / BUILDER_MARKUP
    assert abs(implied_wholesale - 357) < 0.5
    assert retail_whole_dollar(marked, MULTIPLIER_DEFAULT) == 2719


@pytest.mark.skipif(not CVO_XLSX.is_file(), reason="Country View Oak workbook not on disk")
def test_country_view_oak_workbook_detect():
    assert is_country_view_oak_workbook(CVO_XLSX)


@pytest.mark.skipif(not CVO_XLSX.is_file(), reason="Country View Oak workbook not on disk")
def test_country_view_oak_36x48_solid_ci_guard():
    rows, stats = process_country_view_oak_workbook(CVO_XLSX, multiplier=2.7)
    assert stats["wide_rows"] > 500
    spot = spot_check_36x48_solid(rows)
    assert spot["oak"] == "$964"
    assert spot["cherry"] == "$1,302"
    assert spot["qshw"] == "$1,254"
    assert spot["brown_maple_group"] == "$1,157"
    assert spot["walnut"] == "$1,398"