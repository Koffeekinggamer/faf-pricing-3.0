"""Country Comfort finished-only retail sheet checks."""

from __future__ import annotations

from pathlib import Path

import pdfplumber
import pytest

from scripts.country_comfort_book import (
    MULTIPLIER,
    process_country_comfort,
)
from scripts.dealer_book import RowKind
from scripts.pricing_lib import retail_whole_dollar

REPO = Path(__file__).resolve().parents[1]
CATALOG = REPO / "sheets" / "Country_Comfort.pdf"
UPLOADS = Path("/home/ubuntu/.cursor/projects/workspace/uploads")
MAIN = UPLOADS / "2026_CC_Prices_Excel_6fb2.xlsx"
LAUREL = UPLOADS / "Laurel_Bay___Brantley_Prices_3d53.xlsx"

SPOT = [
    ("Adena", "6020", 181, 489),
    ("Adena", "6021", 197, 532),
    ("Christy", "9001", 160, 432),
    ("Fairmont", "3191", 193, 522),
    ("Laurel Bay", "5511T", 190, 513),
]


def test_retail_rounding_spot_math():
    for _coll, _sku, wholesale, expected in SPOT:
        assert retail_whole_dollar(wholesale, MULTIPLIER) == expected


def test_flat_addon_retail_math():
    assert retail_whole_dollar(20, MULTIPLIER) == 54
    assert retail_whole_dollar(40, MULTIPLIER) == 108
    assert retail_whole_dollar(5, MULTIPLIER) == 14
    assert retail_whole_dollar(15, MULTIPLIER) == 41


@pytest.mark.skipif(not MAIN.is_file(), reason="Country Comfort source xlsx not available")
def test_processed_spot_prices_and_addons():
    rows, stats = process_country_comfort(MAIN, LAUREL if LAUREL.is_file() else None)
    assert stats["items"] >= 500
    text_rows = [
        (pr.cells[0], pr.cells)
        for pr in rows
        if pr.kind == RowKind.WIDE and pr.cells
    ]
    by_item: dict[str, list[list]] = {}
    for item, cells in text_rows:
        by_item.setdefault(str(item), []).append(cells)

    assert "6020" in by_item
    assert "$489" in by_item["6020"][0]
    assert "6021" in by_item
    assert "$532" in by_item["6021"][0]
    assert any("$432" in c for c in by_item.get("9001", []))
    assert any("$522" in c for c in by_item.get("3191", []))
    assert any("$513" in c for c in by_item.get("5511T", []))

    flat = "\n".join(
        " | ".join(str(x) for x in pr.cells)
        for pr in rows
        if pr.kind in (RowKind.WIDE, RowKind.PROSE)
    )
    assert "Fabric Seat" in flat and "+$54" in flat
    assert "Leather" in flat and "+$108" in flat
    assert "Cherry Seat" in flat and "+$14" in flat
    assert "Footrest" in flat and "+$41" in flat
    assert "Unfin" not in flat


@pytest.mark.skipif(not CATALOG.is_file(), reason="Country Comfort PDF not built")
def test_catalog_pdf_spot_checks():
    with pdfplumber.open(CATALOG) as pdf:
        assert len(pdf.pages) >= 20
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    for price in ("$489", "$532", "$432", "$522", "$513", "+$54", "+$108", "+$14"):
        assert price in text, price
    assert "wholesale" not in text.lower()
    assert "unfin" not in text.lower()
    assert "Preformance" not in text
    assert "Performance" in text
