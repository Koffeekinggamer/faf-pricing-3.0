"""Country Comfort finished-only retail sheet checks.

Tests must never depend on uploaded source workbooks (CI has no access to
/home/ubuntu/.cursor/projects/workspace/uploads). Exercise the committed PDF
and pure retail math only; optional live-xlsx checks are OSError-safe skips.
"""

from __future__ import annotations

from pathlib import Path

import pdfplumber
import pytest

from scripts.country_comfort_book import MULTIPLIER
from scripts.pricing_lib import retail_whole_dollar

REPO = Path(__file__).resolve().parents[1]
CATALOG = REPO / "sheets" / "Country_Comfort.pdf"

SPOT = [
    ("Adena", "6020", 181, 489),
    ("Adena Hard Maple", "6020", 249, 673),
    ("Adena", "6021", 197, 532),
    ("Christy", "9001", 160, 432),
    ("Fairmont", "3191", 193, 522),
    ("Laurel Bay", "5511T", 190, 513),
    ("Cambria Regular", "2510", 203, 549),
    ("Cambria Performance", "2510", 213, 576),
    ("Cambria Leather", "2510", 243, 657),
    ("Laurel Bay table", "3636 LAB 00", 507, 1369),
]


def test_retail_rounding_spot_math():
    for _coll, _sku, wholesale, expected in SPOT:
        assert retail_whole_dollar(wholesale, MULTIPLIER) == expected


def test_flat_addon_retail_math():
    assert retail_whole_dollar(20, MULTIPLIER) == 54
    assert retail_whole_dollar(40, MULTIPLIER) == 108
    assert retail_whole_dollar(5, MULTIPLIER) == 14
    assert retail_whole_dollar(15, MULTIPLIER) == 41


# Additional Table Options workbook dollars → ceil(×2.7) exactly once
OPTION_DOLLARS = [
    ("36\" or 42\" High", 50, 135),
    ("2 - Tone Stained or Paint", 80, 216),
    ("2 - Tone Paint and Glaze", 130, 351),
    ("1 - Color Paint Whole Table", 130, 351),
    ("Hand Planed", 180, 486),
    ("Ball Bearing Slides", 100, 270),
]


def test_additional_table_option_dollars_once():
    for _label, wholesale, expected in OPTION_DOLLARS:
        assert retail_whole_dollar(wholesale, MULTIPLIER) == expected


def _safe_xlsx_available() -> bool:
    """Return True only when local upload workbooks are readable (never raises)."""
    try:
        main = Path("/home/ubuntu/.cursor/projects/workspace/uploads/2026_CC_Prices_Excel_6fb2.xlsx")
        return main.is_file()
    except OSError:
        return False


@pytest.mark.skipif(not _safe_xlsx_available(), reason="Country Comfort source xlsx not available")
def test_processed_spot_prices_and_addons():
    from scripts.country_comfort_book import process_country_comfort
    from scripts.dealer_book import RowKind

    main = Path("/home/ubuntu/.cursor/projects/workspace/uploads/2026_CC_Prices_Excel_6fb2.xlsx")
    laurel = Path("/home/ubuntu/.cursor/projects/workspace/uploads/Laurel_Bay___Brantley_Prices_3d53.xlsx")
    try:
        laurel_arg = laurel if laurel.is_file() else None
    except OSError:
        laurel_arg = None
    rows, stats = process_country_comfort(main, laurel_arg)
    assert stats["items"] >= 500
    by_item: dict[str, list[list]] = {}
    for pr in rows:
        if pr.kind == RowKind.WIDE and pr.cells:
            by_item.setdefault(str(pr.cells[0]), []).append(pr.cells)

    assert "6020" in by_item
    assert "$489" in by_item["6020"][0]
    assert "$673" in by_item["6020"][0]  # Hard Maple L10=249
    assert "6021" in by_item
    assert "$532" in by_item["6021"][0]
    assert any("$432" in c for c in by_item.get("9001", []))
    assert any("$522" in c for c in by_item.get("3191", []))
    assert any("$513" in c for c in by_item.get("5511T", []))

    # Cambria 2510 mid-sheet block
    cambria = [c for c in by_item.get("2510", [])]
    assert cambria
    flat_c = "\n".join(" | ".join(str(x) for x in c) for c in cambria)
    assert "$549" in flat_c and "$576" in flat_c and "$657" in flat_c

    # Vanessa must not treat paired-collection order # as a price
    van = by_item.get("3311", [])
    assert van
    assert "$2,757" not in van[0]
    assert "$2,760" not in "\n".join(str(x) for row in van for x in row)

    # Laurel Bay tables from Sheet2
    assert "3636 LAB 00" in by_item
    assert "$1,369" in by_item["3636 LAB 00"][0]

    flat = "\n".join(
        " | ".join(str(x) for x in pr.cells)
        for pr in rows
        if pr.kind in (RowKind.WIDE, RowKind.PROSE, RowKind.SUBHEAD, RowKind.SECTION)
    )
    assert "Fabric Seat" in flat and "+$54" in flat
    assert "Leather" in flat and "+$108" in flat
    assert "Footrest" in flat and "+$41" in flat
    assert "Unfin" not in flat

    # Additional Table Options: ×2.7 exactly once
    for label, wholesale, expected in OPTION_DOLLARS:
        assert retail_whole_dollar(wholesale, MULTIPLIER) == expected
        needle = f"+${expected:,}"
        assert needle in flat, f"{label} should be {needle}"
    assert "+$365" not in flat and "+$584" not in flat and "+$948" not in flat

    # Wood-grade lists printed in full
    assert "Free Options" in flat
    assert "Red Oak" in flat and "Sap Cherry" in flat
    assert "Superior Hardwoods" in flat
    assert "Premium Hardwoods" in flat
    assert "Distressed = Free" in flat

    # Main-book table configs
    assert stats.get("table_configs", 0) >= 271


def _count_pdf_table_configs(text: str) -> int:
    """Count priced size×size table rows rendered in the PDF."""
    import re

    n = 0
    for line in text.splitlines():
        if re.match(r"^\d+x\d+\s+.+\s+\d+\s+\$", line.strip(), re.I):
            n += 1
    return n


@pytest.mark.skipif(not CATALOG.is_file(), reason="Country Comfort PDF not built")
def test_catalog_pdf_spot_checks():
    with pdfplumber.open(CATALOG) as pdf:
        assert len(pdf.pages) >= 20
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    for price in (
        "$489",
        "$532",
        "$432",
        "$522",
        "$513",
        "$673",
        "$549",
        "$576",
        "$657",
        "$1,369",
        "+$54",
        "+$108",
        "+$135",
        "+$216",
        "+$351",
        "+$486",
        "+$270",
        "$948",
        "$1,165",
    ):
        assert price in text, price
    assert "wholesale" not in text.lower()
    assert "unfin" not in text.lower()
    assert "Preformance" not in text
    assert "Performance" in text
    assert "+$365" not in text
    # Vanessa must not price paired-collection order # 1021 as $2,757
    assert "3311" in text
    for page_text in ((p.extract_text() or "") for p in pdfplumber.open(CATALOG).pages):
        if "3311" in page_text:
            assert "$2,757" not in page_text
            break
    assert "3636 LAB 00" in text or "3636 LAB" in text
    assert "12x50" in text
    assert "Free Options" in text
    assert "Standard Hardwoods" in text
    assert "Distressed = Free" in text
    assert "Description" in text
    assert "Descriptio\nn" not in text
    # All main-book table configurations must render
    assert _count_pdf_table_configs(text) >= 271

