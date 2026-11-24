"""Guards for Viztech day20261124 extraction edge cases (no upload dependency)."""

from __future__ import annotations

from pathlib import Path

from scripts.pdf_quality_check import find_call_for_contact_wording
from scripts.retail_text import (
    is_call_for_pricing_cell,
    sanitize_retail_text,
    strip_call_for_contact_phrases,
)
from scripts.viztech_book import (
    _markup_calculator_without_finished,
    _prefer_wholesale_sheet_names,
    _transform_row_prices,
)


def test_prefer_skips_prices_with_markup_when_plain_prices_exists():
    sheets = {
        "2026 Prices with Markup": [["ITEM #", "markup"], ["901", 1.0]],
        "2026 Prices": [["ITEM #"], ["901", 315.0]],
    }
    assert _prefer_wholesale_sheet_names(sheets) == {"2026 Prices with Markup"}


def test_markup_calculator_without_finished_detected():
    rows = [
        ["Markup Calculator", 0, "Enter percent of markup", 0, "Cost of Finishing (optional)"],
        ["Brown Maple", "Brown Maple", "Cherry", "Cherry"],
        ["38 Mission Side Chair", 81, 101],
    ]
    assert _markup_calculator_without_finished(rows) is True


def test_finished_grid_not_treated_as_markup_calculator():
    rows = [
        ["Item #", "Description", "Oak Unfinished", "Oak Finished"],
        ["801", "Bookcase", 119.7, 156.7],
    ]
    assert _markup_calculator_without_finished(rows) is False


def test_strip_call_for_parenthetical_keeps_product_label():
    assert (
        strip_call_for_contact_phrases(
            "Rectangle Glass Tops: (Call for oval/round pricing)"
        )
        == "Rectangle Glass Tops:"
    )


def test_strip_special_order_call_for_sentence():
    assert (
        strip_call_for_contact_phrases(
            "Special Order options are available for most items.  "
            "Please call for information and quote."
        )
        == ""
    )


def test_call_for_pricing_cell_left_blank_in_price_row():
    assert is_call_for_pricing_cell("Call for Pricing")
    cells = _transform_row_prices(
        ["HC10126", "Chair", "Upholstered", 250, "Call for Pricing"],
        2.7,
        [3, 4],
        spots=[],
        item_hint="",
    )
    assert cells == ["HC10126", "$675"]
    assert "Call" not in " ".join(cells)


def test_sanitize_drops_call_for_pricing_cell():
    assert sanitize_retail_text("Call for Pricing", 2.7) == ""


def test_day20261124_pdfs_have_no_call_for_contact_wording():
    sheets = Path(__file__).resolve().parent.parent / "sheets"
    for name in (
        "Hoosier_Home_Furniture.pdf",
        "Hoosier_Poly.pdf",
        "Hope_Works_Creations.pdf",
        "HopeWood_Inc.pdf",
    ):
        p = sheets / name
        if not p.is_file():
            continue
        bad = find_call_for_contact_wording(p)
        assert not bad, "\n".join(bad)
