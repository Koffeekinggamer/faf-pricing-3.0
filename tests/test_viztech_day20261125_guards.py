"""Guards for Viztech day20261125 extraction edge cases (no upload dependency)."""

from __future__ import annotations

from scripts.dealer_terms import is_builder_contact_text
from scripts.viztech_book import (
    _is_wholesale_price,
    _prefer_wholesale_sheet_names,
    _row_unf_fin_kind,
)


def test_year_not_treated_as_wholesale_price():
    assert _is_wholesale_price(2026) is False
    assert _is_wholesale_price(2805) is True


def test_integ_unf_fin_row_kinds():
    assert _row_unf_fin_kind([5200, "45 Tv Opening", "unfinished", 2466]) == "unfinished"
    assert _row_unf_fin_kind([50, "finished", 2805, 3298]) == "finished"


def test_prefer_master_over_products_and_markup_twin():
    sheets = {
        "Markup": [[1]],
        "Products": [["SKU"], ["C1"]],
        "Master": [["SKU", "Name", "Finished"], ["C1", "Chest", 564], *[[f"X{i}", "Item", 100 + i] for i in range(12)]],
    }
    skip = _prefer_wholesale_sheet_names(sheets)
    assert "Products" in skip

    ih = {
        "Mar. 1 2026 Price List": [["Oak", 100]],
        "Mar. 1 2026 Price List MARKUP": [["MARKUP", 1]],
    }
    assert "Mar. 1 2026 Price List MARKUP" in _prefer_wholesale_sheet_names(ih)


def test_mailing_address_is_builder_contact():
    assert is_builder_contact_text(
        "Mailing Address: 4254 TR 366 | Millersburg, OH 44654"
    )
