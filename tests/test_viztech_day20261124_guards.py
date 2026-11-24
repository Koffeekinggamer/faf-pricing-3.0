"""Guards for Viztech day20261124 extraction edge cases (no upload dependency)."""

from __future__ import annotations

from scripts.viztech_book import (
    _markup_calculator_without_finished,
    _prefer_wholesale_sheet_names,
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
