"""Guards for Viztech day20261126 extraction edge cases (no upload dependency)."""

from __future__ import annotations

from pathlib import Path

from scripts.viztech_book import (
    ADD_ON_PRICE_COL_RE,
    _addon_cols_above,
    _excluded_price_header_cols,
    _finished_price_cols,
    _infer_price_cols,
    section_title_from_path,
)


def test_kidron_style_section_from_slug_suffix():
    assert section_title_from_path(Path("kidron-woodcraft__Flint_Ridge.xlsx")) == (
        "Flint Ridge"
    )
    assert section_title_from_path(
        Path("kidron-woodcraft__Timberline_Entertainment_Center.xlsx")
    ) == "Timberline Entertainment Center"
    assert section_title_from_path(Path("kidron-woodcraft__Solo Galaxy.xlsx")) == (
        "Solo Galaxy"
    )
    # Standard Viztech coded names still use the Pricelist parser / fallback
    assert section_title_from_path(
        Path("jm-woodworking__JMW_2026_Pricelist_0426.xlsx")
    ) == "Price list"


def test_finished_prices_add_excluded_from_finished_cols():
    headers = [
        [
            "Item No.",
            "Description",
            "Size",
            "Oak",
            "Br. Maple",
            "Finished\nPrices\nAdd\n",
        ]
    ]
    fin = _finished_price_cols(headers, 6)
    assert 5 not in fin
    excl = _excluded_price_header_cols(headers, 6)
    assert 5 in excl


def test_add_for_two_tone_excluded_via_lookback():
    rows = [
        ["Oak", "QSWO", "Brown Maple", "Hickory", "Add for 2-Tone\nStain\norders:"],
        ["Item #", "Description", "", "", ""],
        [900, "Allison Dresser", 954.8, 1184.1, 31.88],
    ]
    assert _addon_cols_above(rows, 2, 5) == {4}
    cols = _infer_price_cols(rows, 2)
    assert 4 not in cols
    assert 2 in cols or 3 in cols


def test_unfin_header_recognized():
    headers = [["Item", "Fin", "Unfin", "Fin", "Unfin"]]
    fin = _finished_price_cols(headers, 5)
    assert fin == {1, 3}


def test_addon_regex_covers_jericho_and_kohler():
    assert ADD_ON_PRICE_COL_RE.search("Finished\nPrices\nAdd")
    assert ADD_ON_PRICE_COL_RE.search("Add for          2-Tone\nStain\norders:")
    assert not ADD_ON_PRICE_COL_RE.search("Finished Prices")


def test_finished_prices_add_excludes_trailing_noise_cols():
    rows = [
        ["Item", "Desc", "Size", "Oak", "Finished\nPrices\nAdd"],
        ["M39", "Table", "12x22", 172, 34, None, None, 3],
    ]
    bad = _addon_cols_above(rows, 1, 8)
    assert 4 in bad and 5 in bad and 7 in bad
    cols = _infer_price_cols(rows, 1)
    assert 3 in cols
    assert 4 not in cols and 7 not in cols
