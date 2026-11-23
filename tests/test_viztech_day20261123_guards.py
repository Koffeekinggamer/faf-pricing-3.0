"""Guards for Viztech day20261123 extraction edge cases (no upload dependency)."""

from __future__ import annotations

from scripts.retail_text import sanitize_retail_text
from scripts.viztech_book import (
    _finished_price_cols,
    _infer_price_cols,
    _row_item_label,
)


def test_unf_fin_abbreviations_are_finished_cols():
    headers = [
        [None, None, None, None, None, "Unf.", "Fin.", "Unf.", "Fin.", "Unf.", "Fin."],
    ]
    assert _finished_price_cols(headers, 11) == {6, 8, 10}


def test_finished_prices_band_excludes_unfinished_band():
    headers = [
        ["DESCRIPTION", None, None, None, None, "FEATURES", None, None, None, None, None, None, "UNFINISHED PRICES", None, None, None, None, "FINISHED PRICES"],
        [None, "Item Code", None, None, None, None, None, None, None, None, None, None, "Oak", "Cherry", "Maple", "QSWO", "Walnut", "Oak", "Cherry", "Maple", "QSWO", "Walnut"],
    ]
    cols = _finished_price_cols(headers, 22)
    assert cols == {17, 18, 19, 20, 21}
    assert 12 not in cols


def test_stained_header_is_finished():
    headers = [["Item #", None, "Description", None, "STAINED", "STAINED"]]
    assert _finished_price_cols(headers, 6) == {4, 5}


def test_side_chair_label_without_item_code():
    row = [None, None, "Side Chair", None, None, None]
    label, idx = _row_item_label(row)
    assert label == "Side Chair"
    assert idx == 2


def test_approx_scrubbed_from_retail_text():
    out = sanitize_retail_text('Antler Chandelier\nApprox. 22" in Diameter', 2.7)
    assert "approx" not in out.lower()
    assert "22" in out


def test_bamms_heuristic_keeps_left_wood_prices():
    rows = [
        ["Item #", "Description", "Dimensions", "Oak", "Cherry", "Maple"],
        ["ASH2", "2 Door", '48"W', 1140, 1254, 1537],
        ["ASH3", "3 Door", '64"W', 1401, 1539, 1891],
    ]
    cols = _infer_price_cols(rows, 1)
    assert cols == {3, 4, 5}
