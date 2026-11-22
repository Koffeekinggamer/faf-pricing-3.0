"""Guards for Viztech day20261122 extraction edge cases (no upload dependency)."""

from __future__ import annotations

from pathlib import Path

from scripts.viztech_book import (
    FinishCostLayout,
    _detect_finish_cost_layout,
    _excluded_price_header_cols,
    _finished_price_cols,
    _infer_price_cols,
    _is_option_note_row,
    _transform_finish_cost_row,
    _transform_sealed_or_base_row,
    _unfinished_only_sheet,
    section_title_from_path,
)


def test_unfinished_price_list_banner_detected():
    rows = [
        [None, None, "Wholesale UNFINISHED Price List"],
        [None, None, "Prices effective March 17, 2026"],
    ]
    assert _unfinished_only_sheet(rows)


def test_retail_markup_cols_excluded():
    headers = [
        [
            "MODEL #",
            "DESCRIPTION",
            "WHOLESALE PRICE",
            "Woodgrain/ Brights",
            "SUGGESTED RETAIL",
        ]
    ]
    assert _excluded_price_header_cols(headers, 5) == {4}
    rows = headers + [["POLY1", "Swing", 411, 482, 780]]
    assert _infer_price_cols(rows, 1) == {2, 3}


def test_sealed_painted_preferred_over_price_ea():
    spots: list = []
    cells = _transform_sealed_or_base_row(
        ["C6100", "Rocker", 194, 15, 142, 189, 283],
        sealed_col=5,
        base_col=2,
        multiplier=2.7,
        spots=spots,
        item_hint="",
    )
    assert cells == ["C6100", "$511"]
    assert spots[0].wholesale == 189


def test_finish_cost_compose_adds_stain():
    layout = FinishCostLayout(wood_cols=[9, 10, 11], stain_col=12, paint_col=None, top_col=None)
    spots: list = []
    cells = _transform_finish_cost_row(
        ["#2050", "Console", None, None, None, None, None, None, None, 368, 382, 442, 64],
        layout,
        2.7,
        spots=spots,
        item_hint="",
    )
    assert cells[0] == "#2050"
    assert cells[1:] == ["$1,167", "$1,205", "$1,367"]


def test_finish_cost_layout_maps_gapped_island_top():
    rows = [
        [None, None, "UNFINISHED", None, None, "FINISH", "FINISH", None, "FINISH"],
        [
            "ITEM #",
            "DESCRIPTION",
            "RUSTIC HICKORY\nRED OAK",
            "HARD MAPLE\nCHERRY",
            "QSWO",
            "COST TO STAIN FINISH BASE",
            "COST TO PAINT FINISH BASE",
            None,
            "Island Tops are\nfinished price\nas pictured",
        ],
        ["#9500", "Madison", None, None, None, None, None, None, None, None, 517.14, 525.3, 642.6, 80, 144, 363],
    ]
    layout = _detect_finish_cost_layout(rows, 2)
    assert layout is not None
    assert layout.wood_cols == [10, 11, 12]
    assert layout.stain_col == 13
    assert layout.top_col == 15


def test_option_bullet_rows_skipped():
    assert _is_option_note_row(
        ["• Standard Edge - Eased", None, None, 85],
        "• Standard Edge - Eased",
    )


def test_hershy_section_titles():
    assert (
        section_title_from_path(Path("hershy-way-ltd__2026_Cyp and Treated Pricelist_0126.xlsx"))
        == "Cypress & Treated"
    )
    assert (
        section_title_from_path(Path("hershy-way-ltd__2026_Poly_Aurora_Pricelist_0426.xlsx"))
        == "Poly Aurora"
    )
    assert (
        section_title_from_path(Path("hershy-way-ltd__2026_Poly_Eclipse_Pricelist_0126.xlsx"))
        == "Poly Eclipse"
    )


def test_cost_to_stain_not_finished_product_col():
    headers = [
        ["ITEM #", "DESCRIPTION", "Unfinished", "COST TO STAIN FINISH PRICE", "Stain"],
    ]
    # Stain (Fishers-style) kept; cost-to-stain add-on dropped
    assert 3 not in _finished_price_cols(headers, 5)
    assert 4 in _finished_price_cols(headers, 5)
