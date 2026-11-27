"""Guards for Viztech day20261127 extraction edge cases (no upload dependency)."""

from __future__ import annotations

from scripts.dealer_book import RowKind
from scripts.pricing_lib import retail_whole_dollar
from pathlib import Path

from scripts.viztech_book import (
    _detect_finished_add_layout,
    _drop_row,
    _finished_price_cols,
    _infer_price_cols,
    _keep_item_mid_text,
    _process_sheet,
    _transform_finished_add_row,
    section_title_from_path,
)


def test_kkt_finished_add_composes_unfinished_plus_add():
    rows = [
        ["Item #", "ADELLA", None, "Oak", "Brown Maple", "Cherry", "QSWO", "Finished\nADD"],
        [None, None, None, "Unfinished", "Unfinished", "Unfinished", "Unfinished", "Add", "Finished"],
        [1010, "Side Chair", None, 124, 138, 147, 170, 36, 0],
        [1012, "Arm Chair", None, 144, 162, 171, 199, 36, 0],
    ]
    layout = _detect_finished_add_layout(rows, 2)
    assert layout is not None
    assert layout.wood_cols == [3, 4, 5, 6]
    assert layout.add_col == 7
    cells = _transform_finished_add_row(
        rows[2], layout, 2.7, spots=[], item_hint=""
    )
    # 124+36=160→$432, 138+36=174→$470, 147+36=183→$495, 170+36=206→$557
    assert cells == ["1010", "Side Chair", "$432", "$470", "$495", "$557"]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert wide["1010"][2:] == ["$432", "$470", "$495", "$557"]
    assert wide["1012"][1] == "Arm Chair"


def test_kkt_fin_unf_pairs_drop_mirrored_reprint():
    headers = [
        ["Item Number", "Description", "Oak", None, "Maple", None, "Cherry", None, "QSWO", None, "Description", "Oak"],
        [None, None, "Finished", "Unfinished", "Finished", "Unfinished", "Finished", "Unfinished", "Finished", "Unfinished", None, "Finished", "Unfinished"],
    ]
    fin = _finished_price_cols(headers, 14)
    assert fin == {2, 4, 6, 8}
    assert 11 not in fin


def test_ln_estimated_finished_band_not_unfinished():
    headers = [
        [None, None, None, "UNFINISHED", None, None, None, None, "ESTIMATED FINISHED PRICE"],
        [
            "Item #",
            "Description",
            "Size",
            "Oak",
            "Cherry",
            "Walnut",
            "R. Cherry Br. Maple",
            "R. Walnut",
            "Oak",
            "Cherry",
            "Walnut",
            "R. Cherry Br. Maple",
            "R. Walnut",
            "Estimated Finish Per Piece",
        ],
    ]
    fin = _finished_price_cols(headers, 14)
    assert fin == {8, 9, 10, 11, 12}
    assert 3 not in fin
    assert 13 not in fin
    rows = headers + [
        ["AD7284", "Corner Desk", "72x36x31", 2081, 2705.3, 4057.95, 2351.53, 3516.89, 2259, 2883.3, 4235.95, 2529.53, 3694.89, 178],
    ]
    cols = _infer_price_cols(rows, 2)
    assert cols == {8, 9, 10, 11, 12}
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = [r.cells for r in out if r.kind == RowKind.WIDE][0]
    assert wide[0] == "AD7284"
    assert "$6,100" in wide and "$5,619" not in wide  # finished Oak, not unfinished


def test_la_patio_poly_columns_all_finished():
    headers = [
        ["CHARITY TABLES", None, "HEIGHT", "STANDARD \nPOLY", "2-TONE\nstd/woodgrain", "WOODGRAIN\nPOLY"],
    ]
    fin = _finished_price_cols(headers, 6)
    assert fin == {3, 4, 5}
    rows = headers + [
        ["LACT36R-D", '36" round Charity Table', "dining", 505.0, 535.0, 555.0],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    cells = [r.cells for r in out if r.kind == RowKind.WIDE][0]
    assert cells[0] == "LACT36R-D"
    assert cells[-3:] == ["$1,364", "$1,445", "$1,499"]


def test_comfort_fabric_leather_grades_kept():
    rows = [
        [None, None, None, None, None, None, None, None, None, None, None, None, "Heartland\nFabric"],
        [
            None,
            "Electric Controls",
            None,
            None,
            "Swivel",
            "Lay Flat",
            None,
            None,
            "Standard\nFabric",
            "Premium \nFabric",
            "Ultra \nLeather",
            "Italian Cowhide Leather",
        ],
        [
            '34" Charleston Stationary Chair, \ntacks add $95',
            "N/A",
            "N/A",
            "N/A",
            None,
            None,
            None,
            None,
            565.0,
            595.0,
            755.0,
            1355.0,
            755.0,
        ],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = [r.cells for r in out if r.kind == RowKind.WIDE][0]
    assert "Standard" not in wide  # swivel option token dropped
    assert "$1,526" in wide and "$1,607" in wide and "$2,039" in wide and "$3,659" in wide
    assert any("tacks add $257" in c for c in wide)


def test_chimes_part_and_size_kept():
    rows = [
        ["Part #", "Description", "Approx\nSize", "Qty.", "Price"],
        [904.0, "Sonic Sound Waves - Jade", '56"', None, 149.0],
        [841.0, "Big Ben - Terra", '82"', None, 232.0],
    ]
    out = _process_sheet("Order Form", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert wide["904"] == ["904", "Sonic Sound Waves - Jade", '56"', "$403"]
    assert wide["841"][-1] == f"${retail_whole_dollar(232.0, 2.7):,}"


def test_discontinued_and_password_rows_dropped():
    assert _drop_row(["Zach Chair is discontinued for 2027"], "Zach Chair is discontinued for 2027")
    assert _drop_row(
        ["TO ACCESS DIGITAL FILES", "www.example.com  :  PASSWORD: secret"],
        "PASSWORD: secret",
    )
    assert not _keep_item_mid_text("Chair")
    assert not _keep_item_mid_text("Standard")


def test_day20261127_section_titles_from_slug():
    assert section_title_from_path(
        Path("ln-woodworking__LNWW_2026_Bookcase_Pricelist_1026.xlsx")
    )
    # Main KKT file keeps a sensible catalog title
    assert "price" in section_title_from_path(
        Path("kountry-knob-tables__KKT_Pricelist_021026.xlsx")
    ).casefold() or section_title_from_path(
        Path("kountry-knob-tables__KKT_Pricelist_021026.xlsx")
    )
