"""Guards for Viztech day20261127 extraction edge cases (no upload dependency)."""

from __future__ import annotations

import re
from pathlib import Path

from scripts.dealer_book import RowKind
from scripts.pricing_lib import retail_whole_dollar
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


def _dollar_tokens(cells: list[str]) -> list[str]:
    return [c for c in cells if isinstance(c, str) and c.startswith("$")]


def _r(wholesale: float) -> str:
    """Expected whole-dollar retail for look-alike wholesale fixtures (multiplier 2.7)."""
    return f"${retail_whole_dollar(wholesale, 2.7):,}"


def _assert_no_bare_yearish_prices(cells: list[str]) -> None:
    """Price rows must not leak bare 4-digit wholesale amounts (2015–2035)."""
    joined = " ".join(cells)
    for m in re.finditer(r"(?<![\d,$])(20[1-3]\d)(?![\d])", joined):
        assert False, f"bare yearish wholesale leaked: {m.group(0)} in {cells}"


def test_kkt_finished_add_composes_unfinished_plus_add():
    add = 11.0
    bases = (101.0, 102.0, 103.0, 104.0)
    rows = [
        ["Item #", "ADELLA", None, "Oak", "Brown Maple", "Cherry", "QSWO", "Finished\nADD"],
        [None, None, None, "Unfinished", "Unfinished", "Unfinished", "Unfinished", "Add", "Finished"],
        [1010, "Side Chair", None, *bases, add, 0],
        [1012, "Arm Chair", None, 111, 112, 113, 114, add, 0],
    ]
    layout = _detect_finished_add_layout(rows, 2)
    assert layout is not None
    assert layout.wood_cols == [3, 4, 5, 6]
    assert layout.add_col == 7
    cells = _transform_finished_add_row(
        rows[2], layout, 2.7, spots=[], item_hint=""
    )
    assert cells == [
        "1010",
        "Side Chair",
        *[_r(b + add) for b in bases],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert wide["1010"][2:] == [_r(b + add) for b in bases]
    assert wide["1012"][1] == "Arm Chair"


def test_kkt_alamo_bare_family_header_composes_add():
    """ALAMO sections omit wood/ADD labels — persist Finished ADD from above."""
    add = 11.0
    rows = [
        ["Item #", None, "Oak", "Brown Maple", "Cherry", "QSWO", "Finished\nADD"],
        [None, None, "Unfinished", "Unfinished", "Unfinished", "Unfinished", "Add"],
        [1010, "Side Chair", 91, 92, 93, 94, add],
        ["Item #", "ALAMO"],
        [1014, "Arm Chair", 95, 96, 97, 98, add],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["1014"]) == [_r(b + add) for b in (95, 96, 97, 98)]
    _assert_no_bare_yearish_prices(wide["1014"])


def test_kkt_lexi_blank_first_wood_shows_remaining():
    add = 11.0
    rows = [
        ["Item #", "LEXI", None, "Oak", "Maple", "Cherry", "Finished\nADD"],
        [None, None, None, "Unfinished", "Unfinished", "Unfinished", "Add"],
        [2044, "Side", None, 88, 92, 99, add],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["2044"]) == [_r(88 + add), _r(92 + add), _r(99 + add)]


def test_kkt_fin_unf_pairs_drop_mirrored_reprint():
    headers = [
        ["Item Number", "Description", "Oak", None, "Maple", None, "Cherry", None, "QSWO", None, "Description", "Oak"],
        [None, None, "Finished", "Unfinished", "Finished", "Unfinished", "Finished", "Unfinished", "Finished", "Unfinished", None, "Finished", "Unfinished"],
    ]
    fin = _finished_price_cols(headers, 14)
    assert fin == {2, 4, 6, 8}
    assert 11 not in fin


def test_kkt_greenville_drops_mirrored_reprint():
    """Two-wood left panel with layout twin — print finished woods once."""
    headers = [
        [
            "Item Number",
            "Description",
            "Rustic Maple, Rustic White Oak",
            None,
            "Brown Maple Wormy Maple",
            None,
            "Description",
            "Elm/Cherry/Hickory/Hard Maple",
            None,
            "R. Cherry/Sap Cherry/R. QSWO",
            None,
        ],
        [
            None,
            None,
            "Finished",
            "Unfinished",
            "Finished",
            "Unfinished",
            None,
            "Finished",
            "Unfinished",
            "Finished",
            "Unfinished",
        ],
    ]
    fin = _finished_price_cols(headers, 12)
    assert fin == {2, 4}
    fin_a, fin_b = 555.0, 444.0
    rows = headers + [
        [
            "1803-4266",
            "Solid Top",
            fin_a,
            111.0,
            fin_b,
            222.0,
            "Solid Top",
            fin_a,
            111.0,
            fin_b,
            222.0,
        ],
    ]
    out = _process_sheet("Greenville", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["1803-4266"]) == [_r(fin_a), _r(fin_b)]


def test_kkt_raleigh_full_left_panel_drops_reprint():
    headers = [
        [
            "Item Number",
            "Description",
            "R. Cherry Sap Cherry Oak",
            None,
            "R. Hickory Brown Maple",
            None,
            "Elm Cherry Hard Maple",
            None,
            "QWSO",
            None,
            "Description",
            "Oak",
            None,
            "Elm/Cherry/Hickory",
            None,
            "R. Cherry group",
            None,
            "QSWO",
        ],
        [
            None,
            None,
            "Finished",
            "Unfinished",
            "Finished",
            "Unfinished",
            "Finished",
            "Unfinished",
            "Finished",
            "Unfinished",
            None,
            "Finished",
            "Unfinished",
            "Finished",
            "Unfinished",
            "Finished",
            "Unfinished",
            "Finished",
            "Unfinished",
        ],
    ]
    fin = _finished_price_cols(headers, 20)
    assert fin == {2, 4, 6, 8}
    assert 11 not in fin


def test_kkt_quick_ship_finished_not_unfinished():
    oregon_fin, oregon_unf = 194.0, 150.0
    gal_side_fin, gal_side_unf = 185.0, 142.0
    gal_arm_fin, gal_arm_unf = 210.0, 160.0
    rows = [
        ["Description", "Elm\nBrown Maple", None, "Rustic Cherry", None, "Rustic Hickory"],
        [None, "Finished", "Unfinished", "Finished", "Unfinished", "Finished", "Unfinished", None, "Finished", "Unfinished", "Finished", "Unfinished"],
        ["Oregon Arm Chair", oregon_fin, oregon_unf, oregon_fin, oregon_unf, None, None, None, oregon_fin, oregon_unf, oregon_fin, oregon_unf],
        ["Galveston G-2 Side Chair", gal_side_fin, gal_side_unf, None, None, None, None, None, gal_side_fin, gal_side_unf],
        ["Galveston G-2 Arm Chair", gal_arm_fin, gal_arm_unf, None, None, None, None, None, gal_arm_fin, gal_arm_unf],
        ["Sierra DP Table", 356.0, None, 310.0, None, None, None, None, 356.0, None, 310.0, None],
        ["2 Tone", None, 12],
        ["Paint", None, 12],
    ]
    out = _process_sheet("Quick Ship", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["Oregon Arm Chair"]) == [_r(oregon_fin), _r(oregon_fin)]
    assert _dollar_tokens(wide["Galveston G-2 Side Chair"]) == [_r(gal_side_fin)]
    assert _dollar_tokens(wide["Galveston G-2 Arm Chair"]) == [_r(gal_arm_fin)]
    assert _dollar_tokens(wide["Sierra DP Table"]) == [_r(356.0), _r(310.0)]


def test_kkt_options_thickness_two_tone_and_no_digital_prompt():
    rows = [
        [None, None, None, None, None, None, None, None, "TO ACCESS DIGITAL FILES, GO TO:"],
        [None, None, None, None, None, None, None, None, "www.example.com : PASSWORD: secret"],
        ["TOPS"],
        ['For 1¼" tops add 15%'],
        ['For 1 ½" and 1 ¾" tops add 30%'],
        ["FOR TWO TONE"],
        ["2-tone Stain : Add $60.00"],
        ["2-tone Paint and Stain : Add $70.00"],
        ["2-tone Paint : Add $80.00"],
        ["For solid paint color add $60.00 to finished price"],
        ["2-tone stain for benches add $30"],
        ["Prices effective January 1, 2019"],
    ]
    out = _process_sheet("Options", rows, multiplier=2.7, spots=[])
    texts = [" ".join(r.cells) for r in out]
    blob = "\n".join(texts)
    assert "DIGITAL FILES" not in blob.upper()
    assert "PASSWORD" not in blob.upper()
    assert "Prices effective" not in blob
    assert any('1¼" tops add 15%' in t for t in texts)
    assert any("1 ¾" in t and "30%" in t for t in texts)
    assert any("2-tone Stain : Add $162" in t for t in texts)
    assert any("2-tone Paint and Stain : Add $189" in t for t in texts)
    assert any("2-tone Paint : Add $216" in t for t in texts)
    assert any("benches add $81" in t for t in texts)


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
    fin = (901.0, 902.0, 903.0, 904.0, 905.0)
    rows = headers + [
        ["AD7284", "Corner Desk", "72x36x31", 801, 802, 803, 804, 805, *fin, 55],
    ]
    cols = _infer_price_cols(rows, 2)
    assert cols == {8, 9, 10, 11, 12}
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = [r.cells for r in out if r.kind == RowKind.WIDE][0]
    assert wide[0] == "AD7284"
    assert _r(fin[0]) in wide and _r(801) not in wide
    assert _r(55) not in wide


def test_ln_estimated_finish_price_alias_excluded():
    """Bookcase sheets label the helper column 'estimated finish price'."""
    headers = [
        [None, None, None, "UNFINISHED", None, None, None, None, "ESTIMATED FINISHED PRICE"],
        [
            "Item #",
            "Description",
            "Adj. Shelves",
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
            "estimated finish price",
        ],
    ]
    fin = _finished_price_cols(headers, 14)
    assert fin == {8, 9, 10, 11, 12}
    assert 13 not in fin
    ad3660_fin = (801.0, 802.0, 803.0, 804.0, 805.0)
    fob_fin = (591.0, 592.0, 593.0, 594.0, 595.0)
    afd24_fin = (851.0, 852.0, 853.0, 854.0, 855.0)
    rows = headers + [
        ["AD3660-O", "Open Bookcase", 3, 701, 702, 703, 704, 705, *ad3660_fin, 44],
        ["FOB-3636-O", "Bookcase", None, 501, 502, 503, 504, 505, *fob_fin],
        ["AFD24", "File Desk", "74x24x31", 701, 702, 703, 704, 705, *afd24_fin, 44],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["AD3660-O"]) == [_r(v) for v in ad3660_fin]
    assert _r(3) not in wide["AD3660-O"]
    assert _r(44) not in wide["AD3660-O"]
    assert wide["FOB-3636-O"][0] == "FOB-3636-O"
    assert len(_dollar_tokens(wide["FOB-3636-O"])) == 5
    assert _dollar_tokens(wide["FOB-3636-O"]) == [_r(v) for v in fob_fin]
    assert _dollar_tokens(wide["AFD24"]) == [_r(v) for v in afd24_fin]
    _assert_no_bare_yearish_prices(wide["AFD24"])


def test_la_patio_poly_columns_all_finished():
    headers = [
        ["CHARITY TABLES", None, "HEIGHT", "STANDARD \nPOLY", "2-TONE\nstd/woodgrain", "WOODGRAIN\nPOLY"],
    ]
    fin = _finished_price_cols(headers, 6)
    assert fin == {3, 4, 5}
    poly = (501.0, 502.0, 503.0)
    rows = headers + [
        ["LACT36R-D", '36" round Charity Table', "dining", *poly],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    cells = [r.cells for r in out if r.kind == RowKind.WIDE][0]
    assert cells[0] == "LACT36R-D"
    assert cells[-3:] == [_r(v) for v in poly]


def test_la_patio_wood_grain_spaced_and_yearish_and_phone_holder():
    headers = [
        ["JADE TABLES", None, None, "STANDARD \nPOLY", "2-TONE\nstd/wood grain", "WOOD GRAIN \nPOLY"],
    ]
    assert _finished_price_cols(headers, 6) == {3, 4, 5}
    jade = (701.0, 702.0, 703.0)
    holder = (26.0, 27.0, 27.0)
    rows = headers + [
        ["LAJA4677F-B", '46" x 77" Jade Table w/Fire Ring', "bar", *jade],
        ["LACH555", "Screw-On Cup/Phone Holder", "", *holder],
        ["Zach Chair is discontinued for 2027"],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["LAJA4677F-B"]) == [_r(v) for v in jade]
    _assert_no_bare_yearish_prices(wide["LAJA4677F-B"])
    assert wide["LACH555"][0] == "LACH555"
    assert _dollar_tokens(wide["LACH555"]) == [_r(v) for v in holder]
    assert not any("Zach" in " ".join(r.cells) for r in out)


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
            551.0,
            552.0,
            553.0,
            554.0,
            555.0,
        ],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = [r.cells for r in out if r.kind == RowKind.WIDE][0]
    assert "Standard" not in wide  # swivel option token dropped
    assert all(_r(v) in wide for v in (551, 552, 553, 554, 555))
    assert any("tacks add $257" in c for c in wide)


def test_comfort_dutch_boy_and_yearish_grades():
    rows = [
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
            "Heartland\nFabric",
        ],
        [
            '37" Dutch Boy Wallhugger Recliner',
            125.0,
            "N/A",
            25.0,
            75.0,
            95.0,
            None,
            None,
            631.0,
            632.0,
            633.0,
            634.0,
            635.0,
        ],
        [
            '52" E-Z Sleeper Jackknife w/Drawer',
            "N/A",
            "N/A",
            "N/A",
            None,
            None,
            None,
            None,
            571.0,
            572.0,
            573.0,
            574.0,
            575.0,
        ],
        [
            "Dutch Swivel Diner",
            "N/A",
            "N/A",
            "N/A",
            None,
            None,
            None,
            None,
            341.0,
            342.0,
            343.0,
            344.0,
            345.0,
        ],
        [
            '83" Encore Zerowall Sofa WITHOUT Folddown',
            "Standard",
            255.0,
            "N/A",
            None,
            None,
            None,
            "Standard",
            161.0,
            162.0,
            163.0,
            164.0,
            165.0,
        ],
        [
            '62" Lazy Lounger Wallhugger Workmate',
            255.0,
            "N/A",
            55.0,
            None,
            "N/A",
            "Standard",
            85.0,
            171.0,
            172.0,
            173.0,
            174.0,
            175.0,
        ],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    by_name = {
        r.cells[0]: r.cells for r in out if r.kind == RowKind.WIDE
    }
    dutch = by_name['37" Dutch Boy Wallhugger Recliner']
    assert _dollar_tokens(dutch) == [
        _r(125),
        _r(25),
        _r(75),
        _r(95),
        *[_r(v) for v in (631, 632, 633, 634, 635)],
    ]
    ez = by_name['52" E-Z Sleeper Jackknife w/Drawer']
    assert _dollar_tokens(ez) == [_r(v) for v in (571, 572, 573, 574, 575)]
    diner = by_name["Dutch Swivel Diner"]
    assert _dollar_tokens(diner) == [_r(v) for v in (341, 342, 343, 344, 345)]
    encore = [c for k, c in by_name.items() if "Encore" in k][0]
    assert _r(163) in encore
    _assert_no_bare_yearish_prices(encore)
    lazy = [c for k, c in by_name.items() if "Lazy Lounger" in k][0]
    assert _r(173) in lazy
    _assert_no_bare_yearish_prices(lazy)


def test_chimes_part_and_size_kept():
    rows = [
        ["Part #", "Description", "Approx\nSize", "Qty.", "Price"],
        [904.0, "Sonic Sound Waves - Jade", '56"', None, 51.0],
        [841.0, "Big Ben - Terra", '82"', None, 82.0],
    ]
    out = _process_sheet("Order Form", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert wide["904"] == ["904", "Sonic Sound Waves - Jade", '56"', _r(51.0)]
    assert wide["841"][-1] == _r(82.0)


def test_chimes_phone_sidebar_keeps_820_and_818():
    rows = [
        ["Part #", "Description", "Approx\nSize", "Qty.", "Price", None, "PHONE"],
        [820.0, "Sparkling Brooke - Tex Black", '30"', None, 44.0, None, "PHONE"],
        [818.0, "Zephyr Valley - Tex Black", '37"', None, 50.0, None, "FAX"],
    ]
    assert not _drop_row(rows[1], "820 Sparkling Brooke PHONE")
    assert not _drop_row(rows[2], "818 Zephyr Valley FAX")
    out = _process_sheet("Order Form", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert wide["820"][-1] == _r(44.0)
    assert wide["818"][-1] == _r(50.0)


def test_discontinued_and_password_rows_dropped():
    assert _drop_row(["Zach Chair is discontinued for 2027"], "Zach Chair is discontinued for 2027")
    assert _drop_row(
        ["TO ACCESS DIGITAL FILES", "www.example.com  :  PASSWORD: secret"],
        "PASSWORD: secret",
    )
    assert not _keep_item_mid_text("Chair")
    assert not _keep_item_mid_text("Standard")


def test_kkt_denali_finished_only_five_woods():
    rows = [
        ["Item #", "DENALI", None, "Oak", "Cherry", "QSWO", "Walnut", None],
        ["Only Available in Finished", None, None, "Finished", "Finished", "Finished", "Finished", None],
        [1032, "Side Chair", None, 111, 112, 113, 114, None],
        [1034, "Arm Chair", None, 115, 116, 117, 118, None],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["1032"]) == [_r(v) for v in (111, 112, 113, 114)]
    assert _dollar_tokens(wide["1034"]) == [_r(v) for v in (115, 116, 117, 118)]


def test_kkt_platinum_not_available_unfinished():
    rows = [
        ["Item #", "PLATINUM", None, "Oak", "Cherry", "QSWO", "Walnut", None],
        ["*Not Available Unfinished", None, None, "Finished", "Finished", "Finished", "Finished", None],
        [1060, "Side Chair", None, 121, 122, 123, 124, None],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["1060"]) == [_r(v) for v in (121, 122, 123, 124)]


def test_kkt_paris_item_code_not_composed_as_wood():
    rows = [
        ["Item #", "NATALIE", None, "Oak", "Cherry", "QSWO", "Finished\nADD"],
        [None, None, None, "Unfinished", "Unfinished", "Unfinished", "Add"],
        [1056, "Side Chair", None, 61, 62, 63, 9],
        ["Item #", "PARIS"],
        [1058, "Chair with Fabric", None, 71, 72, 73, 9],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["1058"]) == [_r(71 + 9), _r(72 + 9), _r(73 + 9)]


def test_ln_bookcase_far_header_five_woods():
    headers = [
        [None, None, None, "UNFINISHED", None, None, None, None, "ESTIMATED FINISHED PRICE"],
        [
            "Item #",
            "Description",
            "Adj. Shelves",
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
        ],
    ]
    # 45 data filler rows mimic a tall L&N section before the SKU
    filler = [["AD3660-O", "Open Bookcase", 3] + [None] * 10] * 45
    rows = headers + filler + [
        ["AD4884-O", "Open Bookcase", 5, 701, 702, 703, 704, 705, 811, 812, 813, 814, 815],
    ]
    out = _process_sheet("Bookcases", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["AD4884-O"]) == [_r(v) for v in (811, 812, 813, 814, 815)]


def test_ln_side_bookcase_subtitle_no_adj_shelf_price():
    headers = [
        [None, None, None, "UNFINISHED", None, None, None, None, "ESTIMATED FINISHED PRICE"],
        [
            "Item #",
            "Description",
            "Adj. Shelves",
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
        ],
    ]
    rows = headers + [
        [None, "Side Bookcases 39½x14x72", 4, None, None, None, None, None, None, None, None, None, None],
    ]
    out = _process_sheet("Bookcases", rows, multiplier=2.7, spots=[])
    assert not any("Side Bookcases 39" in " ".join(r.cells) for r in out if r.kind == RowKind.WIDE)


def test_comfort_72_ez_sleeper_all_grades():
    rows = [
        [None, None, None, None, "Swivel", "Lay Flat", None, None, "Standard\nFabric", "Premium \nFabric", "Ultra \nLeather", "Italian Cowhide Leather", "Heartland\nFabric"],
        [
            '72" E-Z Sleeper Jackknife w/Drawer',
            "N/A",
            "N/A",
            "N/A",
            None,
            None,
            None,
            None,
            461.0,
            462.0,
            463.0,
            464.0,
            465.0,
        ],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = [r.cells for r in out if r.kind == RowKind.WIDE][0]
    assert _dollar_tokens(wide) == [_r(v) for v in (461, 462, 463, 464, 465)]


def test_kkt_pedestal_table_four_finished_woods():
    """Dual-panel Fin|Unf headers must not collapse to a single wood (pedestal tables)."""
    woods = (401.0, 411.0, 421.0, 431.0)
    rows = [
        ["Item Number", "Description", "R. Cherry\nSap Cherry\nOak", None, "R. Hickory\nR. QSWO\nBrown Maple", None, "Elm, Cherry\nHickory\nHard Maple\nR. Walnut", None, "QWSO", None, "Description", "Oak"],
        [None, None, "Finished", "Unfinished", "Finished", "Unfinished", "Finished", "Unfinished", "Finished", "Unfinished", None, "Finished"],
        [None, None, None, None, None, None, None, None, None, None, None, None],
        [None, "•  For Pub Table - Add 10% to Finished Pricee", None, None, None, None, None, None, None, None, None, None],
        ["701-3636", "Solid Top", woods[0], 301, woods[1], 311, woods[2], 321, woods[3], 331, "Solid Top", woods[0]],
    ]
    out = _process_sheet("Single Pedestal Tables", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}["701-3636"]
    assert _dollar_tokens(wide) == [_r(v) for v in woods]


def test_ln_merged_row_finished_band_only_no_adj_shelf_dollars():
    """Bookcase rows: five finished woods only — adj-shelf qty is not a price column."""
    fin = (851.0, 852.0, 853.0, 854.0, 855.0)
    unf = (751.0, 752.0, 753.0, 754.0, 755.0)
    rows = [
        [None, None, None, "UNFINISHED", None, None, None, None, "ESTIMATED FINISHED PRICE"],
        [
            "Item #",
            "Description",
            "Adj. Shelves",
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
        ],
        ["BBC-3836-O", "Open Bookcase", 4, *unf, *fin],
        ["AFD24", "File Desk", "74x24x31", *unf, *fin],
    ]
    out = _process_sheet("Bookcases", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["BBC-3836-O"]) == [_r(v) for v in fin]
    assert _r(4) not in wide["BBC-3836-O"]
    assert _dollar_tokens(wide["AFD24"]) == [_r(v) for v in fin]
    for v in unf:
        assert _r(v) not in wide["AFD24"]


def test_comfort_encore_row_keeps_option_adds():
    tilt, swivel = 51.0, 31.0
    grades = (701.0, 702.0, 703.0, 704.0, 705.0)
    rows = [
        [None, None, None, None, "Swivel", "Lay Flat", None, None, "Standard\nFabric", "Premium \nFabric", "Ultra \nLeather", "Italian Cowhide Leather", "Heartland\nFabric"],
        ['31" Encore Zerowall Recliner', "Standard", tilt, "N/A", swivel, "", "", "Standard", *grades],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = [r.cells for r in out if r.kind == RowKind.WIDE][0]
    assert _dollar_tokens(wide) == [_r(tilt), _r(swivel), *[_r(v) for v in grades]]


def test_kkt_merge_prefers_main_workbook_over_quick_ship_addon():
    import pytest

    from scripts.viztech_book import merge_builder_files

    src = Path("/tmp/p3orig/kountry-knob-tables")
    main = src / "KKT_Pricelist_021026.xlsx"
    qs = src / "KKT Quick Ship Pricelist_0125.xlsx"
    if not main.is_file() or not qs.is_file():
        pytest.skip("viztech originals not on disk")
    res = merge_builder_files([qs, main], builder_name="Kountry Knob Tables", multiplier=2.7)
    assert len(res.rows) > 1000


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
