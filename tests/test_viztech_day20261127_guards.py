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


def _assert_no_bare_yearish_prices(cells: list[str]) -> None:
    """Price rows must not leak bare 4-digit wholesale amounts (2015–2035)."""
    joined = " ".join(cells)
    for m in re.finditer(r"(?<![\d,$])(20[1-3]\d)(?![\d])", joined):
        assert False, f"bare yearish wholesale leaked: {m.group(0)} in {cells}"


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


def test_kkt_alamo_bare_family_header_composes_add():
    """ALAMO sections omit wood/ADD labels — persist Finished ADD from above."""
    rows = [
        ["Item #", None, "Oak", "Brown Maple", "Cherry", "QSWO", "Finished\nADD"],
        [None, None, "Unfinished", "Unfinished", "Unfinished", "Unfinished", "Add"],
        [1010, "Side Chair", 124, 138, 147, 170, 36],
        ["Item #", "ALAMO"],
        [1014, "Arm Chair", 133, 155, 166, 179, 36],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["1014"]) == ["$457", "$516", "$546", "$581"]
    _assert_no_bare_yearish_prices(wide["1014"])


def test_kkt_lexi_blank_first_wood_shows_remaining():
    rows = [
        ["Item #", "LEXI", None, "Oak", "Maple", "Cherry", "Finished\nADD"],
        [None, None, None, "Unfinished", "Unfinished", "Unfinished", "Add"],
        [2044, "Side", None, 126, 134, 172, 36],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["2044"]) == ["$438", "$459", "$562"]


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
    rows = headers + [
        ["1803-4266", "Solid Top", 1409, 1152, 1303, 1065, "Solid Top", 1409, 1152, 1303, 1065],
    ]
    out = _process_sheet("Greenville", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["1803-4266"]) == ["$3,805", "$3,519"]


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
    rows = [
        ["Description", "Elm\nBrown Maple", None, "Rustic Cherry", None, "Rustic Hickory"],
        [None, "Finished", "Unfinished", "Finished", "Unfinished", "Finished", "Unfinished", None, "Finished", "Unfinished", "Finished", "Unfinished"],
        ["Oregon Arm Chair", 187, 146, 187, 146, None, None, None, 187, 146, 187, 146],
        ["Galveston G-2 Side Chair", 179, 138, None, None, None, None, None, 179, 138],
        ["Galveston G-2 Arm Chair", 204, 163, None, None, None, None, None, 204, 163],
        ["2 Tone", None, 12],
        ["Paint", None, 12],
    ]
    out = _process_sheet("Quick Ship", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["Oregon Arm Chair"]) == ["$505", "$505"]
    assert _dollar_tokens(wide["Galveston G-2 Side Chair"]) == ["$484"]
    assert _dollar_tokens(wide["Galveston G-2 Arm Chair"]) == ["$551"]


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
    rows = headers + [
        ["AD7284", "Corner Desk", "72x36x31", 2081, 2705.3, 4057.95, 2351.53, 3516.89, 2259, 2883.3, 4235.95, 2529.53, 3694.89, 178],
    ]
    cols = _infer_price_cols(rows, 2)
    assert cols == {8, 9, 10, 11, 12}
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = [r.cells for r in out if r.kind == RowKind.WIDE][0]
    assert wide[0] == "AD7284"
    assert "$6,100" in wide and "$5,619" not in wide  # finished Oak, not unfinished
    assert "$481" not in wide  # finish-per-piece 178×2.7


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
    rows = headers + [
        ["AD3660-O", "Open Bookcase", 3, 691, 899, 1348, 781, 1168, 795, 1002.3, 1451.45, 884.83, 1271.79, 104],
        ["FOB-3636-O", "Bookcase", None, 542, 705, 759, 613, 678, 605, 768, 822, 676, 741],
        ["AFD24", "File Desk", "74x24x31", 1500, 2000, 2100, 1800, 1900, 2018, 2574, 2759, 2259, 2481, 166],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["AD3660-O"]) == [
        "$2,147",
        "$2,707",
        "$3,919",
        "$2,390",
        "$3,434",
    ]
    assert "$9" not in wide["AD3660-O"]  # adj shelves
    assert "$281" not in wide["AD3660-O"]  # finish-per-piece
    assert wide["FOB-3636-O"][0] == "FOB-3636-O"
    assert len(_dollar_tokens(wide["FOB-3636-O"])) == 5
    assert _dollar_tokens(wide["FOB-3636-O"])[0] == "$1,634"  # 605×2.7 ceil
    assert _dollar_tokens(wide["AFD24"])[0] == "$5,449"  # 2018×2.7 ceil, not bare year
    assert len(_dollar_tokens(wide["AFD24"])) == 5
    _assert_no_bare_yearish_prices(wide["AFD24"])


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


def test_la_patio_wood_grain_spaced_and_yearish_and_phone_holder():
    headers = [
        ["JADE TABLES", None, None, "STANDARD \nPOLY", "2-TONE\nstd/wood grain", "WOOD GRAIN \nPOLY"],
    ]
    assert _finished_price_cols(headers, 6) == {3, 4, 5}
    rows = headers + [
        ["LAJA4677F-B", '46" x 77" Jade Table w/Fire Ring', "bar", 1870.0, 1970.0, 2020.0],
        ["LACH555", "Screw-On Cup/Phone Holder", "", 25.0, 30.0, 30.0],
        ["Zach Chair is discontinued for 2027"],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["LAJA4677F-B"]) == ["$5,049", "$5,319", "$5,454"]
    _assert_no_bare_yearish_prices(wide["LAJA4677F-B"])
    assert wide["LACH555"][0] == "LACH555"
    assert _dollar_tokens(wide["LACH555"]) == ["$68", "$81", "$81"]
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
            635.0,
            675.0,
            865.0,
            1415.0,
            865.0,
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
            1175.0,
            1215.0,
            1405.0,
            1945.0,
            1405.0,
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
            345.0,
            345.0,
            345.0,
            395.0,
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
            1625.0,
            1685.0,
            2025.0,
            3145.0,
            2025.0,
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
            1635.0,
            1695.0,
            2035.0,
            3565.0,
            2035.0,
        ],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    by_name = {
        r.cells[0]: r.cells for r in out if r.kind == RowKind.WIDE
    }
    dutch = by_name['37" Dutch Boy Wallhugger Recliner']
    assert _dollar_tokens(dutch)[:5] == [
        "$1,715",
        "$1,823",
        "$2,336",
        "$3,821",
        "$2,336",
    ]
    ez = by_name['52" E-Z Sleeper Jackknife w/Drawer']
    assert _dollar_tokens(ez) == [
        "$3,173",
        "$3,281",
        "$3,794",
        "$5,252",
        "$3,794",
    ]
    diner = by_name["Dutch Swivel Diner"]
    assert _dollar_tokens(diner) == ["$932", "$932", "$932", "$1,067", "$932"]
    encore = [c for k, c in by_name.items() if "Encore" in k][0]
    assert "$5,468" in encore
    _assert_no_bare_yearish_prices(encore)
    lazy = [c for k, c in by_name.items() if "Lazy Lounger" in k][0]
    assert "$5,495" in lazy
    _assert_no_bare_yearish_prices(lazy)


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


def test_chimes_phone_sidebar_keeps_820_and_818():
    rows = [
        ["Part #", "Description", "Approx\nSize", "Qty.", "Price", None, "PHONE"],
        [820.0, "Sparkling Brooke - Tex Black", '30"', None, 43.0, None, "PHONE"],
        [818.0, "Zephyr Valley - Tex Black", '37"', None, 49.0, None, "FAX"],
    ]
    assert not _drop_row(rows[1], "820 Sparkling Brooke PHONE")
    assert not _drop_row(rows[2], "818 Zephyr Valley FAX")
    out = _process_sheet("Order Form", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert wide["820"][-1] == "$117"
    assert wide["818"][-1] == "$133"


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
        [1032, "Side Chair", None, 240, 240, 256, 270, None],
        [1034, "Arm Chair", None, 262, 262, 286, 305, None],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["1032"]) == ["$648", "$648", "$692", "$729"]
    assert _dollar_tokens(wide["1034"]) == ["$708", "$708", "$773", "$824"]


def test_kkt_platinum_not_available_unfinished():
    rows = [
        ["Item #", "PLATINUM", None, "Oak", "Cherry", "QSWO", "Walnut", None],
        ["*Not Available Unfinished", None, None, "Finished", "Finished", "Finished", "Finished", None],
        [1060, "Side Chair", None, 240, 240, 256, 270, None],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["1060"]) == ["$648", "$648", "$692", "$729"]


def test_kkt_paris_item_code_not_composed_as_wood():
    rows = [
        ["Item #", "NATALIE", None, "Oak", "Cherry", "QSWO", "Finished\nADD"],
        [None, None, None, "Unfinished", "Unfinished", "Unfinished", "Add"],
        [1056, "Side Chair", None, 114, 130, 157, 36],
        ["Item #", "PARIS"],
        [1058, "Chair with Fabric", None, 203, 213, 238, 36],
    ]
    out = _process_sheet("Chairs", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["1058"]) == ["$646", "$673", "$740"]


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
        ["AD4884-O", "Open Bookcase", 5, 907, 1180, 1769, 1025, 1533, 1054, 1327, 1916, 1172, 1680],
    ]
    out = _process_sheet("Bookcases", rows, multiplier=2.7, spots=[])
    wide = {c[0]: c for c in (r.cells for r in out if r.kind == RowKind.WIDE)}
    assert _dollar_tokens(wide["AD4884-O"]) == [
        "$2,846",
        "$3,583",
        "$5,174",
        "$3,165",
        "$4,536",
    ]


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
            1265.0,
            1305.0,
            1525.0,
            2445.0,
            1525.0,
        ],
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    wide = [r.cells for r in out if r.kind == RowKind.WIDE][0]
    assert _dollar_tokens(wide) == [
        "$3,416",
        "$3,524",
        "$4,118",
        "$6,602",
        "$4,118",
    ]


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
