"""Guards for Viztech day20261126 extraction edge cases (no upload dependency)."""

from __future__ import annotations

from pathlib import Path

from scripts.dealer_book import RowKind
from scripts.viztech_book import (
    ADD_ON_PRICE_COL_RE,
    _addon_cols_above,
    _eval_round_product_formulas,
    _excluded_price_header_cols,
    _finished_price_cols,
    _infer_price_cols,
    _process_sheet,
    _wood_cols_before_finished_add,
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
    # Known catalog patterns still win over slug-suffix parsing
    assert section_title_from_path(Path("fishers-furniture__FF_Slabs_1026.xlsx")) == (
        "Slabs"
    )
    assert section_title_from_path(
        Path("hershy-way-ltd__2026_Cyp and Treated Pricelist_0126.xlsx")
    ) == "Cypress & Treated"


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
    # Bare Options/"Add" must not be treated as a finish-Add column header.
    assert not ADD_ON_PRICE_COL_RE.search("Add")
    assert not ADD_ON_PRICE_COL_RE.fullmatch("Add")


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


def test_jericho_six_wood_columns_desc_size_and_m12_prices():
    """Master: 6 woods from Oak×factors; Fenton M-12: 5 woods; no priceless SKUs."""
    headers = [
        [None, None, "Size", "Oak", "Br. Maple", "RQSWO", "Ash", "Alder", "QSWO", "Finished\nPrices\nAdd"],
        ["Item No.", "Description", None, None, "Sap Cherry", None, None, None, None, None],
        [None, None, None, None, 1.15, 1.2, 1.25, 1.35, 1.45, None],
    ]
    master_rows = headers + [
        ["M38", "Small Wedge Table", '10" front x 18" back', 161, None, None, None, None, None, 33],
        ["    Tables"],
    ]
    formulas = [[None] * 10 for _ in headers] + [
        [
            None,
            None,
            None,
            None,
            "=ROUND(D4*$E$3,0)",
            "=ROUND(D4*$F$3,0)",
            "=ROUND(D4*$G$3,0)",
            "=ROUND(D4*$H$3,0)",
            "=ROUND(D4*$I$3,0)",
            None,
        ],
        [None] * 10,
    ]
    _eval_round_product_formulas(
        master_rows, formulas, sheet_name="Master", cross={"Master": master_rows}
    )
    assert master_rows[3][3:9] == [161, 185.0, 193.0, 201.0, 217.0, 233.0]
    assert _wood_cols_before_finished_add(master_rows, 12) == {3, 4, 5, 6, 7, 8}

    out = _process_sheet("Master", master_rows, multiplier=2.7, spots=[])
    wide = [r for r in out if r.kind == RowKind.WIDE]
    assert wide, out
    cells = wide[0].cells
    assert cells[0] == "M38"
    assert "Small Wedge Table" in cells[1]
    assert any("10" in c for c in cells[1:3])
    prices = [c for c in cells if c.startswith("$")]
    assert prices == ["$435", "$500", "$522", "$543", "$586", "$630"]
    assert 33 * 2.7 != float(prices[-1].replace("$", "").replace(",", ""))  # Add excluded

    # Fenton dual panel: right-panel woods only; Custom title without prices dropped from WIDE
    fenton = [
        [
            None,
            None,
            "Size",
            "Oak\nBr. Maple",
            "Ash",
            "Alder",
            "Elm",
            "QSWO",
            "Finished\nPrices\nAdd",
            None,
            None,
            "Size",
            "Oak\nBr. Maple",
            "Ash",
            "Alder",
            "Elm",
            "QSWO",
            "Finished\nPrices\nAdd",
        ],
        ["Item No.", "Description", None, None, None, None, None, None, None, "Item No.", "Description"],
        [None] * 12 + [None, 1.15, 1.2, 1.35, 1.45, None],
        ["Custom M-77 with Solid Side Panel and Back", None, None, None, None, None, None, None, None,
         "Custom M-77 with Solid Side Panel and Back"],
        [
            "M-12",
            "Small Lamp Table",
            '15" x 16" x 26"',
            None,
            None,
            None,
            None,
            None,
            27,
            "M-12",
            "Small Lamp Table",
            '15" x 16" x 26"',
            145,
            167.0,
            174.0,
            196.0,
            210.0,
            25,
        ],
    ]
    assert _wood_cols_before_finished_add(fenton, 20) == {12, 13, 14, 15, 16}
    fout = _process_sheet("Fenton Only", fenton, multiplier=2.7, spots=[])
    fwide = [r for r in fout if r.kind == RowKind.WIDE]
    assert len(fwide) == 1
    assert fwide[0].cells[0] == "M-12"
    assert "Small Lamp Table" in fwide[0].cells[1]
    fprices = [c for c in fwide[0].cells if c.startswith("$")]
    assert len(fprices) == 5
    assert "$392" in fprices  # 145 * 2.7
    # no priceless WIDE rows
    assert all(any(c.startswith("$") for c in r.cells) for r in fwide)
