"""Guards for Viztech day20261123 extraction edge cases (no upload dependency)."""

from __future__ import annotations

from scripts.pdf_quality_check import find_add_upcharge_issues
from scripts.retail_text import (
    excel_percent_to_display,
    rewrite_add_fraction_as_percent,
    sanitize_retail_text,
)
from scripts.viztech_book import (
    SpotSample,
    _finished_price_cols,
    _infer_price_cols,
    _process_sheet,
    _row_item_label,
    _transform_add_upcharge_row,
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


def test_excel_percent_display_no_multiplier():
    assert excel_percent_to_display(0.04) == "4%"
    assert excel_percent_to_display(0.06) == "6%"
    assert excel_percent_to_display(0.1) == "10%"
    assert excel_percent_to_display(0.12) == "12%"


def test_add_upcharge_percent_tokens_not_multiplied():
    spots: list[SpotSample] = []
    cells = _transform_add_upcharge_row(
        [None, "Standard Color and Tropical Color, Add", None, None, None, None, "4%"],
        2.7,
        spots=spots,
        item_hint="",
    )
    assert cells == ["Standard Color and Tropical Color, Add 4%"]
    assert spots == []


def test_add_upcharge_keeps_dollar_values_with_multiplier():
    spots: list[SpotSample] = []
    cells = _transform_add_upcharge_row(
        [None, None, "Seat, ADD", None, None, None, None, None, None, 53],
        2.7,
        spots=spots,
        item_hint="",
    )
    assert len(cells) == 1
    assert cells[0].startswith("Seat, ADD $")
    assert "53" not in cells[0]  # retail, not wholesale


def test_process_sheet_keeps_all_percent_add_rows():
    rows = [
        [None, "Standard Color and Tropical Color, Add", None, None, None, None, "4%"],
        [None, "Standard Color and Premium Color, Add", None, None, None, None, "6%"],
        [None, "All Tropical Color, Add", None, None, None, None, "10%"],
        [None, "All Premium Color, Add", None, None, None, None, "12%"],
        [None, "Options", "Add"],  # header — dropped
    ]
    out = _process_sheet("Wholesale", rows, multiplier=2.7, spots=[])
    joined = " | ".join(" ".join(r.cells) for r in out if r.cells)
    assert "Add 4%" in joined
    assert "Add 6%" in joined
    assert "Add 10%" in joined
    assert "Add 12%" in joined
    assert not joined.rstrip().endswith("Add")
    assert "Options" not in joined


def test_rewrite_add_fraction_of_regular_price():
    assert (
        rewrite_add_fraction_as_percent("Ajay built for full size mattress add 0.1 of regular price")
        == "Ajay built for full size mattress add 10% of regular price"
    )


def test_day20261123_pdfs_have_no_add_upcharge_issues():
    from pathlib import Path

    sheets = Path(__file__).resolve().parent.parent / "sheets"
    targets = [
        "Hillside_Chair.pdf",
        "Hilltop_Hickory_Furniture.pdf",
        "Hochstetler_Lawn_Furniture.pdf",
        "Hogback_Design_And_Finishing.pdf",
        "Hoosier_Crafts.pdf",
    ]
    bad: list[str] = []
    for name in targets:
        p = sheets / name
        if not p.is_file():
            continue
        bad.extend(find_add_upcharge_issues(p))
    assert not bad, "\n".join(bad[:40])


def test_hillside_embedded_dollar_add_notes_not_dropped():
    """Prose 'Seat - Add $10' cells must keep multiplied $ lines (not treated as empty Add)."""
    rows = [
        [
            "Profile Comfort Seat - Add $10\nFabrice Seat - Add $15\nLeather Seat - Add $45\n"
            "Cherry, Hickory, Rustic Hickory, Rustic QSWO, Elm Seat - Add $10\n"
            "QSWO, Rustic Walnut Seat - Add $20",
            "Walnut Seat -  Add $35\nDouble X Fabric Seat & Back - Add $50\n"
            "Double X Leather Seat & Back - Add $110",
        ],
        ["Hardwood XX's in Back (for Double X Chair) Add", 15, 18, 20, 23, 26],
        ["Options", "Add"],
    ]
    out = _process_sheet("Sheet3", rows, multiplier=2.7, spots=[])
    text = "\n".join(" ".join(c for c in r.cells if c) for r in out)
    for needle in (
        "Profile Comfort Seat - Add $27",
        "Fabrice Seat - Add $41",
        "Leather Seat - Add $122",
        "Cherry, Hickory, Rustic Hickory, Rustic QSWO, Elm Seat - Add $27",
        "QSWO, Rustic Walnut Seat - Add $54",
        "Walnut Seat - Add $95",
        "Double X Fabric Seat & Back - Add $135",
        "Double X Leather Seat & Back - Add $297",
        "Hardwood XX's in Back (for Double X Chair) Add $41",
    ):
        assert needle in text, needle
    assert "Options" not in text
    assert text.count("Hardwood XX's in Back") == 1


def test_day20261123_dollar_line_counts_vs_source(tmp_path=None):
    """Each builder's processed $ line count must not drop vs source value-bearing rows."""
    from pathlib import Path

    from scripts.viztech_book import load_workbook_sheets, process_viztech_workbook

    sources = Path("/tmp/day20261123/day20261123")
    if not sources.is_dir():
        return  # upload not present in CI; synthetic coverage above still runs

    builders = [
        ("Hillside Chair", "hillside-chair__HC_2026_Pricelist_0426-1.xlsx"),
        (
            "Hilltop Hickory Furniture",
            "hilltop-hickory-furniture__HHF_2026_Pricelist_0326-CompressedImages.xlsx",
        ),
        (
            "Hochstetler Lawn Furniture",
            "hochstetler-lawn-furniture__HLF_2027_Pricelist_1026.xlsx",
        ),
        (
            "Hogback Design And Finishing",
            "hogback-design-and-finishing__HDF_2026_Pricelist_0326-1.xlsx",
        ),
        ("Hoosier Crafts", "hoosier-crafts__HC_2026_Pricelist_0326.xlsx"),
    ]

    def source_price_rows(path: Path) -> int:
        sheets = load_workbook_sheets(path)
        n = 0
        for rows in sheets.values():
            for row in rows:
                if any(
                    isinstance(c, (int, float))
                    and not isinstance(c, bool)
                    and 1 <= float(c) <= 250_000
                    for c in row
                ):
                    n += 1
                elif any(isinstance(c, str) and "$" in c for c in row):
                    n += 1
        return n

    for name, fname in builders:
        path = sources / fname
        if not path.is_file():
            continue
        src = source_price_rows(path)
        res = process_viztech_workbook(path, multiplier=2.7, builder_name=name)
        out_dollar = sum(1 for r in res.rows if any("$" in str(c) for c in r.cells))
        # Allow some filtering (unfinished/retail sheets) but never a collapse.
        assert out_dollar >= max(10, int(src * 0.35)), (
            f"{name}: $ lines {out_dollar} vs source price rows {src}"
        )
        if name == "Hillside Chair":
            text = "\n".join(" ".join(map(str, r.cells)) for r in res.rows)
            assert "Profile Comfort Seat - Add $27" in text
            assert "Leather Seat - Add $122" in text
            assert "Double X Leather Seat & Back - Add $297" in text
