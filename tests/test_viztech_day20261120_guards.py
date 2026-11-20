"""Guards for Viztech day20261120 extraction edge cases (no upload dependency)."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from scripts.dealer_terms import is_builder_contact_text
from scripts.price_list_date import parse_price_list_phrase
from scripts.viztech_book import (
    _drop_row,
    _finished_price_cols,
    _infer_price_cols,
    section_title_from_path,
)


def test_price_row_not_mistaken_for_phone_contact():
    row = [None, '42x60 Solid Top', None, 673.1285313623807, 913.25, 946.71]
    joined = "42x60 Solid Top 673.1285313623807 913.25 946.71"
    assert not is_builder_contact_text("42x60 Solid Top")
    assert not _drop_row(row, joined)


def test_parse_numeric_effective_date():
    assert parse_price_list_phrase("Effective Date 3-15-2026") == date(2026, 3, 15)
    assert parse_price_list_phrase("Effective May 1, 2026") == date(2026, 5, 1)


def test_finished_cols_include_stain_not_unfinished():
    headers = [["Item#", "Description", "Unfinished", "Stain", "2-Tone & Glazing"]]
    cols = _finished_price_cols(headers, 5)
    assert cols == {3, 4}


def test_infer_skips_item_number_as_price():
    rows = [
        ["Item #", "Item Name", "Finished"],
        [5136, "Brown Swirl Sofa Table", 580],
        [5138, "Black Round Cedar Slab", 997],
    ]
    cols = _infer_price_cols(rows, 1)
    assert 0 not in cols
    assert 2 in cols


def test_section_titles_from_upload_names():
    assert section_title_from_path(Path("fishers-furniture__FF_Slabs_1026.xlsx")) == "Slabs"
    assert section_title_from_path(Path("fighting-creek-furniture__FCF_2026_Barrel_Pricelist_0326.xlsx")) == "Barrel"
    assert section_title_from_path(Path("weaver-woodcraft__FST_2026_Dining_Pricelist_0326-1.xlsx")) == "Dining"
    assert section_title_from_path(Path("weaver-woodcraft__FST_2026_Saw_Marked_Pricelist_0326.xlsx")) == "Saw Marked"
