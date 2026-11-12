import math

import pytest

from scripts.builders_data import load_builders, save_builders, sort_key, upsert_builder
from scripts.pricing_lib import (
    assert_no_approx,
    retail_whole_dollar,
    transform_text_prices,
)


def test_retail_rounds_up_to_whole_dollar():
    assert retail_whole_dollar(683.20, 1.0) == 684
    assert retail_whole_dollar(253.0, 2.7) == 684
    assert retail_whole_dollar(100.0, 2.7) == 270
    assert retail_whole_dollar(99.01, 2.0) == 199


def test_no_approx_in_transform_or_assert():
    out = transform_text_prices("Chair $683.20", 1.0)
    assert "approx" not in out.lower()
    assert "$684" in out
    with pytest.raises(ValueError):
        assert_no_approx("pricing is approx")


def test_alphabetical_insert_in_json(tmp_path):
    path = tmp_path / "builders.json"
    save_builders([], path)
    upsert_builder("Zebra Woodworks", "sheets/z.pdf", "2025-01-01", path=path)
    upsert_builder("Acme Shop", "sheets/a.pdf", "2025-02-01", path=path)
    upsert_builder("Middling Co", "sheets/m.pdf", "2025-03-01", path=path)
    names = [b["name"] for b in load_builders(path)]
    assert names == ["Acme Shop", "Middling Co", "Zebra Woodworks"]


def test_sort_key_ignores_leading_punctuation():
    assert sort_key("& Sons Furniture") == sort_key("Sons Furniture")
    assert sort_key("acme") == sort_key("Acme")
    assert sort_key("Acme Shop") < sort_key("Beta LLC")


def test_replace_on_readd(tmp_path):
    path = tmp_path / "builders.json"
    save_builders(
        [{"name": "Look-alike Builder", "pdf": "sheets/old.pdf", "updated": "2024-01-01"}],
        path,
    )
    entry, replaced = upsert_builder(
        "Look-alike Builder",
        "sheets/new.pdf",
        "2026-11-12",
        note="Refresh",
        path=path,
    )
    assert replaced is True
    builders = load_builders(path)
    assert len(builders) == 1
    assert builders[0]["pdf"] == "sheets/new.pdf"
    assert builders[0]["updated"] == "2026-11-12"
    assert builders[0]["note"] == "Refresh"


def test_percentage_unchanged():
    text = "Leather +15%"
    assert transform_text_prices(text, 2.7) == text
