from datetime import date

from scripts.price_list_date import (
    evaluate_price_list_update,
    mmyy_from_path,
    parse_iso_price_list_date,
    parse_price_list_phrase,
)
from pathlib import Path


def test_parse_effective_phrase():
    assert parse_price_list_phrase("Effective March 1, 2026") == date(2026, 3, 1)
    assert parse_price_list_phrase("April 15, 2025") == date(2025, 4, 15)


def test_mmyy_from_filename():
    assert mmyy_from_path(Path("BF_2026_Pricelist_0526.xlsx")) == date(2026, 5, 1)


def test_evaluate_newer_and_skip():
    existing = {"name": "X", "price_list_date": "2026-05-01"}
    ok, _ = evaluate_price_list_update(existing, date(2026, 6, 1))
    assert ok
    ok, reason = evaluate_price_list_update(existing, date(2026, 5, 1))
    assert not ok
    assert "same" in reason
    ok, reason = evaluate_price_list_update(existing, date(2026, 4, 1))
    assert not ok
    assert "older" in reason


def test_evaluate_undetermined_incoming():
    existing = {"name": "X", "price_list_date": "2026-05-01"}
    ok, reason = evaluate_price_list_update(existing, None)
    assert not ok
    assert "could not be determined" in reason


def test_evaluate_new_builder():
    ok, _ = evaluate_price_list_update(None, date(2026, 1, 1))
    assert ok
    ok, _ = evaluate_price_list_update(None, None)
    assert ok


def test_iso_parse():
    assert parse_iso_price_list_date("2026-11-12") == date(2026, 11, 12)
    assert parse_iso_price_list_date("bad") is None
