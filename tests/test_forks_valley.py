"""Unit tests for Forks Valley retail transforms."""

from math import ceil

from scripts.forks_valley_book import COLOR_BAKED_MARKUP, _transform_price
from scripts.pricing_lib import retail_whole_dollar


def test_color_upcharge_uses_baked_wholesale_base():
    mult = 2.7
    # Open end table cat 2: cached 66.96 = 62 * 1.08
    base = 66.96 / COLOR_BAKED_MARKUP
    assert abs(base - 62) < 0.01
    assert _transform_price(66.96, mult, color_sheet=True) == f"${retail_whole_dollar(base, mult):,}"
    assert retail_whole_dollar(base, mult) == ceil(62 * mult)


def test_main_list_price_not_double_marked():
    mult = 2.7
    wholesale = 326.2262
    assert _transform_price(wholesale, mult, color_sheet=False) == f"${retail_whole_dollar(wholesale, mult):,}"


def test_glass_std_preserved():
    assert _transform_price("STD", 2.7, color_sheet=False) == "STD"


def test_glass_add_retail():
    mult = 2.7
    assert _transform_price(35.0, mult, color_sheet=False) == f"${retail_whole_dollar(35.0, mult):,}"
