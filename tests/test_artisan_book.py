from scripts.artisan_book import artisan_forbidden_text, process_artisan_book
from scripts.pricing_lib import retail_whole_dollar


def test_artisan_add_row_multiplies_flat_dollar():
    grid = [
        [None] * 15,
        [None, "2026 Wholesale Price List"],
        [None, "• Finished and Unfinished •"],
        [None, "Prices Effective Starting March 1, 2026"],
        [
            None,
            "Fabric Seat",
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            "ADD",
            None,
            28,
        ],
        [None, "Side Chair", None, 108, 114, 142, 147, 157, 162, 182],
    ]
    processed, title = process_artisan_book(
        grid, builder_name="Artisan Chairs", multiplier=2.7
    )
    assert "Retail Price List" in title
    assert "Wholesale" not in title
    assert artisan_forbidden_text(processed) == []
    wide = [r for r in processed if r.kind.value == "wide"]
    add_row = next(r for r in wide if "Fabric Seat" in str(r.cells))
    assert any("$76" in str(c) for c in add_row.cells)
    assert retail_whole_dollar(108, 2.7) == 292


def test_retail_rounding_28():
    assert retail_whole_dollar(28, 2.7) == 76
