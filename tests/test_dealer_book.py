from scripts.dealer_book import forbidden_retail_text, process_dealer_book


def test_dealer_book_strips_wholesale_and_fob():
    grid = [
        [None] * 5,
        [
            None,
            "January 11th, 2027 PREFERRED DEALER Wholesale Price List",
            None,
            None,
            None,
        ],
        [None, "All prices are F.O.B. Berlin.", None, None, None],
        [None, "Price List Index", None, None, None],
        [None, "Albany", 6, None, "Victoria's Trad-", 62],
        [None, "Standard Options *No Upcharge*", None, None, None],
        [None, "Description", "Item Number", "Standard Wood", "Premium Wood"],
        [None, "Full  ", "MBA55FL", 999, 1198],
        [None, "Hidden Jewelry Tray: $30", None, None, "Lock: $25"],
    ]
    processed, title = process_dealer_book(
        grid, builder_name="Millcraft", multiplier=2.7
    )
    assert "Retail Price List" in title
    assert "Wholesale" not in title
    assert forbidden_retail_text(processed) == []
    sku = [r for r in processed if r.kind.value == "sku"][0]
    assert sku.cells[2] == "$2,698"
    assert sku.cells[3] == "$3,235"
