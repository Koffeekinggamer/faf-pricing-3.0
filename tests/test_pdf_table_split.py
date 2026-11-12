from scripts.pdf_table_layout import (
    assert_split_headers_unique,
    chunks_for_split,
    non_key_headers_in_chunks,
)


def test_chunks_are_disjoint_not_cumulative():
    widths = [0.5] * 2 + [0.4] * 12
    chunks = chunks_for_split(widths, prefix=2, usable=2.5)
    assert chunks[0][0] >= 2
    for a, b in zip(chunks, chunks[1:]):
        assert a[1] == b[0], "chunks must meet end-to-end without overlap"
    assert chunks[-1][1] == len(widths)


def test_split_headers_must_not_repeat():
    rows = [
        ["Item", "Cat", "Oak", "Maple", "Cherry", "Walnut", "Hickory"],
        ["Chair", "Cat. 1", "$1", "$2", "$3", "$4", "$5"],
    ]
    chunks = [(2, 4), (4, 6), (6, 7)]
    per = non_key_headers_in_chunks(rows, prefix=2, chunks=chunks, header_rows=1)
    assert per[0] == ["Oak", "Maple"]
    assert per[1] == ["Cherry", "Walnut"]
    assert per[2] == ["Hickory"]
    assert_split_headers_unique(rows, prefix=2, chunks=chunks, header_rows=1)

    bad = [(2, 5), (4, 7)]
    try:
        assert_split_headers_unique(rows, prefix=2, chunks=bad, header_rows=1)
    except ValueError as exc:
        assert "repeats column header" in str(exc)
    else:
        raise AssertionError("expected duplicate header failure")
