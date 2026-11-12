from scripts.pdf_plain_text import clean_pdf_plain_text, escape_for_paragraph_markup


def test_strips_x000d_and_literal_br():
    raw = "Line one_x000D_Line two<br/>Line three"
    assert clean_pdf_plain_text(raw) == "Line one\nLine two\nLine three"
    esc = escape_for_paragraph_markup(raw)
    assert "<br/>" in esc
    assert "&lt;br" not in esc
    assert "_x000D_" not in esc


def test_escape_shows_no_literal_angle_brackets():
    esc = escape_for_paragraph_markup("A < B and C > D")
    assert esc == "A &lt; B and C &gt; D"
