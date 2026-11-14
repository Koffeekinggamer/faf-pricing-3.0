from pathlib import Path

import pytest

from scripts.indian_trail_book import process_indian_trail_workbook
from scripts.jr_woodworking_book import process_jr_woodworking_workbook
from scripts.pricing_lib import retail_whole_dollar

UPLOAD = Path("/home/ubuntu/.cursor/projects/workspace/uploads")
IT_XLSX = UPLOAD / "Indian_Trail_2026_648127_aff1.xlsx"
JR_XLSX = UPLOAD / "J_and_R_2026_306462_e390.xlsx"

IT_PDF = Path(__file__).resolve().parents[1] / "sheets" / "Indian_Trail_Furniture.pdf"
JR_PDF = Path(__file__).resolve().parents[1] / "sheets" / "J_R_Woodworking.pdf"


def test_retail_from_wholesale_spots():
    assert retail_whole_dollar(1163, 2.7) == 3141
    assert retail_whole_dollar(1267, 2.7) == 3421
    assert retail_whole_dollar(1012, 2.7) == 2733


@pytest.mark.skipif(not IT_XLSX.is_file(), reason="Indian Trail upload missing")
def test_indian_trail_wholesale_finished_only():
    res = process_indian_trail_workbook(IT_XLSX, multiplier=2.7)
    assert not res.warnings or res.rows
    texts = " ".join(str(c) for pr in res.rows for c in pr.cells)
    assert "Troyer" not in texts
    assert "2% after 30" not in texts
    assert "$3,141" in texts or "$3,141" in str(
        retail_whole_dollar(1163, 2.7)
    )


@pytest.mark.skipif(not JR_XLSX.is_file(), reason="J&R upload missing")
def test_jr_wholesale_finished_only():
    res = process_jr_woodworking_workbook(JR_XLSX, multiplier=2.7)
    texts = " ".join(str(c) for pr in res.rows for c in pr.cells)
    assert "Net 30" not in texts
    assert "Miller & Family" not in texts
    assert "$3,421" in texts
    assert "$2,733" in texts


@pytest.mark.skipif(not IT_PDF.is_file(), reason="Indian Trail PDF not built")
def test_indian_trail_pdf_spot():
    import pdfplumber

    with pdfplumber.open(IT_PDF) as doc:
        full = "\n".join(p.extract_text() or "" for p in doc.pages)
    assert "ITAB-032-T-CB" in full and "$3,141" in full


@pytest.mark.skipif(not JR_PDF.is_file(), reason="J&R PDF not built")
def test_jr_pdf_spot():
    import pdfplumber

    with pdfplumber.open(JR_PDF) as doc:
        full = "\n".join(p.extract_text() or "" for p in doc.pages)
    assert "JRAB-073" in full and "$3,421" in full
    assert "JRAB-059" in full and "$2,733" in full