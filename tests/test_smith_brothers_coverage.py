from pathlib import Path

import pdfplumber
import pytest

from scripts.smith_brothers_book import (
    SPOT_CHECKS,
    pdf_has_reversed_artifacts,
    process_smith_brothers_pdfs,
    RowKind,
)
from scripts.smith_brothers_inventory import (
    catalog_has_split_headers,
    catalog_inventory,
    coverage_report,
    load_expected_inventory,
)

CATALOG = Path(__file__).resolve().parents[1] / "sheets" / "Smith_Brothers.pdf"
REVERSED_MARKERS = ("serutaeF", "snoisnemiD", "nommoC", "swolliP", "gnicirP")


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_group_inventory_matches_source():
    exp_groups, exp_skus = load_expected_inventory()
    site_groups, site_skus = catalog_inventory(CATALOG)
    report = coverage_report(exp_groups, exp_skus, site_groups, site_skus)
    assert not report["missing_groups"], report["missing_groups"][:10]
    assert not report["missing_skus"], report["missing_skus"][:10]
    assert report["source_groups"] == 102
    assert report["site_groups"] == 102
    assert report["source_skus"] == report["site_skus"]


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_no_reversed_word_artifacts():
    assert pdf_has_reversed_artifacts(CATALOG) == []
    text = "\n".join((p.extract_text() or "") for p in pdfplumber.open(CATALOG).pages)
    for marker in REVERSED_MARKERS:
        assert marker not in text


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_no_split_header_cells():
    hits = catalog_has_split_headers(CATALOG)
    assert not hits, hits[:5]


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_no_header_only_group_tables():
    """Each group section should include at least one Grade row in the retail PDF."""
    text = "\n".join((p.extract_text() or "") for p in pdfplumber.open(CATALOG).pages)
    import re

    for m in re.finditer(r"\b(\d{3})\s+Group\b", text):
        start = m.start()
        chunk = text[start : start + 4000]
        assert re.search(r"Grade\s+10", chunk), f"{m.group(1)} Group has no Grade 10 row"


@pytest.mark.skipif(not CATALOG.is_file(), reason="Smith Brothers catalog PDF not built")
def test_smith_brothers_extended_spot_checks():
    text = "\n".join((p.extract_text() or "") for p in pdfplumber.open(CATALOG).pages)
    for item, grade, price in SPOT_CHECKS:
        assert price in text, f"{item} {grade} expected {price}"
        if grade.lower().startswith("grade"):
            assert "Grade" in text


def test_smith_brothers_processed_blocks_have_grade_rows():
    """Importer internal sanity: every parsed grade page block includes grade rows."""
    uploads = Path("/home/ubuntu/.cursor/projects/workspace/uploads")
    parts = sorted(uploads.glob("sb_part*.pdf"))[:4]
    if len(parts) < 4:
        pytest.skip("source part PDFs not available")
    rows, _ = process_smith_brothers_pdfs(parts)
    in_section = False
    saw_grade = False
    for pr in rows:
        if pr.kind == RowKind.SECTION and pr.cells[0].endswith("Group"):
            if in_section:
                assert saw_grade, "prior group section had no grade rows"
            in_section = True
            saw_grade = False
        elif in_section and pr.kind == RowKind.WIDE and pr.cells[0].startswith("Grade"):
            saw_grade = True
    if in_section:
        assert saw_grade