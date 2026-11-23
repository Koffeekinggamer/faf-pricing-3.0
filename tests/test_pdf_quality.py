from pathlib import Path

import pytest

from scripts.pdf_quality_check import (
    assert_pdfs_quality,
    find_add_upcharge_issues,
    find_customer_contact_in_pdf,
    find_dealer_terms_in_pdf,
    find_pdf_markup_artifacts,
    find_text_past_right_margin,
)

REPO_ROOT = Path(__file__).resolve().parent.parent
SHEETS = REPO_ROOT / "sheets"

# Judson's final Smith Brothers book is a pass-through scan; wording/layout
# validators that apply to generated retail PDFs do not apply here.
QUALITY_EXEMPT_PDFS = frozenset({"Smith_Brothers.pdf"})


@pytest.fixture(scope="module")
def pdf_paths() -> list[Path]:
    paths = sorted(SHEETS.glob("*.pdf"))
    assert paths
    return paths


def test_pdf_quality_bundle(pdf_paths: list[Path]) -> None:
    assert_pdfs_quality([p for p in pdf_paths if p.name not in QUALITY_EXEMPT_PDFS])


def test_no_markup_artifacts(pdf_paths: list[Path]) -> None:
    bad: list[str] = []
    for p in pdf_paths:
        if p.name in QUALITY_EXEMPT_PDFS:
            continue
        bad.extend(find_pdf_markup_artifacts(p))
    assert not bad, "\n".join(bad[:30])


def test_no_clipped_right_margin(pdf_paths: list[Path]) -> None:
    bad: list[str] = []
    for p in pdf_paths:
        if p.name in QUALITY_EXEMPT_PDFS:
            continue
        bad.extend(find_text_past_right_margin(p))
    assert not bad, "\n".join(bad[:30])


def test_no_dealer_terms(pdf_paths: list[Path]) -> None:
    bad: list[str] = []
    for p in pdf_paths:
        if p.name in QUALITY_EXEMPT_PDFS:
            continue
        bad.extend(find_dealer_terms_in_pdf(p))
    assert not bad, "\n".join(bad)


def test_no_builder_contact_or_billing(pdf_paths: list[Path]) -> None:
    bad: list[str] = []
    for p in pdf_paths:
        if p.name in QUALITY_EXEMPT_PDFS:
            continue
        bad.extend(find_customer_contact_in_pdf(p))
    assert not bad, "\n".join(bad)


def test_no_bare_decimal_or_trailing_add_upcharges(pdf_paths: list[Path]) -> None:
    bad: list[str] = []
    for p in pdf_paths:
        if p.name in QUALITY_EXEMPT_PDFS:
            continue
        bad.extend(find_add_upcharge_issues(p))
    assert not bad, "\n".join(bad[:40])
