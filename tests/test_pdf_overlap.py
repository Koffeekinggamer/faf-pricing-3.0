"""PDF layout checks: no overlapping text; spot price anchors."""

from __future__ import annotations

import re
from pathlib import Path

import pdfplumber
import pytest

from scripts.pdf_overlap_check import assert_pdfs_no_overlap, find_pdf_overlaps

REPO_ROOT = Path(__file__).resolve().parent.parent
SHEETS = REPO_ROOT / "sheets"

FORBIDDEN = re.compile(
    r"\b(approx|approximately|wholesale|dealer|unfinished|\bunf\b)\b",
    re.I,
)

SPOT_CHECKS = [
    ("Millcraft.pdf", "MBA55FL", "$2,698"),
    ("Frog_Pond_Furniture.pdf", "Weston", "$2,609"),
    ("Frog_Pond_Furniture.pdf", "Weston", "$3,392"),
    ("Townline_Furniture.pdf", "TL-25", "$5,390"),
    ("Woodside_Woodworks.pdf", "14 x 48", "$1,332"),
    ("J_Troyer_and_Co.pdf", "M1284", "$2,792"),
]


@pytest.fixture(scope="module")
def pdf_paths() -> list[Path]:
    paths = sorted(SHEETS.glob("*.pdf"))
    assert len(paths) >= 12, f"expected 12 builder PDFs, found {len(paths)}"
    return paths


def test_no_word_bbox_overlap(pdf_paths: list[Path]) -> None:
    assert_pdfs_no_overlap(pdf_paths)


def test_spot_prices_and_forbidden_words(pdf_paths: list[Path]) -> None:
    by_name = {p.name: p for p in pdf_paths}
    for fname, needle, price in SPOT_CHECKS:
        path = by_name[fname]
        text = ""
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
        assert needle in text, f"{fname} missing {needle}"
        assert price in text, f"{fname} missing {price} near {needle}"
        assert not FORBIDDEN.search(text), f"{fname} contains forbidden retail wording"
