"""Post-build PDF quality checks for retail sheets."""

from __future__ import annotations

import re
from pathlib import Path

import pdfplumber

from scripts.dealer_terms import find_customer_forbidden_in_text, find_forbidden_dealer_phrases

# Literal markup / entity artifacts in extracted text.
MARKUP_ARTIFACT_RE = re.compile(
    r"(?:<br\b|&lt;|&gt;|&amp;|&nbsp;|_x000D_)",
    re.I,
)

# Match SimpleDocTemplate side margins (0.55").
RIGHT_MARGIN_PT = 0.55 * 72
BBOX_TOLERANCE_PT = 3.0


def _page_right_limit(page) -> float:
    return float(page.width) - RIGHT_MARGIN_PT + BBOX_TOLERANCE_PT


def find_pdf_markup_artifacts(pdf_path: Path) -> list[str]:
    issues: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for n, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            for m in MARKUP_ARTIFACT_RE.finditer(text):
                snippet = text[max(0, m.start() - 20) : m.end() + 20]
                issues.append(
                    f"{pdf_path.name} p{n}: markup artifact {m.group()!r} near {snippet!r}"
                )
    return issues


def find_text_past_right_margin(pdf_path: Path) -> list[str]:
    issues: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for n, page in enumerate(pdf.pages, start=1):
            limit = _page_right_limit(page)
            page_h = float(page.height)
            for w in page.extract_words() or []:
                top = float(w.get("top", 0))
                bottom = float(w.get("bottom", 0))
                if top < 58 or page_h - bottom < 42:
                    continue
                x1 = float(w.get("x1", 0))
                if x1 > limit:
                    issues.append(
                        f"{pdf_path.name} p{n}: "
                        f"'{w.get('text', '')}' x1={x1:.1f} > {limit:.1f}"
                    )
    return issues


def find_dealer_terms_in_pdf(pdf_path: Path) -> list[str]:
    issues: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        text = "".join((p.extract_text() or "") + "\n" for p in pdf.pages)
        for hit in find_forbidden_dealer_phrases(text):
            issues.append(f"{pdf_path.name}: forbidden dealer phrase ({hit})")
    return issues


def find_customer_contact_in_pdf(pdf_path: Path) -> list[str]:
    issues: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        text = "".join((p.extract_text() or "") + "\n" for p in pdf.pages)
        seen: set[str] = set()
        for hit in find_customer_forbidden_in_text(text):
            if hit in seen:
                continue
            seen.add(hit)
            issues.append(f"{pdf_path.name}: customer PDF must not contain ({hit})")
    return issues


def assert_pdfs_quality(paths: list[Path]) -> None:
    all_issues: list[str] = []
    for p in paths:
        all_issues.extend(find_pdf_markup_artifacts(p))
        all_issues.extend(find_text_past_right_margin(p))
        all_issues.extend(find_dealer_terms_in_pdf(p))
        all_issues.extend(find_customer_contact_in_pdf(p))
    if all_issues:
        sample = "\n".join(all_issues[:50])
        extra = (
            f"\n... and {len(all_issues) - 50} more"
            if len(all_issues) > 50
            else ""
        )
        raise AssertionError(f"PDF quality check failed:\n{sample}{extra}")
