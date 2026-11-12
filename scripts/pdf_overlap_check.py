"""Detect overlapping word bounding boxes in retail PDFs."""

from __future__ import annotations

from pathlib import Path

import pdfplumber

# Ignore sub-point grazing at shared edges between adjacent words.
MIN_OVERLAP_PT = 0.75


def _overlap_area(
    a: tuple[float, float, float, float],
    b: tuple[float, float, float, float],
) -> float:
    x0 = max(a[0], b[0])
    x1 = min(a[2], b[2])
    y0 = max(a[1], b[1])
    y1 = min(a[3], b[3])
    if x1 - x0 <= MIN_OVERLAP_PT or y1 - y0 <= MIN_OVERLAP_PT:
        return 0.0
    return (x1 - x0) * (y1 - y0)


def find_pdf_overlaps(pdf_path: Path) -> list[str]:
    issues: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            words = page.extract_words() or []
            boxes = [
                (
                    w["text"],
                    (float(w["x0"]), float(w["top"]), float(w["x1"]), float(w["bottom"])),
                )
                for w in words
                if w.get("text")
            ]
            for i in range(len(boxes)):
                t1, b1 = boxes[i]
                for j in range(i + 1, len(boxes)):
                    t2, b2 = boxes[j]
                    if _overlap_area(b1, b2) > 0:
                        issues.append(
                            f"{pdf_path.name} page {page_no}: "
                            f"'{t1}' overlaps '{t2}'"
                        )
    return issues


def assert_pdfs_no_overlap(paths: list[Path]) -> None:
    all_issues: list[str] = []
    for p in paths:
        all_issues.extend(find_pdf_overlaps(p))
    if all_issues:
        sample = "\n".join(all_issues[:40])
        extra = f"\n... and {len(all_issues) - 40} more" if len(all_issues) > 40 else ""
        raise AssertionError(f"PDF text overlap detected:\n{sample}{extra}")
