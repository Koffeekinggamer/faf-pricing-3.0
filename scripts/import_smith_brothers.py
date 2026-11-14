#!/usr/bin/env python3
"""Import Smith Brothers upholstery retail book (FAF customer prices, 1.0× pass-through)."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from scripts.builders_data import slugify, sort_key, upsert_builder
from scripts.multipliers_data import (
    PASSTHROUGH_PATH,
    load_passthrough_builders,
    set_stored_multiplier,
)
from scripts.pdf_builder import build_retail_pdf_processed
from scripts.pdf_overlap_check import assert_pdfs_no_overlap
from scripts.pdf_quality_check import assert_pdfs_quality
from scripts.smith_brothers_book import (
    BUILDER,
    MULTIPLIER,
    process_smith_brothers_pdfs,
    smith_brothers_forbidden,
)

SHEETS_DIR = REPO_ROOT / "sheets"
DEFAULT_PRICE_LIST_DATE = "2026-01-12"


def _save_passthrough_builder(name: str) -> None:
    names = set(load_passthrough_builders())
    if sort_key(name) not in {sort_key(n) for n in names}:
        names.add(name)
        PASSTHROUGH_PATH.parent.mkdir(parents=True, exist_ok=True)
        ordered = sorted(names, key=sort_key)
        with PASSTHROUGH_PATH.open("w", encoding="utf-8") as f:
            json.dump(ordered, f, indent=2, ensure_ascii=False)
            f.write("\n")


def collect_pdf_paths(args: argparse.Namespace) -> list[Path]:
    paths: list[Path] = []
    for raw in args.pdfs:
        p = Path(raw).expanduser().resolve()
        if p.is_dir():
            paths.extend(sorted(p.glob("*.pdf")))
        elif p.is_file():
            paths.append(p)
        else:
            raise FileNotFoundError(f"Not found: {p}")
    if not paths:
        raise ValueError("No PDF inputs")
    return paths


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Import Smith Brothers retail upholstery book (multiplier 1.0)."
    )
    parser.add_argument(
        "pdfs",
        nargs="+",
        help="One or more PDF paths, or a directory of part files (concatenated in name order)",
    )
    parser.add_argument("--date", default=date.today().isoformat(), help="Catalog updated YYYY-MM-DD")
    parser.add_argument(
        "--price-list-date",
        default=DEFAULT_PRICE_LIST_DATE,
        help="Printed book date YYYY-MM-DD",
    )
    parser.add_argument("--force-replace", action="store_true", help="Replace when price list date unchanged")
    args = parser.parse_args()

    try:
        date.fromisoformat(args.date)
        date.fromisoformat(args.price_list_date)
    except ValueError:
        print("Dates must be YYYY-MM-DD", file=sys.stderr)
        return 1

    try:
        pdf_paths = collect_pdf_paths(args)
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    processed, stats = process_smith_brothers_pdfs(pdf_paths, multiplier=MULTIPLIER)
    bad = smith_brothers_forbidden(processed)
    if bad:
        print("Forbidden wording:", "; ".join(bad[:5]), file=sys.stderr)
        return 1

    slug = slugify(BUILDER)
    pdf_path = SHEETS_DIR / f"{slug}.pdf"
    rel_pdf = f"sheets/{slug}.pdf"
    pages = build_retail_pdf_processed(
        processed,
        builder_name=BUILDER,
        updated=args.date,
        output_path=pdf_path,
    )
    assert_pdfs_no_overlap([pdf_path])
    assert_pdfs_quality([pdf_path])

    upsert_builder(
        BUILDER,
        rel_pdf,
        args.date,
        note="FAF retail pass-through (multiplier 1.0)",
        price_list_date=args.price_list_date,
    )
    set_stored_multiplier(BUILDER, MULTIPLIER)
    _save_passthrough_builder(BUILDER)

    report = {"pages": pages, "source_pages": stats["source_pages"], "pdf": str(pdf_path)}
    print(json.dumps(report, indent=2))
    print(f"Imported {BUILDER}: {pages} pages from {stats['source_pages']} book pages", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())