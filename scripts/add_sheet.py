#!/usr/bin/env python3
"""Add or refresh a builder retail price sheet on Pricing 3.0."""

from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from scripts.builders_data import REPO_ROOT as ROOT, slugify, upsert_builder
from scripts.artisan_book import (
    artisan_forbidden_text,
    is_artisan_style_grid,
    process_artisan_book,
)
from scripts.dealer_book import (
    extract_catalog_date,
    forbidden_retail_text,
    process_dealer_book,
)
from scripts.pdf_builder import build_retail_pdf, build_retail_pdf_processed
from scripts.pricing_lib import assert_no_approx
from scripts.read_dealer_file import (
    DealerFileError,
    apply_multiplier_to_grid,
    is_dealer_book_grid,
    read_dealer_file,
)

SHEETS_DIR = ROOT / "sheets"


def main() -> int:
    parser = argparse.ArgumentParser(description="Add a builder retail price sheet")
    parser.add_argument("--builder", required=True, help="Builder display name")
    parser.add_argument(
        "--multiplier",
        type=float,
        required=True,
        help="Retail multiplier (wholesale × multiplier, round up)",
    )
    parser.add_argument("--file", required=True, type=Path, help="Dealer source file")
    parser.add_argument("--note", default=None, help="Optional short note on the index")
    parser.add_argument(
        "--date",
        default=None,
        help="Updated date YYYY-MM-DD (default: today)",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Worksheet name for .xlsx (default: Pricelist if present)",
    )
    args = parser.parse_args()

    if args.multiplier <= 0:
        print("Multiplier must be positive.", file=sys.stderr)
        return 1

    source = args.file.expanduser().resolve()
    if not source.is_file():
        print(f"File not found: {source}", file=sys.stderr)
        return 1

    updated = args.date or date.today().isoformat()
    try:
        date.fromisoformat(updated)
    except ValueError:
        print("--date must be YYYY-MM-DD", file=sys.stderr)
        return 1

    assert_no_approx(args.builder, args.note or "")

    try:
        grid = read_dealer_file(source, sheet_name=args.sheet)
    except DealerFileError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    slug = slugify(args.builder)
    pdf_name = f"{slug}.pdf"
    pdf_path = SHEETS_DIR / pdf_name
    rel_pdf = f"sheets/{pdf_name}"

    if is_artisan_style_grid(grid):
        processed, _title = process_artisan_book(
            grid,
            builder_name=args.builder,
            multiplier=args.multiplier,
        )
        bad = artisan_forbidden_text(processed)
        if bad:
            print(
                "Retail PDF would contain dealer-only wording: "
                + "; ".join(bad[:3]),
                file=sys.stderr,
            )
            return 1
        pages = build_retail_pdf_processed(
            processed,
            builder_name=args.builder,
            updated=updated,
            output_path=pdf_path,
        )
        print(f"Wrote {pages}-page retail PDF", file=sys.stderr)
    elif is_dealer_book_grid(grid):
        catalog_date = extract_catalog_date(grid)
        processed, _title = process_dealer_book(
            grid,
            builder_name=args.builder,
            multiplier=args.multiplier,
            catalog_date=catalog_date,
        )
        bad = forbidden_retail_text(processed)
        if bad:
            print(
                "Retail PDF would contain dealer-only wording: "
                + "; ".join(bad[:3]),
                file=sys.stderr,
            )
            return 1
        pages = build_retail_pdf_processed(
            processed,
            builder_name=args.builder,
            updated=updated,
            output_path=pdf_path,
        )
        print(f"Wrote {pages}-page retail PDF", file=sys.stderr)
    else:
        grid = apply_multiplier_to_grid(grid, args.multiplier)
        build_retail_pdf(
            grid,
            builder_name=args.builder,
            updated=updated,
            output_path=pdf_path,
        )

    entry, replaced = upsert_builder(
        args.builder,
        rel_pdf,
        updated,
        args.note,
    )

    action = "Replaced" if replaced else "Added"
    print(f"{action} {entry['name']} → {rel_pdf} (updated {entry['updated']})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
