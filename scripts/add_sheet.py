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
from scripts.farmside_book import (
    farmside_forbidden_text,
    is_farmside_workbook,
    process_farmside_workbook,
)
from scripts.fn_chairs_book import (
    fn_chairs_forbidden_text,
    is_fn_chairs_workbook,
    process_fn_chairs_workbook,
)
from scripts.forks_valley_book import (
    forks_valley_forbidden_text,
    is_forks_valley_bundle,
    process_forks_valley_bundle,
)
from scripts.lamb_book import (
    is_lamb_workbook,
    lamb_forbidden_text,
    process_lamb_workbook,
)
from scripts.meadow_wood_book import (
    is_meadow_wood_workbook,
    meadow_wood_forbidden_text,
    meadow_builder_display_name,
    process_meadow_wood_workbook,
)
from scripts.red_barn_book import (
    is_red_barn_workbook,
    process_red_barn_workbook,
    red_barn_forbidden_text,
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
        help="Worksheet name for single-sheet .xlsx (ignored for multi-sheet books)",
    )
    args = parser.parse_args()

    if args.multiplier <= 0:
        print("Multiplier must be positive.", file=sys.stderr)
        return 1

    source = args.file.expanduser().resolve()
    if not source.exists():
        print(f"File not found: {source}", file=sys.stderr)
        return 1

    updated = args.date or date.today().isoformat()
    try:
        date.fromisoformat(updated)
    except ValueError:
        print("--date must be YYYY-MM-DD", file=sys.stderr)
        return 1

    builder_name = args.builder
    if source.is_file() and is_meadow_wood_workbook(source):
        if args.builder.lower() in ("meadow wood", "meadow"):
            builder_name = meadow_builder_display_name()

    assert_no_approx(builder_name, args.note or "")

    slug = slugify(builder_name)
    pdf_name = f"{slug}.pdf"
    pdf_path = SHEETS_DIR / pdf_name
    rel_pdf = f"sheets/{pdf_name}"

    if is_forks_valley_bundle(source):
        processed = process_forks_valley_bundle(source, multiplier=args.multiplier)
        bad = forks_valley_forbidden_text(processed)
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
    elif is_red_barn_workbook(source):
        processed = process_red_barn_workbook(source, multiplier=args.multiplier)
        bad = red_barn_forbidden_text(processed)
        if bad:
            print(
                "Retail PDF would contain dealer-only wording: "
                + "; ".join(bad[:3]),
                file=sys.stderr,
            )
            return 1
        pages = build_retail_pdf_processed(
            processed,
            builder_name=builder_name,
            updated=updated,
            output_path=pdf_path,
        )
        print(f"Wrote {pages}-page retail PDF", file=sys.stderr)
    elif is_meadow_wood_workbook(source):
        processed = process_meadow_wood_workbook(
            source, multiplier=args.multiplier
        )
        bad = meadow_wood_forbidden_text(processed)
        if bad:
            print(
                "Retail PDF would contain dealer-only wording: "
                + "; ".join(bad[:3]),
                file=sys.stderr,
            )
            return 1
        pages = build_retail_pdf_processed(
            processed,
            builder_name=builder_name,
            updated=updated,
            output_path=pdf_path,
        )
        print(f"Wrote {pages}-page retail PDF", file=sys.stderr)
    elif is_lamb_workbook(source):
        processed = process_lamb_workbook(source, multiplier=args.multiplier)
        bad = lamb_forbidden_text(processed)
        if bad:
            print(
                "Retail PDF would contain dealer-only wording: "
                + "; ".join(bad[:3]),
                file=sys.stderr,
            )
            return 1
        pages = build_retail_pdf_processed(
            processed,
            builder_name=builder_name,
            updated=updated,
            output_path=pdf_path,
        )
        print(f"Wrote {pages}-page retail PDF", file=sys.stderr)
    elif is_farmside_workbook(source):
        processed = process_farmside_workbook(
            source, builder_name=args.builder, multiplier=args.multiplier
        )
        bad = farmside_forbidden_text(processed)
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
    elif is_fn_chairs_workbook(source):
        processed = process_fn_chairs_workbook(
            source, builder_name=args.builder, multiplier=args.multiplier
        )
        bad = fn_chairs_forbidden_text(processed)
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
        try:
            grid = read_dealer_file(source, sheet_name=args.sheet)
        except DealerFileError as exc:
            print(str(exc), file=sys.stderr)
            return 1

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
        builder_name,
        rel_pdf,
        updated,
        args.note,
    )

    action = "Replaced" if replaced else "Added"
    print(f"{action} {entry['name']} → {rel_pdf} (updated {entry['updated']})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
