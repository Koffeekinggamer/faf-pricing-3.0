#!/usr/bin/env python3
"""Add or refresh a builder retail price sheet on Pricing 3.0.

Standing rule: when the builder already exists in builders.json, replace its PDF
only if the incoming upload's price list effective/printed date is newer than the
stored price_list_date. Skip (no file changes) when the date is the same, older,
or cannot be determined. Pass --price-list-date YYYY-MM-DD when auto-detect fails.
"""

from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from scripts.builders_data import REPO_ROOT as ROOT, find_builder_by_name, slugify, upsert_builder
from scripts.price_list_date import (
    evaluate_price_list_update,
    infer_from_source,
    parse_iso_price_list_date,
)
from scripts.multipliers_data import resolve_multiplier, set_stored_multiplier
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
from scripts.troyer_book import (
    is_j_troyer_bundle,
    j_troyer_forbidden_text,
    process_j_troyer_bundle,
)
from scripts.frog_pond_book import (
    frog_pond_forbidden_text,
    is_frog_pond_workbook,
    process_frog_pond_workbook,
)
from scripts.townline_book import (
    is_townline_workbook,
    process_townline_workbook,
    townline_forbidden_text,
)
from scripts.woodside_book import (
    is_woodside_workbook,
    process_woodside_workbook,
    woodside_forbidden_text,
)
from scripts.country_view_oak_book import (
    country_view_oak_forbidden,
    is_country_view_oak_pdf,
    process_country_view_oak_pdf,
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
        default=None,
        help="Retail multiplier (wholesale × multiplier, round up). "
        "Omit for an existing builder to use the saved multiplier.",
    )
    parser.add_argument(
        "--override-multiplier",
        action="store_true",
        help="Allow changing the saved multiplier for this builder.",
    )
    parser.add_argument("--file", required=True, type=Path, help="Dealer source file")
    parser.add_argument("--note", default=None, help="Optional short note on the index")
    parser.add_argument(
        "--date",
        default=None,
        help="Updated date YYYY-MM-DD (default: today)",
    )
    parser.add_argument(
        "--price-list-date",
        default=None,
        metavar="YYYY-MM-DD",
        help="Printed/effective price list date (default: auto-detect from source)",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Worksheet name for single-sheet .xlsx (ignored for multi-sheet books)",
    )
    args = parser.parse_args()

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

    try:
        multiplier, persist_mult = resolve_multiplier(
            builder_name,
            args.multiplier,
            override=args.override_multiplier,
        )
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    assert_no_approx(builder_name, args.note or "")

    existing = find_builder_by_name(builder_name)
    if args.price_list_date:
        incoming_pld = parse_iso_price_list_date(args.price_list_date)
        if incoming_pld is None:
            print("--price-list-date must be YYYY-MM-DD", file=sys.stderr)
            return 1
    else:
        incoming_pld = infer_from_source(source, builder_name, sheet=args.sheet)

    allow, pld_reason = evaluate_price_list_update(existing, incoming_pld)
    if not allow:
        print(f"Skipped {builder_name}: {pld_reason}", file=sys.stderr)
        print(f"Skipped {builder_name}: {pld_reason}")
        return 2 if incoming_pld is None else 0

    slug = slugify(builder_name)
    pdf_name = f"{slug}.pdf"
    pdf_path = SHEETS_DIR / pdf_name
    rel_pdf = f"sheets/{pdf_name}"

    if is_country_view_oak_pdf(source):
        processed, stats = process_country_view_oak_pdf(
            source, multiplier=multiplier
        )
        bad = country_view_oak_forbidden(processed)
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
        print(
            f"Wrote {pages}-page retail PDF "
            f"({stats['expected_wholesale_dollars']} prices @ {multiplier})",
            file=sys.stderr,
        )
    elif is_forks_valley_bundle(source):
        processed = process_forks_valley_bundle(source, multiplier=multiplier)
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
    elif is_j_troyer_bundle(source):
        processed = process_j_troyer_bundle(source, multiplier=multiplier)
        bad = j_troyer_forbidden_text(processed)
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
    elif is_frog_pond_workbook(source):
        processed = process_frog_pond_workbook(source, multiplier=multiplier)
        bad = frog_pond_forbidden_text(processed)
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
    elif is_woodside_workbook(source):
        processed = process_woodside_workbook(source, multiplier=multiplier)
        bad = woodside_forbidden_text(processed)
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
    elif is_townline_workbook(source):
        processed = process_townline_workbook(source, multiplier=multiplier)
        bad = townline_forbidden_text(processed)
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
    elif is_red_barn_workbook(source):
        processed = process_red_barn_workbook(source, multiplier=multiplier)
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
            source, multiplier=multiplier
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
        processed = process_lamb_workbook(source, multiplier=multiplier)
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
            source, builder_name=args.builder, multiplier=multiplier
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
            source, builder_name=args.builder, multiplier=multiplier
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
                multiplier=multiplier,
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
                multiplier=multiplier,
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
            grid = apply_multiplier_to_grid(grid, multiplier)
            build_retail_pdf(
                grid,
                builder_name=args.builder,
                updated=updated,
                output_path=pdf_path,
            )

    pld_iso = incoming_pld.isoformat() if incoming_pld else None
    entry, replaced = upsert_builder(
        builder_name,
        rel_pdf,
        updated,
        args.note,
        price_list_date=pld_iso,
    )

    if persist_mult:
        set_stored_multiplier(builder_name, multiplier)

    action = "Replaced" if replaced else "Added"
    mult_note = f" (multiplier {multiplier})" if args.multiplier is None else ""
    print(
        f"{action} {entry['name']} → {rel_pdf} (updated {entry['updated']}){mult_note}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
