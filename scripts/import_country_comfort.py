#!/usr/bin/env python3
"""Import Country Comfort retail sheet from finished wholesale workbooks (×2.7)."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from scripts.builders_data import slugify, upsert_builder, validate_index_note
from scripts.country_comfort_book import (
    BUILDER,
    MULTIPLIER,
    country_comfort_forbidden,
    process_country_comfort,
)
from scripts.multipliers_data import set_stored_multiplier
from scripts.pdf_builder import build_retail_pdf_processed
from scripts.pdf_overlap_check import assert_pdfs_no_overlap
from scripts.pdf_quality_check import assert_pdfs_quality

SHEETS_DIR = REPO_ROOT / "sheets"


def main() -> int:
    parser = argparse.ArgumentParser(description="Import Country Comfort retail PDF")
    parser.add_argument("main_xlsx", type=Path, help="2026 CC Prices Excel.xlsx")
    parser.add_argument(
        "laurel_xlsx",
        nargs="?",
        type=Path,
        default=None,
        help="Laurel Bay & Brantley Prices.xlsx",
    )
    parser.add_argument("--date", default=date.today().isoformat())
    parser.add_argument(
        "--price-list-date",
        default="2026-09-23",
        help="Workbook last-modified stand-in (undated printed list)",
    )
    parser.add_argument(
        "--note",
        default="Undated printed list; workbook dated 2026-09-23 (Laurel Bay & Brantley 2026-11-02)",
    )
    parser.add_argument("--force-replace", action="store_true")
    args = parser.parse_args()

    main_xlsx = args.main_xlsx.expanduser().resolve()
    if not main_xlsx.is_file():
        print(f"Not found: {main_xlsx}", file=sys.stderr)
        return 1
    laurel = args.laurel_xlsx.expanduser().resolve() if args.laurel_xlsx else None
    if args.laurel_xlsx and not laurel.is_file():
        print(f"Not found: {laurel}", file=sys.stderr)
        return 1

    try:
        validate_index_note(args.note)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    processed, stats = process_country_comfort(
        main_xlsx, laurel, multiplier=MULTIPLIER
    )
    bad = country_comfort_forbidden(processed)
    if bad:
        print("Forbidden wording:", "; ".join(bad[:5]), file=sys.stderr)
        return 1

    slug = slugify(BUILDER)
    pdf_path = SHEETS_DIR / f"{slug}.pdf"
    pages = build_retail_pdf_processed(
        processed,
        builder_name=BUILDER,
        updated=args.date,
        output_path=pdf_path,
        all_landscape=True,
    )
    assert_pdfs_no_overlap([pdf_path])
    assert_pdfs_quality([pdf_path])

    upsert_builder(
        BUILDER,
        f"sheets/{slug}.pdf",
        args.date,
        note=args.note,
        price_list_date=args.price_list_date,
    )
    set_stored_multiplier(BUILDER, MULTIPLIER)

    report = {
        "pages": pages,
        "items": stats["items"],
        "collections": stats["collections"],
        "pdf": str(pdf_path),
    }
    print(json.dumps(report, indent=2))
    print(
        f"Imported {BUILDER}: {pages} pages, {stats['items']} priced items",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
