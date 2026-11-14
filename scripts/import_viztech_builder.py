#!/usr/bin/env python3
"""Import a single Viztech-style builder workbook (finished wholesale × multiplier)."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from scripts.builders_data import find_builder_by_name, slugify, upsert_builder
from scripts.indian_trail_book import (
    BUILDER as INDIAN_TRAIL,
    process_indian_trail_workbook,
)
from scripts.jr_woodworking_book import (
    BUILDER as JR_BUILDER,
    process_jr_woodworking_workbook,
)
from scripts.multipliers_data import set_stored_multiplier
from scripts.pdf_builder import build_retail_pdf_processed
from scripts.pdf_overlap_check import assert_pdfs_no_overlap
from scripts.pdf_quality_check import assert_pdfs_quality
from scripts.price_list_date import evaluate_price_list_update
from scripts.pricing_lib import assert_no_approx
from scripts.viztech_book import viztech_forbidden

SHEETS_DIR = REPO_ROOT / "sheets"

PROCESSORS = {
    "indian-trail": (INDIAN_TRAIL, process_indian_trail_workbook, "2026-01-01"),
    "jr": (JR_BUILDER, process_jr_woodworking_workbook, None),
}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--kind",
        required=True,
        choices=sorted(PROCESSORS.keys()),
        help="Builder workbook layout",
    )
    parser.add_argument("--file", type=Path, required=True)
    parser.add_argument("--multiplier", type=float, default=2.7)
    parser.add_argument("--date", default=date.today().isoformat())
    parser.add_argument("--price-list-date", default=None)
    parser.add_argument("--force-replace", action="store_true")
    args = parser.parse_args()

    path = args.file.expanduser().resolve()
    if not path.is_file():
        print(f"Not found: {path}", file=sys.stderr)
        return 1

    builder, processor, default_pld = PROCESSORS[args.kind]
    assert_no_approx(builder)
    incoming_pld = None
    if args.price_list_date:
        incoming_pld = date.fromisoformat(args.price_list_date)
    elif default_pld:
        incoming_pld = date.fromisoformat(default_pld)

    existing = find_builder_by_name(builder)
    allow, reason = evaluate_price_list_update(existing, incoming_pld)
    if args.force_replace and existing:
        allow = True
    if not allow:
        print(f"Skipped {builder}: {reason}", file=sys.stderr)
        return 2 if incoming_pld is None and existing else 0

    res = processor(path, multiplier=args.multiplier)
    if res.warnings and not res.rows:
        print("; ".join(res.warnings), file=sys.stderr)
        return 1
    if not res.rows:
        print("No content", file=sys.stderr)
        return 1

    bad = viztech_forbidden(res.rows)
    if bad:
        print("Forbidden:", "; ".join(bad[:5]), file=sys.stderr)
        return 1

    slug = slugify(builder)
    pdf_path = SHEETS_DIR / f"{slug}.pdf"
    pages = build_retail_pdf_processed(
        res.rows,
        builder_name=builder,
        updated=args.date,
        output_path=pdf_path,
    )
    assert_pdfs_no_overlap([pdf_path])
    assert_pdfs_quality([pdf_path])

    pld_iso = incoming_pld.isoformat() if incoming_pld else None
    upsert_builder(builder, f"sheets/{slug}.pdf", args.date, None, price_list_date=pld_iso)
    set_stored_multiplier(builder, args.multiplier)

    report = {
        "builder": builder,
        "pages": pages,
        "spot_checks": [
            {"item": s.item, "wholesale": s.wholesale, "retail": s.retail}
            for s in res.spots[:6]
        ],
    }
    print(json.dumps(report, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())