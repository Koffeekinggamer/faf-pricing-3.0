#!/usr/bin/env python3
"""One-shot / maintenance: set price_list_date on builders.json from retail PDFs."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from scripts.builders_data import REPO_ROOT, load_builders, save_builders
from scripts.price_list_date import backfill_from_retail_pdf


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--write",
        action="store_true",
        help="Update data/builders.json (default: report only)",
    )
    args = parser.parse_args()

    builders = load_builders()
    missing: list[str] = []
    updated = 0
    for b in builders:
        pdf = REPO_ROOT / b.get("pdf", "")
        d = backfill_from_retail_pdf(pdf)
        if d:
            iso = d.isoformat()
            if b.get("price_list_date") != iso:
                b["price_list_date"] = iso
                updated += 1
        else:
            missing.append(b["name"])

    report = {
        "backfilled_count": len(builders) - len(missing),
        "missing_count": len(missing),
        "missing_builders": sorted(missing),
        "changed": updated,
    }
    print(json.dumps(report, indent=2))
    if args.write:
        save_builders(builders)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
