#!/usr/bin/env python3
"""Import Viztech batch from manifest.tsv (sources stay outside repo)."""

from __future__ import annotations

import argparse
import csv
import json
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from scripts.builders_data import slugify, upsert_builder
from scripts.multipliers_data import set_stored_multiplier
from scripts.pdf_builder import build_retail_pdf_processed
from scripts.pdf_overlap_check import assert_pdfs_no_overlap
from scripts.pdf_quality_check import assert_pdfs_quality
from scripts.pricing_lib import assert_no_approx
from scripts.viztech_book import (
    merge_builder_files,
    normalize_builder_name,
    viztech_forbidden,
)

SHEETS_DIR = REPO_ROOT / "sheets"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--sources", type=Path, required=True, help="Directory with workbook files")
    parser.add_argument("--date", default=date.today().isoformat())
    args = parser.parse_args()

    sources = args.sources.resolve()
    groups: dict[str, list[dict]] = defaultdict(list)
    with args.manifest.open(encoding="utf-8-sig") as f:
        for row in csv.DictReader(f, delimiter="\t"):
            bname = normalize_builder_name(row["builder_name"])
            groups[bname].append(row)

    report = {"live": [], "skipped": [], "spot_checks": {}}
    pdf_paths: list[Path] = []

    for builder_name in sorted(groups.keys(), key=lambda s: s.casefold()):
        files: list[Path] = []
        for row in groups[builder_name]:
            p = sources / row["file"]
            if not p.is_file():
                report["skipped"].append({"builder": builder_name, "reason": f"missing {row['file']}"})
                continue
            files.append(p)

        mult = float(groups[builder_name][0].get("multiplier") or 2.7)
        assert_no_approx(builder_name)
        res = merge_builder_files(files, builder_name=builder_name, multiplier=mult)
        if res.warnings and not res.rows:
            report["skipped"].append(
                {"builder": builder_name, "reason": "; ".join(res.warnings)}
            )
            continue
        if not res.rows:
            report["skipped"].append({"builder": builder_name, "reason": "no rows"})
            continue

        bad = viztech_forbidden(res.rows)
        if bad:
            report["skipped"].append(
                {"builder": builder_name, "reason": f"forbidden text: {bad[:2]}"}
            )
            continue

        slug = slugify(builder_name)
        pdf_path = SHEETS_DIR / f"{slug}.pdf"
        pages = build_retail_pdf_processed(
            res.rows,
            builder_name=builder_name,
            updated=args.date,
            output_path=pdf_path,
        )
        pdf_paths.append(pdf_path)
        upsert_builder(builder_name, f"sheets/{slug}.pdf", args.date, None)
        set_stored_multiplier(builder_name, mult)
        report["live"].append({"name": builder_name, "pages": pages, "pdf": str(pdf_path)})
        report["spot_checks"][builder_name] = [
            {"item": s.item, "wholesale": s.wholesale, "retail": s.retail} for s in res.spots
        ]
        if res.warnings:
            report.setdefault("warnings", []).append(
                {builder_name: res.warnings}
            )

    if pdf_paths:
        assert_pdfs_no_overlap(pdf_paths)
        assert_pdfs_quality(pdf_paths)

    out = REPO_ROOT / "data" / "viztech_batch_import_report.json"
    out.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))
    return 0 if report["live"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
