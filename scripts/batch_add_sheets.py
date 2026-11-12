#!/usr/bin/env python3
"""Batch refresh builder PDFs from a JSON manifest (sources stay outside the repo)."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run add_sheet.py for each entry in a manifest JSON file"
    )
    parser.add_argument(
        "--manifest",
        required=True,
        type=Path,
        help="JSON array: [{builder, file, multiplier?, finished_only?, note?, date?}, ...]",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print commands without running",
    )
    args = parser.parse_args()
    path = args.manifest.expanduser().resolve()
    if not path.is_file():
        print(f"Manifest not found: {path}", file=sys.stderr)
        return 1
    entries = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(entries, list):
        print("Manifest must be a JSON array", file=sys.stderr)
        return 1

    add_sheet = REPO_ROOT / "scripts" / "add_sheet.py"
    rc = 0
    for i, row in enumerate(entries):
        if not isinstance(row, dict):
            print(f"Row {i}: not an object", file=sys.stderr)
            rc = 1
            continue
        builder = row.get("builder")
        source = row.get("file")
        if not builder or not source:
            print(f"Row {i}: missing builder or file", file=sys.stderr)
            rc = 1
            continue
        cmd = [sys.executable, str(add_sheet), "--builder", str(builder), "--file", str(source)]
        mult = row.get("multiplier")
        if mult is not None:
            cmd.extend(["--multiplier", str(mult)])
        elif row.get("use_saved_multiplier", True):
            pass
        if row.get("override_multiplier"):
            cmd.append("--override-multiplier")
        if row.get("note"):
            cmd.extend(["--note", str(row["note"])])
        if row.get("date"):
            cmd.extend(["--date", str(row["date"])])
        if row.get("sheet"):
            cmd.extend(["--sheet", str(row["sheet"])])
        fin = row.get("finished_only")
        if fin is False:
            print(
                f"Warning: {builder}: finished_only=false not supported yet; processors enforce finished-only",
                file=sys.stderr,
            )
        if args.dry_run:
            print(" ".join(cmd))
            continue
        print(f"--- {builder} ---", file=sys.stderr)
        r = subprocess.run(cmd, cwd=REPO_ROOT)
        if r.returncode:
            rc = r.returncode
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
