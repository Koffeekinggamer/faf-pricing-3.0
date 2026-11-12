#!/usr/bin/env bash
# Regenerate retail PDFs from local upload sources (not committed).
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
UP="${UPLOAD_DIR:-$HOME/.cursor/projects/workspace/uploads}"
cd "$ROOT"

run() {
  local builder="$1"
  local source="$2"
  shift 2
  echo "=== $builder ==="
  python3 scripts/add_sheet.py --builder "$builder" --file "$source" "$@"
}

run "Artisan Chairs" "$UP/Artisan_Chairs_952f.xlsx"
run "Farmside Woodworking" "$UP/drop3_502c.xlsx"
run "FN Chairs" "$UP/FN_Chairs_39e9.zip"
run "Forks Valley Woodworking" "$UP/Forks_Valley_5aeb.zip"
run "Frog Pond Furniture" "$UP/Frog_Pond_a0de.xlsx"
run "J. Troyer and Co" "$UP/J_Troyer_858f.zip"
run "Lamb Woodworking" "$UP/Lamb_Woodworking_e69a.xlsx"
run "Meadow Wood" "$UP/Meadow_Wood_fe17.xlsx"
run "Millcraft" "$UP/Millcraft_aed4.xlsx"
run "Red Barn Woodworking" "$UP/Red_Barn_739a.zip"
run "Townline Furniture" "$UP/Townline_ba42.xlsx"
run "Woodside Woodworks" "$UP/Woodside_86ac.zip"

echo "Done."
