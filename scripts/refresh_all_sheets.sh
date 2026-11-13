#!/usr/bin/env bash
# Regenerate retail PDFs from local upload sources (not committed).
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
UP="${UPLOAD_DIR:-$HOME/.cursor/projects/workspace/uploads}"
cd "$ROOT"
TODAY="$(date +%Y-%m-%d)"

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

B1="/tmp/viztech_batch01/batch01"
B2="/tmp/viztech_batch02/batch02"
EB="/tmp/viztech_ebony/ebony"
mkdir -p "$(dirname "$B1")" "$(dirname "$B2")" "$(dirname "$EB")"
unzip -q -o "$UP/batch01_411f.zip" -d /tmp/viztech_batch01
unzip -q -o "$UP/batch02_7dee.zip" -d /tmp/viztech_batch02
unzip -q -o "$UP/ebony_f099.zip" -d /tmp/viztech_ebony

if [[ ! -s "$B2/manifest.tsv" ]]; then
  cat >"$B2/manifest.tsv" <<'TSV'
builder_name	slug	file	multiplier	notes
Buckeye Rockers	buckeye-rockers	buckeye-rockers__AL_2026_Pricelist_0426.xlsx	2.7	Authentex + standard
Buckeye Rockers	buckeye-rockers	buckeye-rockers__BR_2026_Pricleist_1026.xlsx	2.7	Authentex + standard
Buss Woodcraft	buss-woodcraft	buss-woodcraft__BW_2026_Pricelist_0426.xlsx	2.7	
Cloverleaf Woodworking	cloverleaf-woodworking	cloverleaf-woodworking__CLWW_2026_Pricelist_0226.xlsx	2.7	
Comfort Time Furniture	comfort-time-furniture-yoders-woodshop	comfort-time-furniture-yoders-woodshop__CTF_2026_Pricing_0326.xlsx	2.7	
Country Cupola	cupola-furniture	cupola-furniture__CF_2026_Pricelist_0326.xlsx	2.7	Standard + LuxWood
Country Cupola	cupola-furniture	cupola-furniture__CF_2026_LuxWood_Pricelist_0326.xlsx	2.7	Standard + LuxWood
Country Road Furniture	country-road-furniture	country-road-furniture__CRF_2026_Pricelist_0226.xlsx	2.7	
Country Side Rustic Log Furniture	country-side-rustic-log-furniture	country-side-rustic-log-furniture__2026-Wholesale-Price-List-with-Markup-Calculator.xlsx	2.7	
Country View Furniture	country-view-furniture	country-view-furniture__CVF_2026_Pricelist_0526.xlsx	2.7	
Countryside Tables	countryside-tables	countryside-tables__CST_2026_Pricelist_0326.xlsx	2.7	
Curveside Wood Products	curveside-wood-products	curveside-wood-products__CSWP_2026_Pricelist_1026.xls	2.7	
D&E Workshop	de-workshop	de-workshop__DEW_2026_Pricelist_0526.xlsx	2.7	
D&N Wood Design	dn-wood-design	dn-wood-design__DNWD_2026_Pricelist_0326.xlsx	2.7	
Dover Road Wood	dover-road-wood	dover-road-wood__DRW_2026_Bath_Cabinets_Pricelist_1026.xlsx	2.7	
Dover Road Wood	dover-road-wood	dover-road-wood__DRW_2026_Bedroom_Tables_Barstools_Pricelist_1026.xlsx	2.7	
Dover Road Wood	dover-road-wood	dover-road-wood__DRW_2026_Work_Stations_Pricelist_0226.xlsx	2.7	
Dutch Boy Poly	dutch-boy-poly	dutch-boy-poly__DBP_2026-2027_Pricelist_0926.xlsx	2.7	
TSV
fi
if [[ ! -s "$EB/manifest.tsv" ]]; then
  cat >"$EB/manifest.tsv" <<'TSV'
builder_name	slug	file	multiplier	notes
Ebony Woodworking	ebony-woodworking	ebony-woodworking__EWW_2025_Pricelist_0825.xlsx	2.7	2025 list
TSV
fi

if [[ -s "$B1/manifest.tsv" ]]; then
  echo "=== Viztech batch 01 ==="
  python3 scripts/import_viztech_batch.py --manifest "$B1/manifest.tsv" --sources "$B1" --date "$TODAY"
fi
if [[ -s "$B2/manifest.tsv" ]]; then
  echo "=== Viztech batch 02 ==="
  python3 scripts/import_viztech_batch.py --manifest "$B2/manifest.tsv" --sources "$B2" --date "$TODAY"
fi
if [[ -s "$EB/manifest.tsv" ]]; then
  echo "=== Viztech Ebony ==="
  python3 scripts/import_viztech_batch.py --manifest "$EB/manifest.tsv" --sources "$EB" --date "$TODAY"
fi

echo "Done."
