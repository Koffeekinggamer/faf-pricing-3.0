# FAF Pricing 3.0

Internal **builder price sheet** index for [Foothills Amish Furniture](https://foothillsamishfurniture.com/) (Landrum, SC). Static site on GitHub Pages plus a small Python tool that turns dealer wholesale files into branded retail PDFs.

## Live site

After merging to `main` and enabling GitHub Pages (source: **GitHub Actions**), the site is served from this repo’s default Pages URL:

`https://<github-owner>.github.io/<repo-name>/`

For this repository: `https://koffeekinggamer.github.io/faf-pricing-3.0/`

The page is `noindex` and `robots.txt` blocks crawlers.

## Local preview

```bash
python3 -m http.server 8080
# open http://127.0.0.1:8080/
```

## Add or refresh a builder

1. Install dependencies (from repo root):

   ```bash
   pip install -r requirements.txt
   ```

2. Run the add script with a dealer file (`.xlsx`, `.xls`, `.csv`, or text-extractable `.pdf`):

   ```bash
   python scripts/add_sheet.py \
     --builder "Look-alike Builder" \
     --multiplier 2.7 \
     --file tests/fixtures/look_alike_wholesale.csv \
     --note "Optional note for staff"
   ```

   On **first add**, pass `--multiplier` (saved in `data/multipliers.json`, not used by the public index). On **refresh**, omit `--multiplier` to reuse the saved value. To change a builder’s multiplier later, pass the new value with `--override-multiplier`.

3. Commit the updated `data/builders.json`, `data/multipliers.json` (when the multiplier changes), and `sheets/<slug>.pdf`, push to `main`, and wait for the Pages workflow.

### Pricing rules

- **Finished pricing only** — retail PDFs show **finished** prices. Drop unfinished price columns, unfinished-only sheets, and notes such as “X% less for unfinished.” When a builder lists only unfinished plus a finish upcharge, compute finished wholesale (unfinished + finish cost) once, then apply the multiplier.
- Retail = wholesale × multiplier, rounded **up** to the next whole dollar (e.g. $683.20 → $684).
- Output uses **whole dollars** only (no cents in transformed prices).
- The word **`approx`** must not appear in generated output.
- **Percent** add-ons stay as percentages; dimensions and item numbers are not multiplied.

Re-adding the same builder name **replaces** the PDF and updates metadata (no duplicate rows). `builders.json` is kept sorted A→Z (case-insensitive, leading punctuation ignored).

### Price list date (standing rule)

Each builder stores a **`price_list_date`** (YYYY-MM-DD): the effective or printed date on the dealer price list, not the day the PDF was generated (`updated`).

When a builder **already exists**, `add_sheet.py` and `import_viztech_batch.py` **replace the PDF only if the incoming list date is newer** than the stored `price_list_date`. If the date is the same or older, the run skips that builder and leaves files unchanged. If the incoming date **cannot be determined**, nothing is replaced (use `--price-list-date` on `add_sheet.py` when auto-detect fails). Viztech uploads infer the date from filename version stamps such as `_0526` (May 2026).

Backfill stored dates from existing retail PDF titles:

```bash
python scripts/backfill_price_list_dates.py --write
```

## Repository layout

| Path | Purpose |
|------|---------|
| `index.html` | Staff price sheet index |
| `data/builders.json` | Builder list (public index) |
| `data/multipliers.json` | Staff-only retail multipliers per builder (not read by the site) |
| `sheets/` | Retail PDFs |
| `assets/` | FAF logo, CSS, JS |
| `scripts/add_sheet.py` | CLI entrypoint |
| `tests/` | Unit tests and look-alike fixtures |

## Tests

```bash
pip install -r requirements.txt
pytest tests/ -q
```

## GitHub Pages setup (manual, once)

In the repository **Settings → Pages**:

- **Build and deployment** → Source: **GitHub Actions** (not “Deploy from a branch”).

Pushes to `main` run `.github/workflows/pages.yml` automatically.

## Input formats

| Format | Support |
|--------|---------|
| `.xlsx` | Supported (first worksheet) |
| `.xls` | Supported (first sheet) |
| `.csv` | Supported |
| `.pdf` | Supported only when text/tables are extractable; scanned image PDFs fail with a clear error |
| Other | Not supported yet (e.g. Word, Google Sheets export except via xlsx/csv) |
