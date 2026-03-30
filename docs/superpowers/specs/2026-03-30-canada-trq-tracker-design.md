# Canada TRQ Weekly Tracker — Design Spec

## Overview

A single Python script that downloads Canadian steel Tariff Rate Quota (TRQ) utilization data and B1 import data weekly, parsing them into an Excel workbook that matches Laura's template format. Automated via GitHub Actions on a Monday schedule.

## Problem Statement

Laura (US-based analyst) needs to monitor how Canadian steel import quotas are being utilized over time. The government publishes TRQ utilization data at `eics-scei.gc.ca` but as complex report-formatted files (550+ merged cells per file). She also wants to cross-reference quota utilization with actual import volumes from the B1 report.

Currently there is no monitoring — data must be manually downloaded and reformatted each time.

## Requirements

### From Laura's letter and template:
1. **Weekly Monday updates** — grab latest data each Monday morning
2. **8 product categories only** (not all 23 in the TRQ system):
   - Flat: Hot-Rolled Sheet, Steel Plate, Cold-Rolled Sheet, Coated Steel Sheet
   - Long: Rebar, Hot-Rolled Bar, Wire Rod, Structural Steel
3. **Country-level breakdown** — show each country's share of utilization
4. **Time-series layout** — each Monday = new column showing utilization %
5. **Both FTA and non-FTA** — separate sheets for each
6. **B1 import data** — match actual imports by HTS codes alongside quota utilization
7. **OVER flag** — indicate when a country exceeds its maximum share cap
8. **Quarterly transition** — detect when new quarter data appears, create new sheets

## Architecture

### Single script: `canada_trq_tracker.py`

```
main()
  1. detect_current_quarter()    -> "Q4", date range, URLs
  2. download_trq_csv("FTA")     -> raw CSV text
  3. download_trq_csv("NFTA")    -> raw CSV text
  4. parse_trq_csv(csv_text)     -> {product: {country: share%}}
  5. scrape_b1_imports()         -> {product: {country: tonnes, value}}
  6. load_or_create_workbook()   -> openpyxl Workbook
  7. update_trq_sheet(wb, ...)   -> adds date column
  8. update_b1_sheet(wb, ...)    -> updates B1 sheet
  9. wb.save()
```

### File structure

```
Canada-Quota/
  canada_trq_tracker.py       # Main script (~300-400 lines)
  requirements.txt             # requests, beautifulsoup4, openpyxl
  .github/workflows/
    weekly_scrape.yml           # GitHub Actions weekly cron
  data/
    canada_trq_tracker.xlsx     # Output (auto-generated, committed to repo)
  Canadian Quota Template.xlsx  # Laura's reference template
  Small Task/                   # Original analysis files (kept for reference)
```

## Data Sources

### Source 1: TRQ Utilization CSVs

**URLs:**
- `https://www.eics-scei.gc.ca/report-rapport/TRQ_FTA-{quarter}.csv`
- `https://www.eics-scei.gc.ca/report-rapport/TRQ_NFTA-{quarter}.csv`

Where `{quarter}` is `P1`, `Q2`, `Q3`, `Q4`, etc.

**CSV format** (actual structure from analysis):
- Lines 1-4: Report headers with auto-generated SSRS field names
- Lines 5-27: Part A — summary row per product (23 products)
  - Format: `Product Category,Maximum quota (KGM),Maximum country share (%),Current utilization (KGM),Current utilization (%),Remaining quota (KGM),{item_number},{product_name},"{max_quota}",{max_share}%,"{util_kgm}",{util_pct}%,"{remaining}"`
- Lines 29+: Part B — country-level breakdown sections
  - Each section: metadata header line, then country data rows, then blank line
  - Country row format: `Country,Current utilization (KGM),Share of total utilization (%),{country},"{country_kgm}",{share_pct}%,"{total_kgm}",{total_pct}%`

**Parsing strategy:**
1. Parse Part A to build product index: `{product_name: {item_number, max_quota, max_share, total_util_kgm}}`
2. Parse Part B sections — each country row contains the section's total utilization KGM
3. Match Part B sections to products by comparing total utilization KGM against Part A values
4. Extract `share_pct` for each country in each of the 8 tracked products

### Source 2: B1 Import Data

**URL:** `https://www.eics-scei.gc.ca/report-rapport/b1.htm`

**Format:** HTML table with columns: Month | Country of Origin | Tonnes | C$1000 | C$/Tonne, grouped by HS code.

**HTS code mapping** (from Laura's template "HTS code covered" sheet):

| Product Category | HTS Codes |
|-----------------|-----------|
| Hot-Rolled Sheet | 7208.10.00.00, 7208.25.00.00, 7208.26.00.00, 7208.27.00.00, 7208.36.00.00, 7208.37.00.00, 7208.38.00.00, 7208.39.00.00 |
| Steel Plate | 7208.40.00.00, 7208.51.00.00, 7208.52.00.00 |
| Cold-Rolled Sheet | 7209.15.00.00, 7209.16.00.00, 7209.17.00.00, 7209.18.00.00, 7209.25.00.00, 7209.26.00.00, 7209.27.00.00, 7209.28.00.00 |
| Coated Steel Sheet | 7210.41.00.00, 7210.49.00.00, 7210.61.00.00, 7210.69.00.00, 7210.70.00.00, 7210.90.00.00 |
| Rebar | 7214.20.00.00 |
| Hot-Rolled Bar | 7214.30.00.00, 7214.91.00.00, 7214.99.00.00 |
| Wire Rod | 7213.10.00.00, 7213.20.00.00, 7213.91.00.00, 7213.99.00.00 |
| Structural Steel | 7216.10.00.00, 7216.21.00.00, 7216.22.00.00, 7216.31.00.00, 7216.32.00.00, 7216.33.00.00, 7216.40.00.00, 7216.50.00.00, 7216.69.00.00 |

**Scraping strategy:**
1. Fetch B1 HTML with requests
2. Parse with BeautifulSoup
3. Extract rows matching the HTS codes above (normalize dotted format to B1's plain format)
4. Aggregate tonnes and C$1000 by product category and country, across all months in the quarter

## Excel Output Format

### Sheet: "non-FTA {quarter}" and "FTA {quarter}"

Matches Laura's template exactly:

```
Row 1: "Non-FTA Quota - Quarter 4: March 26, 2026 to June 25, 2026"
Row 2: [blank] | Country | Max Quota (KG) | Max Share (%) | Apr 6 2026 | Apr 13 2026 | ... | OVER

Row 3:  Hot-Rolled Sheet | China          | 2,370,500 | 41% | 2.95% | ... | NO
Row 4:  Hot-Rolled Sheet | India          | 2,370,500 | 41% | 0.11% | ... | NO
Row 5:  Hot-Rolled Sheet | Taiwan         | 2,370,500 | 41% | 40.92% | ... | YES
Row 6:  Hot-Rolled Sheet | TOTAL          |           |     | 43.97% | ... |
Row 7:  Steel Plate      | China          | 5,201,200 | 36% | 36.00% | ... | YES
...
```

- **Column A**: Product category name (repeated for each country row within that product)
- **Column B**: Country name, with "TOTAL" row at end of each product group
- **Column C**: Max Quota (KG) — static, same for all countries within a product
- **Column D**: Max Share (%) — static, the cap for any single country
- **Columns E+**: Weekly date columns — each contains the country's `Share of total utilization (%)`
- **Last column**: "OVER" — "YES" if any weekly value exceeds Max Share, else blank

Each new script run appends a new date column. Existing columns are never modified.

### Sheet: "B1 Imports"

Aggregated import volumes by product category and country:

```
Row 1: "B1 Import Data - Quarter 4: March 26, 2026 to June 25, 2026"
Row 2: Product Category | Country | Total Tonnes | Total C$1000 | Avg C$/Tonne

Row 3: Hot-Rolled Sheet | China    | 15,234 | 12,456 | 817.89
Row 4: Hot-Rolled Sheet | Japan    | 8,901  | 9,234  | 1,037.41
...
```

Updated each run with latest B1 data (B1 data is cumulative for the quarter).

### Sheet: "HTS code covered"

Static reference sheet copied from Laura's template — lists HTS codes per product category.

## Quarterly Transition Logic

**Quarter date ranges:**
- Q1/P1: ~late June/August to September 25
- Q2: September 26 to December 25
- Q3: December 26 to March 25
- Q4: March 26 to June 25

**Auto-detection:** Based on today's date, calculate expected quarter. Construct URL and attempt download.

**Fallback:** If the current quarter's CSV returns HTTP 404 (data not posted yet), try the previous quarter. Log a message: "Q4 data not yet available, using Q3."

**New sheets:** When the quarter changes from the existing sheets, create new sheets (e.g., "non-FTA Q4", "FTA Q4"). Previous quarter sheets remain as historical records.

## GitHub Actions Workflow

```yaml
name: Weekly Canada TRQ Update
on:
  schedule:
    - cron: '0 12 * * 1'    # Every Monday at 12:00 UTC (~8:00 AM ET)
  workflow_dispatch:          # Manual trigger
jobs:
  update:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'
      - run: pip install -r requirements.txt
      - run: python canada_trq_tracker.py
      - run: |
          git config user.name "GitHub Actions"
          git config user.email "actions@github.com"
          git add data/
          git diff --staged --quiet || git commit -m "Weekly TRQ update $(date +%Y-%m-%d)"
          git push
```

## Error Handling

| Scenario | Behavior |
|----------|----------|
| CSV download fails (network error) | Retry once after 5 seconds. If still fails, exit with error message. |
| B1 page unreachable | Log warning, skip B1 update, still update TRQ sheets. |
| CSV format changes | Log the unexpected format and exit with a descriptive error. |
| Product has zero utilization (no countries) | Show only the product name with TOTAL = 0.00%. |
| Duplicate column (same date already exists) | Skip adding a new column, log "already up to date". |
| New country appears mid-quarter | Append new row at the end of the product group, backfill previous weeks as blank. |

## Dependencies

```
requests>=2.31.0
beautifulsoup4>=4.12.0
openpyxl>=3.1.0
```

No pandas needed — openpyxl handles Excel I/O directly, and the CSV parsing is custom (non-standard format).
