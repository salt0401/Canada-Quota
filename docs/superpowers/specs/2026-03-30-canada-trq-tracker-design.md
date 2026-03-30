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
  canada_trq_tracker.py       # Main script
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

**URLs (server is case-insensitive, so consistent uppercase works):**
- `https://www.eics-scei.gc.ca/report-rapport/TRQ_FTA-{quarter}.csv`
- `https://www.eics-scei.gc.ca/report-rapport/TRQ_NFTA-{quarter}.csv`

Where `{quarter}` is `Q2`, `Q3`, `Q4`, etc. Special case: FTA's first period uses `P1` instead of `Q1`.

**Verified:** Server returns identical content for `TRQ_nFTA-q2` and `TRQ_NFTA-Q2` (same file size). No casing fallback needed.

#### CSV Format Variants

There are two CSV format variants. The parser must detect and handle both.

**Detection:** Check line 1. If it starts with `ExecutionTime` → old format. If it starts with `Textbox` → new format.

##### Old Format (P1, Q1, Q2)

```
Line 1: ExecutionTime                                    ← identifier
Line 2: Report executed on : 03/29/2026 00:45:35 AM
Line 3: (blank)
Line 4: Textbox18,CONTROL_ITEMS_Level_1_Commodity_Code,...  ← column header IDs
Lines 5-27: Part A summary — 23 rows, fields:
    item_number, product_name, "max_quota", max_share%, "util_kgm"|"Max Utilized", util_pct%|"Max Utilized", "remaining"|0
Line 28: (blank)
Lines 29+: Part B country sections — metadata header, then data rows:
    country, "util_kgm", share_pct%, "total_kgm", total_pct%
```

##### New Format (Q3, Q4, and likely future quarters)

```
Line 1: Textbox137,Textbox143,...                         ← identifier
Line 2: Title + all 23 product section headers
Line 3: (blank)
Line 4: More Textbox IDs + CONTROL_ITEMS column names
Lines 5-27: Part A summary — 23 rows, fields with INLINE LABELS:
    "Product Category","Maximum quota (KGM)","Maximum country share (%)","Current utilization (KGM)","Current utilization (%)","Remaining quota (KGM)",item_number,product_name,"max_quota",max_share%,"util_kgm",util_pct%,"remaining"
Line 28: (blank)
Lines 29+: Part B country sections — metadata header, then data rows with INLINE LABELS:
    "Country","Current utilization (KGM)","Share of total utilization (%)",country,"util_kgm",share_pct%,"total_kgm",total_pct%
```

**Key difference:** New format prepends 6 label fields to Part A rows and 3 label fields to Part B rows. The actual data fields are in the same order.

**Parsing strategy (works for both formats):**
1. Detect format by line 1
2. Set field offset: old format offset=0, new format Part A offset=6, Part B offset=3
3. Parse Part A (lines 5-27): extract item_number, product_name, max_quota, max_share, total_util_kgm for each product. Store as ordered list (items 1-23).
4. Parse Part B sections: each section preceded by metadata header, separated by blank lines. Part B sections appear in the same sequential order as Part A items (1-23), skipping products with zero utilization that have no country data.
5. **Primary matching: positional.** Walk Part A items in order; for each item with non-zero utilization, consume the next Part B section. **Confirmation: compare `total_kgm`** in the country rows against Part A's `total_util_kgm` — log a warning if they differ.
6. Extract `share_pct` for each country in each of the 8 tracked products

**Special values to handle:**
- `"Max Utilized"` string (instead of numeric values) — treat as 100% utilization
- `"0.00%"` as string — strip `%` and parse as float 0.0
- Comma-separated thousands in quoted numbers: `"1,042,387"` — strip commas before parsing

### Source 2: B1 Import Data

**URL:** `https://www.eics-scei.gc.ca/report-rapport/b1.htm`

**Format:** Single large HTML table covering all steel HS codes for the current quarter. Columns:
- HS code (10-digit format, e.g., `7208100000`)
- Month (1, 2, 3 = months within the quarter)
- Country of Origin
- Tonnes
- C$1000 (value in thousands of Canadian dollars)
- C$/Tonne (unit price)

**HTS code matching:** Laura's template uses dotted 10-digit format (e.g., `7208.10.00.00`). B1 uses plain 10-digit format (e.g., `7208100000`). Normalize by stripping dots: `7208.10.00.00` → `7208100000`.

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
1. Fetch B1 HTML with `requests.get()`
2. Parse with BeautifulSoup — find the main data table
3. Iterate rows, check if the HS code (stripped of dots) matches any HTS code in our mapping
4. For matching rows, accumulate tonnes and C$1000 by product category and country
5. Sum across all months (1, 2, 3) in the quarter to get quarterly totals

## Excel Output Format

### Sheet: "non-FTA {quarter}" and "FTA {quarter}"

Matches Laura's template exactly:

```
Row 1: "Non-FTA Quota - Quarter 4: March 26, 2026 to June 25, 2026"
Row 2: [blank] | Country | Max Quota (KG) | Max Share (%) | Mar 30 2026 | Apr 6 2026 | ... | OVER

Row 3:  Hot-Rolled Sheet | China          | 2,370,500 | 41% | 2.95%  | ... |
Row 4:  Hot-Rolled Sheet | India          | 2,370,500 | 41% | 0.11%  | ... |
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
- **Last column**: "OVER" — `"YES"` if the country's latest share value >= Max Share %; blank otherwise. TOTAL rows have no OVER flag. The comparison is `share >= max_share` (inclusive of equality, since hitting the cap exactly means the country is at its limit).

Each new script run appends a new date column. Existing columns are never modified. The OVER column is recalculated each run based on the latest values.

### Sheet: "B1 Imports"

Aggregated import volumes by product category and country:

```
Row 1: "B1 Import Data - Quarter 4: March 26, 2026 to June 25, 2026"
Row 2: Product Category | Country | Total Tonnes | Total C$1000 | Avg C$/Tonne

Row 3: Hot-Rolled Sheet | China    | 15,234 | 12,456 | 817.89
Row 4: Hot-Rolled Sheet | Japan    | 8,901  | 9,234  | 1,037.41
...
```

This sheet is completely rewritten each run with the latest B1 data (B1 data is cumulative for the quarter, so the latest pull always contains all prior months).

### Sheet: "HTS code covered"

Static reference sheet copied from Laura's template — lists HTS codes per product category. Created once on first run, never modified.

## Quarterly Transition Logic

**Quarter date ranges (for both FTA and NFTA):**
- Q1: June 27 to September 25
- Q2: September 26 to December 25
- Q3: December 26 to March 25
- Q4: March 26 to June 26

Note: FTA's first-ever period was "P1" (August 1 - September 25, 2025), but all subsequent quarters use standard "Q" numbering for both FTA and NFTA.

**Detection algorithm:**
```python
def get_current_quarter(today):
    month, day = today.month, today.day
    if (month == 3 and day >= 26) or (month in [4, 5]) or (month == 6 and day <= 26):
        return "Q4"
    elif (month == 6 and day >= 27) or (month in [7, 8]) or (month == 9 and day <= 25):
        return "Q1"
    elif (month == 9 and day >= 26) or (month in [10, 11]) or (month == 12 and day <= 25):
        return "Q2"
    else:  # Dec 26+ or Jan/Feb/Mar 1-25
        return "Q3"
```

**URL construction:**
- FTA: `TRQ_FTA-{quarter}.csv` (e.g., `TRQ_FTA-Q4.csv`)
- NFTA: `TRQ_NFTA-{quarter}.csv` (e.g., `TRQ_NFTA-Q4.csv`)

**Verified as of 2026-03-30:** Q4 data is already available 5 days after Q3 ended (March 25). The government appears to publish new quarter data promptly.

**Fallback:** If the current quarter's CSV returns HTTP 404 (data not posted yet), try the previous quarter. Print: "Q4 data not yet available, using Q3."

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
| CSV format unrecognized (line 1 not "ExecutionTime" or "Textbox") | Exit with descriptive error. |
| `"Max Utilized"` string in utilization fields | Treat as 100% utilization. |
| `"0.00%"` or other string percentages | Strip `%` suffix and parse as float. |
| Product has zero utilization (no countries) | Show product name + TOTAL row with 0.00%. |
| Duplicate column (same date already exists) | Skip adding a new column, print "already up to date". |
| New country appears mid-quarter | Append new row at the end of the product group, backfill previous weeks as blank. |
| Country present in prior week but absent in current data | Keep existing row, leave current week's cell blank (country may have had quota reversed). |

## Data Validation

After parsing, perform basic sanity checks:
- All share percentages should be between 0% and 100% (warn if violated but don't fail)
- Sum of country shares for a product should approximately equal the total utilization % from Part A (warn if >1% discrepancy)
- Max quota and max share values should be positive numbers

## Dependencies

```
requests>=2.31.0
beautifulsoup4>=4.12.0
openpyxl>=3.1.0
```

No pandas needed — openpyxl handles Excel I/O directly, and the CSV parsing is custom (non-standard format).
