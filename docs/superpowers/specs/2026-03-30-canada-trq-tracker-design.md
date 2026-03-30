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
  1. detect_current_quarter()      -> "Q4", date range, URLs
  2. download_trq_csv("FTA")       -> raw CSV text
  3. download_trq_csv("NFTA")      -> raw CSV text
  4. parse_trq_csv(csv_text)       -> {product: {country: share%, ...}, max_quota, max_share}
  5. should_fetch_b1(trq_quarter)  -> bool (is matching B1 calendar quarter available?)
  6. scrape_b1_imports()           -> {product: {country: {tonnes, value}}} or None
  7. load_or_create_workbook()     -> openpyxl Workbook
  8. update_trq_sheet(wb, ...)     -> adds date column with formatting
  9. update_b1_sheet(wb, ...)      -> rewrites B1 sheet (if B1 data available)
  10. wb.save()
```

### File structure

```
Canada-Quota/
  canada_trq_tracker.py       # Main script
  requirements.txt             # requests, beautifulsoup4, lxml, openpyxl
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
4. Parse Part B sections: iterate lines after the blank line following Part A. Each section has:
   - A **Textbox metadata header line** (starts with `Textbox` or `CONTROL_ITEMS_Level_1`). Every Part A item (1-23) gets a header, **including zero-utilization products**.
   - Optionally followed by **data rows** (only if the product has non-zero utilization):
     - **New format**: data rows start with `"Country,"` label prefix (offset=3)
     - **Old format**: data rows start directly with the country name (offset=0). Distinguish from headers by checking that the first field is NOT a known prefix (`"Textbox"`, `"CONTROL_ITEMS_Level_1"`).
   - Followed by a blank line before the next section.
   - **Simpler detection**: after a header line, all non-blank lines until the next blank line are data rows.
5. **Matching strategy: positional with zero-utilization awareness.**
   - Walk Part B sections sequentially. Each Textbox header = one Part A item (in order 1-23).
   - If the header is followed immediately by a blank line (no `Country,...` rows), the product has zero utilization → skip.
   - If `Country,...` rows follow, consume them as the country data for the current Part A item.
   - **Confirmation: compare `total_kgm`** (from the last two fields of any country row) against Part A's `total_util_kgm` — log a warning if they differ by more than 1 KGM.
6. Extract `share_pct` for each country in each of the 8 tracked products. For zero-utilization tracked products, record 0.0% with no countries.

**Special values to handle:**
- `"Max Utilized"` string (instead of numeric values) — treat as 100% utilization
- `"0.00%"` as string — strip `%` and parse as float 0.0
- Comma-separated thousands in quoted numbers: `"1,042,387"` — strip commas before parsing

### Source 2: B1 Import Data

**URL:** `https://www.eics-scei.gc.ca/report-rapport/b1.htm`

**Format:** Single large HTML page (2.8 MB, ~7,600 rows in the main table) generated by Microsoft Report 8.0 (SSRS via Power BI). Covers all steel HS codes for the current **calendar** quarter with monthly breakdown.

The page META tags expose the reporting period:
```
META pFromDate: 1/1/2026 12:00:00 AM
META pToDate:   3/21/2026 12:00:00 AM
```

The URL is stable — `b1.htm` always shows the current calendar quarter's data. When the calendar quarter changes (e.g., Apr 1), the page automatically shows the new period. Historical calendar quarters are not available from this URL.

#### B1 vs TRQ: Different Data, Different Time Periods

**Critical context:** TRQ utilization data and B1 import data measure fundamentally different things:

| | TRQ Utilization | B1 Imports |
|---|---|---|
| **Measures** | Permit quantities issued (administrative) | Actual customs entries (physical) |
| **Triggered by** | Import permit issuance in EICS/NEICS | Goods clearing customs |
| **Quarter system** | Offset quarters (Mar 26–Jun 26, etc.) | Calendar quarters (Jan 1–Mar 31, etc.) |
| **Internal DB field** | `ID_PERMIT_ITEM_FACT_Permit_Quantity` | Customs records |
| **Update lag** | Near-real-time (reports generated on demand) | Several days behind |

Because of this, the two datasets will never match perfectly — permits can be issued but goods may not arrive, actual imports may differ from permitted quantities, and the quarterly boundaries are offset by ~5 days.

#### Approximate Quarter Alignment Strategy

B1 data includes a `Month` field (1, 2, 3) for each row. Use this to approximate alignment with TRQ offset quarters:

**Mapping table (TRQ quarter → B1 months to include):**

| TRQ Quarter | TRQ Dates | B1 Calendar Quarter | B1 Months Used | Overlap Gap |
|-------------|-----------|-------------------|----------------|-------------|
| Q1 | Jun 27 – Sep 25 | Jul–Sep (Q3 cal) | Month 1, 2, 3 | ~4 days at start, ~5 days at end |
| Q2 | Sep 26 – Dec 25 | Oct–Dec (Q4 cal) | Month 1, 2, 3 | ~5 days at start, ~6 days at end |
| Q3 | Dec 26 – Mar 25 | Jan–Mar (Q1 cal) | Month 1, 2, 3 | ~6 days at start, ~6 days at end |
| Q4 | Mar 26 – Jun 26 | Apr–Jun (Q2 cal) | Month 1, 2, 3 | ~5 days at start, ~4 days at end |

**Logic:** For a given TRQ quarter, fetch the B1 page during the **next** calendar quarter that overlaps most with it. Sum all 3 months. This gives ~85-90% time overlap, with ~5-6 day misalignment at each boundary.

**Implementation:**
```python
TRQ_TO_B1_CALENDAR_QUARTER = {
    "Q1": (7, 9),   # Jul-Sep → B1 shows as months 1,2,3 during Q3 calendar
    "Q2": (10, 12),  # Oct-Dec → B1 shows as months 1,2,3 during Q4 calendar
    "Q3": (1, 3),    # Jan-Mar → B1 shows as months 1,2,3 during Q1 calendar
    "Q4": (4, 6),    # Apr-Jun → B1 shows as months 1,2,3 during Q2 calendar
}
```

Since B1 always shows the *current* calendar quarter, the script should only fetch B1 when the calendar quarter overlaps with the active TRQ quarter. If there's no overlap yet (e.g., TRQ Q4 just started but B1 still shows Jan–Mar), skip B1 update and log: "B1 data for matching calendar quarter not yet available."

#### B1 HTML Row Format

The HTML table has **variable-length rows**. The parser must maintain state to track the current HS code:

| Row Type | Cell Count | Fields | Meaning |
|----------|-----------|--------|---------|
| **HS-code row** | 8 | `['', HS_code, commodity_desc, month, country, tonnes, C$1000, C$/Tonne]` | First appearance of an HS code |
| **Month-continuation** | 6 | `['', month, country, tonnes, C$1000, C$/Tonne]` | Same HS code, different month |
| **Country-continuation** | 5 | `['', country, tonnes, C$1000, C$/Tonne]` | Same HS code + same month, different country |
| **Summary row** | 7 | `['', 'Summary of HS 10 Code{code}', '', '', tonnes, C$1000, C$/Tonne]` | Subtotal for the HS code |

**Stateful parsing algorithm:**
```python
current_hs = None
current_month = None

for row in table_rows:
    cells = [td.get_text(strip=True) for td in row.find_all(['td', 'th'])]
    n = len(cells)

    if n == 8 and cells[1] and cells[1].isdigit() and len(cells[1]) == 10:
        # New HS code row
        current_hs = cells[1]
        current_month = cells[3]
        country, tonnes, value = cells[4], cells[5], cells[6]
        # → process(current_hs, current_month, country, tonnes, value)

    elif n == 6 and cells[1].isdigit() and len(cells[1]) <= 2:
        # Month-continuation (same HS code)
        current_month = cells[1]
        country, tonnes, value = cells[2], cells[3], cells[4]
        # → process(current_hs, current_month, country, tonnes, value)

    elif n == 5 and current_hs:
        # Country-continuation (same HS code + month)
        country, tonnes, value = cells[1], cells[2], cells[3]
        # → process(current_hs, current_month, country, tonnes, value)

    elif n == 7 and 'Summary' in cells[1]:
        # Summary row — skip (we calculate our own totals)
        current_hs = None
        continue
```

**Number parsing:** Tonnes and C$1000 values use comma separators (e.g., `"1,330.38"`). Strip commas before converting to float. Some values are `"0.00"` — parse as 0.0.

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

#### Country Name Normalization

TRQ and B1 data both come from the Canadian government but may use slightly different country name conventions. Maintain a normalization map for known discrepancies:

```python
COUNTRY_NAME_MAP = {
    # B1 name → TRQ name (add entries as discrepancies are discovered)
    # Both sources appear to use the same names (same government system),
    # but log warnings for any country in B1 that doesn't match TRQ exactly.
}
```

On first run, log all unique country names from both sources and flag any mismatches for manual review. After initial verification, this map can be hardcoded.

## Excel Output Format

### Value Conversion: CSV String → Excel Cell

All percentage values from CSV must be converted before writing to Excel:

```
CSV string "2.95%" → strip "%" → float 2.95 → divide by 100 → 0.0295 → write to cell
```

The cell's `number_format` then renders `0.0295` as `3.0%`. This matches Laura's template, where values like `0.023` are stored as raw decimals and displayed via format `0.0%` as `2.3%`.

For max_share values: `"41%"` → strip `%` → 41 → divide by 100 → `0.41`.

For max_quota values: `"2,370,500"` → strip commas → `2370500` → write as integer.

### Sheet: "non-FTA {quarter}" and "FTA {quarter}"

Matches Laura's template exactly:

```
Row 1: "Non-FTA Quota - Quarter 4: March 26, 2026 to June 26, 2026"
Row 2: [blank] | Country | Max Quota (KG) | Max Share (%) | March 30 2026 | April 6 2026 | ... | OVER

Row 3:  Hot-Rolled Sheet | China          | 2,370,500 | 41.0% | 3.0%   | ... |
Row 4:  Hot-Rolled Sheet | India          | 2,370,500 | 41.0% | 0.1%   | ... |
Row 5:  Hot-Rolled Sheet | Taiwan         | 2,370,500 | 41.0% | 40.9%  | ... | YES
Row 6:  Hot-Rolled Sheet | TOTAL          |           |       | =SUM() | ... |
Row 7:  Steel Plate      | China          | 5,201,200 | 36.0% | 36.0%  | ... | YES
...
```

#### Column specifications:

| Column | Content | Number Format | Notes |
|--------|---------|---------------|-------|
| **A** | Product category name | General | Repeated for each country row within that product |
| **B** | Country name | General | Last row of each product group = `"TOTAL"` |
| **C** | Max Quota (KG) | `#,##0` | Static. Same for all countries within a product. Blank on TOTAL rows. |
| **D** | Max Share (%) | `0.0%` | Static. The single-country cap. Blank on TOTAL rows. |
| **E+** | Weekly date columns | `0.0%` | Country's `Share of total utilization (%)` as decimal (e.g., 0.0295 for 2.95%) |
| **Last** | OVER | General | `"YES"` if country's latest share ≥ Max Share (inclusive). Blank on TOTAL rows. |

#### Date column header format:

Use **full month name** to match Laura's template: `"March 30 2026"`, `"April 6 2026"`, etc.

```python
from datetime import date
header = date.today().strftime("%B %-d %Y")  # "March 30 2026" (no leading zero on day)
# On Windows: date.today().strftime("%B %#d %Y")
```

#### TOTAL row formula:

TOTAL rows use Excel **SUM formulas** (not hardcoded values), matching Laura's template:

```python
# Example: if Hot-Rolled Sheet has countries in rows 3-5, and current date column is E:
# Row 6, Column E = "=SUM(E3:E5)"
from openpyxl.utils import get_column_letter
col_letter = get_column_letter(date_col_index)
total_cell.value = f"=SUM({col_letter}{first_country_row}:{col_letter}{last_country_row})"
total_cell.number_format = '0.0%'
```

This provides cross-validation: if the SUM of country shares doesn't approximately match Part A's total utilization %, it indicates a parsing issue.

#### OVER column logic:

For each country row (not TOTAL rows):
- Compare the **latest** date column's value against Max Share
- If `share >= max_share`: write `"YES"`
- Otherwise: leave blank
- The OVER column is **recalculated** each run (delete and rewrite the entire column)

#### Excel visual formatting (matching Laura's template):

Laura's template uses yellow highlighting (`FFFFFF00`) on specific cells:

```python
from openpyxl.styles import PatternFill
yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
```

Apply yellow fill to:
1. **"OVER" header cell** (Row 2, last column)
2. **All TOTAL rows' date column cells** (the SUM formula cells)

Apply **bold** to:
1. **All header cells** (Row 2: Country, Max Quota, Max Share, date columns, OVER)

#### Appending new date columns:

Each new script run appends a new date column. Existing columns are **never modified**. The process:
1. Find the OVER column (scan Row 2 for cell with value `"OVER"`)
2. If today's date header already exists in Row 2 → skip ("already up to date")
3. **Overwrite** the OVER column with the new date data:
   - Row 2: write date header (e.g., `"March 30 2026"`) with bold
   - Data rows: write country share values with `number_format = '0.0%'`
   - TOTAL rows: write SUM formulas with `number_format = '0.0%'` and yellow fill
4. Write OVER in the **next** column (one position to the right):
   - Row 2: `"OVER"` header with bold + yellow fill
   - Data rows: `"YES"` or blank based on latest share vs max_share

This approach avoids complex column insertion — the OVER column simply shifts right by one each week.

#### First run behavior:

When the output Excel file doesn't exist yet:
1. Create new workbook, remove default sheet
2. Create `"non-FTA {quarter}"` and `"FTA {quarter}"` sheets with Row 1 title and Row 2 headers
3. Populate all country rows from TRQ data, with first date column in E and OVER in F
4. Create `"B1 Imports"` sheet with disclaimer notes and header row (data filled if B1 available)
5. Create `"HTS code covered"` sheet, copying structure from Laura's template

#### New country handling:

- **New country appears**: Append new row at the end of the product group (before TOTAL), backfill previous week columns as blank
- **Country absent in current data**: Keep existing row, leave current week's cell blank
- **TOTAL SUM range**: Update the SUM formula range to include any newly added country rows

### Sheet: "B1 Imports"

Aggregated actual import volumes by product category and country, from customs data:

```
Row 1: "B1 Import Data - Calendar Quarter: January 1 to March 31, 2026"
Row 2: (blank)
Row 3: "NOTE: B1 data reflects actual customs entries using calendar quarters (Jan-Mar, Apr-Jun, etc.)."
Row 4: "TRQ utilization data uses offset quarters (Mar 26-Jun 26, etc.) and tracks permit issuance, not customs entries."
Row 5: "The two datasets have ~5-6 day boundary misalignment and measure different aspects of the import process."
Row 6: (blank)
Row 7: Product Category | Country | Total Tonnes | Total C$1000 | Avg C$/Tonne

Row 8:  Hot-Rolled Sheet | China    | 15,234 | 12,456 | 817.89
Row 9:  Hot-Rolled Sheet | Japan    | 8,901  | 9,234  | 1,037.41
...
```

| Column | Number Format |
|--------|---------------|
| Product Category | General |
| Country | General |
| Total Tonnes | `#,##0.00` |
| Total C$1000 | `#,##0.00` |
| Avg C$/Tonne | `#,##0.00` |

This sheet is completely rewritten each run with the latest B1 data (B1 data is cumulative for the calendar quarter, so the latest pull always contains all prior months).

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
| B1 calendar quarter doesn't overlap with TRQ quarter | Log info, skip B1 update: "B1 data for matching calendar quarter not yet available." |
| CSV format unrecognized (line 1 not "ExecutionTime" or "Textbox") | Exit with descriptive error. |
| `"Max Utilized"` string in utilization fields | Treat as 100% utilization (store as `1.0` in cell, displays as `100.0%`). |
| `"0.00%"` or other string percentages | Strip `%` suffix, parse as float, divide by 100 (e.g., `"0.00%"` → `0.0`). |
| Comma-separated thousands in quoted values | Strip commas before numeric conversion (e.g., `"2,370,500"` → `2370500`). |
| Product has zero utilization (no countries) | Show product name + TOTAL row. TOTAL formula = `=SUM(Erow:Erow)` which evaluates to 0. |
| Duplicate column (same date already exists) | Skip adding a new column, print "already up to date". |
| New country appears mid-quarter | Append new row at the end of the product group (before TOTAL), backfill previous weeks as blank. Update TOTAL SUM formula range. |
| Country present in prior week but absent in current data | Keep existing row, leave current week's cell blank (permit may have been cancelled). |
| B1 country name doesn't match any TRQ country | Log warning with both names. Use B1 name as-is; add to `COUNTRY_NAME_MAP` for future normalization. |
| B1 HTML structure changed (unexpected cell count) | Log warning with row details, skip that row, continue parsing. |
| Part A total_util_kgm ≠ Part B total_kgm (>1 KGM diff) | Log warning but continue — does not block execution. |

## Data Validation

After parsing (values are in decimal form, e.g., 0.41 = 41%), perform basic sanity checks:
- All share values should be between 0.0 and 1.0 (warn if violated but don't fail)
- Sum of country shares for a product should approximately equal the total utilization from Part A (warn if discrepancy > 0.01, i.e., >1 percentage point)
- Max quota values should be positive integers
- Max share values should be between 0.0 and 1.0

## Dependencies

```
requests>=2.31.0
beautifulsoup4>=4.12.0
lxml>=5.0.0
openpyxl>=3.1.0
```

- **lxml**: Required as the HTML parser for BeautifulSoup. The B1 page is 2.8 MB with 7,600+ rows — `lxml` is ~10x faster than the default `html.parser` for pages this large. Use: `BeautifulSoup(html, 'lxml')`.
- **No pandas needed** — openpyxl handles Excel I/O directly, and the CSV parsing is custom (non-standard format, not suitable for `pandas.read_csv`).
