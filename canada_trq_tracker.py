"""
Canada TRQ Weekly Tracker

Downloads Canadian steel Tariff Rate Quota (TRQ) utilization data and B1 import
data, parsing them into an Excel workbook that matches Laura's template format.
Designed to run weekly via GitHub Actions.
"""

import csv
import io
import logging
import os
import platform
import re
import sys
import time
from datetime import date, datetime

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

BASE_URL = "https://www.eics-scei.gc.ca/report-rapport"
B1_URL = f"{BASE_URL}/b1.htm"
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "canada_trq_tracker.xlsx")

# The 8 product categories Laura tracks (must match TRQ product names exactly)
TRACKED_PRODUCTS = [
    "Hot-Rolled Sheet",
    "Steel Plate",
    "Cold-Rolled Sheet",
    "Coated Steel Sheet",
    "Rebar",
    "Hot-Rolled Bar",
    "Wire Rod",
    "Structural Steel",
]

# HTS code mapping (dotted format from Laura's template)
HTS_CODE_MAP = {
    "Hot-Rolled Sheet": [
        "7208.10.00.00", "7208.25.00.00", "7208.26.00.00", "7208.27.00.00",
        "7208.36.00.00", "7208.37.00.00", "7208.38.00.00", "7208.39.00.00",
    ],
    "Steel Plate": [
        "7208.40.00.00", "7208.51.00.00", "7208.52.00.00",
    ],
    "Cold-Rolled Sheet": [
        "7209.15.00.00", "7209.16.00.00", "7209.17.00.00", "7209.18.00.00",
        "7209.25.00.00", "7209.26.00.00", "7209.27.00.00", "7209.28.00.00",
    ],
    "Coated Steel Sheet": [
        "7210.41.00.00", "7210.49.00.00", "7210.61.00.00",
        "7210.69.00.00", "7210.70.00.00", "7210.90.00.00",
    ],
    "Rebar": ["7214.20.00.00"],
    "Hot-Rolled Bar": [
        "7214.30.00.00", "7214.91.00.00", "7214.99.00.00",
    ],
    "Wire Rod": [
        "7213.10.00.00", "7213.20.00.00", "7213.91.00.00", "7213.99.00.00",
    ],
    "Structural Steel": [
        "7216.10.00.00", "7216.21.00.00", "7216.22.00.00", "7216.31.00.00",
        "7216.32.00.00", "7216.33.00.00", "7216.40.00.00", "7216.50.00.00",
        "7216.69.00.00",
    ],
}

# Build reverse lookup: 8-digit HS prefix -> product category
# Laura's codes use "00" as the last 2 statistical digits (e.g., 7214.20.00.00 → 7214200000)
# but B1 uses specific statistical suffixes (e.g., 7214200011, 7214200012).
# Matching on first 8 digits captures all sub-codes correctly.
_HTS_REVERSE: dict[str, str] = {}
for _prod, _codes in HTS_CODE_MAP.items():
    for _code in _codes:
        _prefix_8 = _code.replace(".", "")[:8]
        _HTS_REVERSE[_prefix_8] = _prod

# Quarter date ranges  (month, day) boundaries
QUARTER_RANGES = {
    "Q1": ("June 27", "September 25"),
    "Q2": ("September 26", "December 25"),
    "Q3": ("December 26", "March 25"),
    "Q4": ("March 26", "June 26"),
}

# TRQ offset quarter -> overlapping B1 calendar quarter months
TRQ_TO_B1_CALENDAR_MONTHS: dict[str, tuple[int, int]] = {
    "Q1": (7, 9),
    "Q2": (10, 12),
    "Q3": (1, 3),
    "Q4": (4, 6),
}

# Country name normalization (B1 name -> TRQ name)
COUNTRY_NAME_MAP: dict[str, str] = {
    # Both sources use the same government system; add entries if mismatches found.
}

# Excel formatting constants
YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
BOLD_FONT = Font(bold=True)

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def parse_number(s: str) -> float:
    """Strip commas and parse a numeric string. Returns 0.0 for empty/blank."""
    s = s.strip().strip('"')
    if not s:
        return 0.0
    return float(s.replace(",", ""))


def parse_percent(s: str) -> float:
    """Parse a percentage string to a decimal (e.g. '2.95%' -> 0.0295).
    Handles 'Max Utilized' as 1.0 (100%)."""
    s = s.strip().strip('"')
    if not s:
        return 0.0
    if s == "Max Utilized":
        return 1.0
    if s.endswith("%"):
        return float(s[:-1].replace(",", "")) / 100.0
    # Already a decimal number
    return float(s.replace(",", ""))


def format_date_header(d: date) -> str:
    """Format a date as 'March 30 2026' (full month, no leading zero on day)."""
    if platform.system() == "Windows":
        return d.strftime("%B %#d %Y")
    return d.strftime("%B %-d %Y")


def get_current_quarter(today: date) -> str:
    """Determine the TRQ quarter for a given date."""
    month, day = today.month, today.day
    if (month == 3 and day >= 26) or month in (4, 5) or (month == 6 and day <= 26):
        return "Q4"
    if (month == 6 and day >= 27) or month in (7, 8) or (month == 9 and day <= 25):
        return "Q1"
    if (month == 9 and day >= 26) or month in (10, 11) or (month == 12 and day <= 25):
        return "Q2"
    # Dec 26+ or Jan/Feb/Mar 1-25
    return "Q3"


def get_quarter_date_range(quarter: str, ref_date: date) -> tuple[str, str]:
    """Return human-readable start/end dates for the given quarter relative to ref_date."""
    year = ref_date.year
    if quarter == "Q1":
        return f"June 27, {year}", f"September 25, {year}"
    if quarter == "Q2":
        return f"September 26, {year}", f"December 25, {year}"
    if quarter == "Q3":
        # Q3 spans the year boundary
        start_year = year - 1 if ref_date.month <= 3 else year
        end_year = start_year + 1 if start_year == year else year
        return f"December 26, {start_year}", f"March 25, {end_year}"
    # Q4
    return f"March 26, {year}", f"June 26, {year}"


def should_fetch_b1(trq_quarter: str, today: date) -> bool:
    """Check if B1 page currently shows data for the calendar quarter matching this TRQ quarter."""
    cal_start, cal_end = TRQ_TO_B1_CALENDAR_MONTHS[trq_quarter]
    return cal_start <= today.month <= cal_end


# ---------------------------------------------------------------------------
# TRQ CSV Download & Parsing
# ---------------------------------------------------------------------------

def download_csv(trq_type: str, quarter: str) -> str | None:
    """Download TRQ CSV. trq_type is 'FTA' or 'NFTA'. Returns text or None."""
    url = f"{BASE_URL}/TRQ_{trq_type}-{quarter}.csv"
    for attempt in range(2):
        try:
            log.info("Downloading %s (attempt %d)...", url, attempt + 1)
            resp = requests.get(url, timeout=30)
            if resp.status_code == 404:
                log.warning("HTTP 404 for %s", url)
                return None
            resp.raise_for_status()
            resp.encoding = "utf-8"
            return resp.text
        except requests.RequestException as exc:
            log.warning("Download failed: %s", exc)
            if attempt == 0:
                time.sleep(5)
    return None


def parse_trq_csv(csv_text: str) -> dict:
    """Parse TRQ CSV (either format) and return structured data.

    Returns:
        {
            "products": {
                "Hot-Rolled Sheet": {
                    "item_number": 3,
                    "max_quota": 2370500,
                    "max_share": 0.41,
                    "total_util": 0.4397,
                    "countries": {"China": 0.0295, "India": 0.0011, ...}
                },
                ...
            }
        }
    """
    lines = csv_text.splitlines()
    if not lines:
        raise ValueError("Empty CSV")

    # Detect format
    if lines[0].startswith("ExecutionTime"):
        fmt = "old"
        part_a_offset = 0
        part_b_offset = 0
    elif lines[0].startswith("Textbox"):
        fmt = "new"
        part_a_offset = 6
        part_b_offset = 3
    else:
        raise ValueError(f"Unrecognized CSV format. Line 1: {lines[0][:80]}")

    log.info("CSV format detected: %s", fmt)

    # --- Part A: lines 5-27 (1-indexed: lines[4:27]) ---
    part_a_items: list[dict] = []
    for i in range(4, 27):
        if i >= len(lines) or not lines[i].strip():
            continue
        fields = list(csv.reader(io.StringIO(lines[i])))[0]
        o = part_a_offset
        item_number = int(fields[o])
        product_name = fields[o + 1].strip()
        max_quota = int(parse_number(fields[o + 2]))
        max_share = parse_percent(fields[o + 3])
        util_kgm_raw = fields[o + 4].strip().strip('"')
        util_pct_raw = fields[o + 5].strip().strip('"')
        total_util_kgm = 0.0 if util_kgm_raw == "Max Utilized" else parse_number(util_kgm_raw)
        total_util_pct = 1.0 if util_pct_raw == "Max Utilized" else parse_percent(util_pct_raw)

        part_a_items.append({
            "item_number": item_number,
            "product_name": product_name,
            "max_quota": max_quota,
            "max_share": max_share,
            "total_util_kgm": total_util_kgm,
            "total_util_pct": total_util_pct,
        })

    # --- Part B: lines 29+ (after blank line 28) ---
    part_b_start = 28  # 0-indexed
    # Walk Part B sections, matching sequentially to Part A items
    part_b_sections: list[list[dict]] = []
    current_section: list[dict] | None = None
    in_header = False

    for i in range(part_b_start, len(lines)):
        line = lines[i].strip()

        if not line:
            # Blank line: end of current section
            if current_section is not None:
                part_b_sections.append(current_section)
                current_section = None
            in_header = False
            continue

        # Check if this is a metadata header line
        if line.startswith("Textbox") or line.startswith("CONTROL_ITEMS_Level_1"):
            if current_section is not None:
                # Previous section ended (no blank line separator — shouldn't happen, but be safe)
                part_b_sections.append(current_section)
            current_section = []
            in_header = True
            continue

        # Data row (not blank, not header)
        if current_section is not None:
            fields = list(csv.reader(io.StringIO(line)))[0]
            o = part_b_offset
            if len(fields) > o + 4:
                country = fields[o].strip()
                util_kgm = parse_number(fields[o + 1])
                share_pct = parse_percent(fields[o + 2])
                total_kgm = parse_number(fields[o + 3])
                total_pct = parse_percent(fields[o + 4])
                current_section.append({
                    "country": country,
                    "util_kgm": util_kgm,
                    "share_pct": share_pct,
                    "total_kgm": total_kgm,
                    "total_pct": total_pct,
                })

    # Don't forget the last section if file doesn't end with blank line
    if current_section is not None:
        part_b_sections.append(current_section)

    # --- Match Part B sections to Part A items ---
    result: dict[str, dict] = {}
    b_idx = 0

    for a_item in part_a_items:
        product_name = a_item["product_name"]
        countries: dict[str, float] = {}

        if a_item["total_util_kgm"] > 0 or a_item["total_util_pct"] > 0:
            # This product should have a Part B section
            if b_idx < len(part_b_sections):
                section = part_b_sections[b_idx]
                b_idx += 1

                for row in section:
                    countries[row["country"]] = row["share_pct"]

                # Validation: compare total_kgm
                if section:
                    b_total_kgm = section[0]["total_kgm"]
                    if abs(b_total_kgm - a_item["total_util_kgm"]) > 1:
                        log.warning(
                            "Part A/B mismatch for %s: Part A util=%.0f, Part B total=%.0f",
                            product_name, a_item["total_util_kgm"], b_total_kgm,
                        )
            else:
                log.warning("No Part B section found for %s (expected non-zero utilization)", product_name)
        else:
            # Zero utilization — still has a header in Part B but no data rows.
            # The header+blank was counted as an empty section.
            if b_idx < len(part_b_sections) and len(part_b_sections[b_idx]) == 0:
                b_idx += 1  # consume the empty section

        # Only store tracked products
        if product_name in TRACKED_PRODUCTS:
            result[product_name] = {
                "item_number": a_item["item_number"],
                "max_quota": a_item["max_quota"],
                "max_share": a_item["max_share"],
                "total_util_pct": a_item["total_util_pct"],
                "countries": countries,
            }

    # Validate: all tracked products should be present
    for prod in TRACKED_PRODUCTS:
        if prod not in result:
            log.warning("Tracked product '%s' not found in CSV data", prod)
            result[prod] = {
                "item_number": 0,
                "max_quota": 0,
                "max_share": 0.0,
                "total_util_pct": 0.0,
                "countries": {},
            }

    # Data validation
    for prod_name, prod_data in result.items():
        for country, share in prod_data["countries"].items():
            if not (0.0 <= share <= 1.0):
                log.warning("Share out of range for %s/%s: %.4f", prod_name, country, share)
        country_sum = sum(prod_data["countries"].values())
        if abs(country_sum - prod_data["total_util_pct"]) > 0.01:
            log.warning(
                "Country share sum (%.4f) != total utilization (%.4f) for %s",
                country_sum, prod_data["total_util_pct"], prod_name,
            )

    return result


# ---------------------------------------------------------------------------
# B1 Import Data Scraping
# ---------------------------------------------------------------------------

def scrape_b1_imports() -> dict[str, dict[str, dict[str, float]]] | None:
    """Scrape B1 HTML page and return import data for tracked products.

    Returns:
        {
            "Hot-Rolled Sheet": {
                "China": {"tonnes": 1234.5, "value": 5678.9},
                ...
            },
            ...
        }
    """
    try:
        log.info("Downloading B1 page (%s)...", B1_URL)
        resp = requests.get(B1_URL, timeout=60)
        resp.raise_for_status()
        resp.encoding = "utf-8"
    except requests.RequestException as exc:
        log.warning("B1 download failed: %s", exc)
        return None

    log.info("Parsing B1 HTML (%.1f MB)...", len(resp.text) / 1_000_000)
    soup = BeautifulSoup(resp.text, "lxml")

    # Find the main data table (the one with >100 rows)
    main_table = None
    for table in soup.find_all("table"):
        if len(table.find_all("tr")) > 100:
            main_table = table
            break

    if main_table is None:
        log.warning("Could not find main data table in B1 page")
        return None

    rows = main_table.find_all("tr")
    log.info("B1 table has %d rows", len(rows))

    # Stateful parsing
    result: dict[str, dict[str, dict[str, float]]] = {}
    current_hs: str | None = None
    current_month: str | None = None

    def process(hs_code: str, month: str, country: str, tonnes_str: str, value_str: str):
        product = _HTS_REVERSE.get(hs_code[:8])
        if product is None:
            return  # Not a tracked HS code
        country = COUNTRY_NAME_MAP.get(country, country)
        tonnes = parse_number(tonnes_str)
        value = parse_number(value_str)
        if product not in result:
            result[product] = {}
        if country not in result[product]:
            result[product][country] = {"tonnes": 0.0, "value": 0.0}
        result[product][country]["tonnes"] += tonnes
        result[product][country]["value"] += value

    for row in rows:
        cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
        n = len(cells)

        if n == 8 and len(cells[1]) == 10 and cells[1].isdigit():
            current_hs = cells[1]
            current_month = cells[3]
            process(current_hs, current_month, cells[4], cells[5], cells[6])

        elif n == 6 and cells[1].isdigit() and len(cells[1]) <= 2:
            current_month = cells[1]
            process(current_hs or "", current_month, cells[2], cells[3], cells[4])

        elif n == 5 and current_hs:
            process(current_hs, current_month or "", cells[1], cells[2], cells[3])

        elif n == 7 and "Summary" in cells[1]:
            current_hs = None
            continue

    log.info("B1 parsing complete: %d tracked products found", len(result))
    return result if result else None


# ---------------------------------------------------------------------------
# Excel Workbook Management
# ---------------------------------------------------------------------------

def _find_over_col(ws) -> int | None:
    """Find the column index (1-based) of the OVER header in row 2."""
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=2, column=col).value
        if val and str(val).strip().upper() == "OVER":
            return col
    return None


def _find_date_col(ws, date_header: str) -> int | None:
    """Find the column index of a specific date header in row 2."""
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=2, column=col).value
        if val and str(val).strip() == date_header:
            return col
    return None


def _get_product_row_ranges(ws) -> dict[str, dict]:
    """Scan the sheet to build a map of product -> {first_row, last_country_row, total_row}.

    Assumes data starts at row 3. Column A = product name, Column B = country.
    """
    ranges: dict[str, dict] = {}
    current_product = None
    first_row = None
    last_country_row = None

    for row in range(3, ws.max_row + 1):
        prod = ws.cell(row=row, column=1).value
        country = ws.cell(row=row, column=2).value

        if prod and str(prod).strip():
            prod = str(prod).strip()
            if prod != current_product:
                # Finalize previous product
                if current_product and first_row:
                    ranges[current_product]["last_country_row"] = last_country_row
                current_product = prod
                first_row = row
                last_country_row = row
                if prod not in ranges:
                    ranges[prod] = {"first_row": row, "last_country_row": row, "total_row": None}

        if country and str(country).strip() == "TOTAL" and current_product:
            ranges[current_product]["total_row"] = row
            ranges[current_product]["last_country_row"] = last_country_row
        elif country and str(country).strip() != "TOTAL" and current_product:
            last_country_row = row
            ranges[current_product]["last_country_row"] = last_country_row

    return ranges


def create_trq_sheet(wb: Workbook, sheet_name: str, quarter: str,
                     trq_type: str, data: dict, today: date):
    """Create a new TRQ sheet with initial data."""
    ws = wb.create_sheet(title=sheet_name)
    start_date, end_date = get_quarter_date_range(quarter, today)
    type_label = "Non-FTA" if trq_type == "NFTA" else "FTA"

    # Row 1: Title
    ws.cell(row=1, column=1, value=f"{type_label} Quota - Quarter {quarter[1]}: {start_date} to {end_date}")

    # Row 2: Headers
    date_header = format_date_header(today)
    headers = [None, "Country", "Max Quota (KG)", "Max Share (%)", date_header, "OVER"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        if header is not None:
            cell.font = BOLD_FONT
    # Yellow on OVER header
    ws.cell(row=2, column=6).fill = YELLOW_FILL

    # Data rows
    row_num = 3
    date_col = 5  # Column E
    over_col = 6  # Column F

    for product in TRACKED_PRODUCTS:
        prod_data = data.get(product, {
            "max_quota": 0, "max_share": 0.0, "countries": {},
        })
        countries = prod_data.get("countries", {})
        max_quota = prod_data.get("max_quota", 0)
        max_share = prod_data.get("max_share", 0.0)

        # Sort countries alphabetically for consistent ordering
        sorted_countries = sorted(countries.keys())

        if not sorted_countries:
            # No countries — just write TOTAL row with 0
            ws.cell(row=row_num, column=1, value=product)
            ws.cell(row=row_num, column=2, value="TOTAL")
            col_letter = get_column_letter(date_col)
            total_cell = ws.cell(row=row_num, column=date_col)
            total_cell.value = 0.0
            total_cell.number_format = "0.0%"
            total_cell.fill = YELLOW_FILL
            row_num += 1
            continue

        first_country_row = row_num
        for country in sorted_countries:
            share = countries[country]
            ws.cell(row=row_num, column=1, value=product)
            ws.cell(row=row_num, column=2, value=country)

            quota_cell = ws.cell(row=row_num, column=3, value=max_quota)
            quota_cell.number_format = "#,##0"

            share_cap_cell = ws.cell(row=row_num, column=4, value=max_share)
            share_cap_cell.number_format = "0.0%"

            util_cell = ws.cell(row=row_num, column=date_col, value=share)
            util_cell.number_format = "0.0%"

            # OVER flag
            if share >= max_share and max_share > 0:
                ws.cell(row=row_num, column=over_col, value="YES")

            row_num += 1

        last_country_row = row_num - 1

        # TOTAL row
        ws.cell(row=row_num, column=1, value=product)
        ws.cell(row=row_num, column=2, value="TOTAL")
        col_letter = get_column_letter(date_col)
        total_cell = ws.cell(row=row_num, column=date_col)
        total_cell.value = f"=SUM({col_letter}{first_country_row}:{col_letter}{last_country_row})"
        total_cell.number_format = "0.0%"
        total_cell.fill = YELLOW_FILL
        row_num += 1

    log.info("Created sheet '%s' with %d rows", sheet_name, row_num - 1)


def update_trq_sheet(ws, data: dict, today: date):
    """Add a new date column to an existing TRQ sheet."""
    date_header = format_date_header(today)

    # Check if this date already exists
    if _find_date_col(ws, date_header):
        log.info("Sheet '%s' already has column '%s' — skipping", ws.title, date_header)
        return

    # Find OVER column
    over_col = _find_over_col(ws)
    if over_col is None:
        log.warning("No OVER column found in '%s' — cannot update", ws.title)
        return

    # New date column replaces OVER position; OVER moves right by 1
    new_date_col = over_col
    new_over_col = over_col + 1

    # Write date header
    header_cell = ws.cell(row=2, column=new_date_col, value=date_header)
    header_cell.font = BOLD_FONT

    # Get product row ranges
    ranges = _get_product_row_ranges(ws)

    # Write data for each row
    for row in range(3, ws.max_row + 1):
        prod = ws.cell(row=row, column=1).value
        country = ws.cell(row=row, column=2).value

        if not prod or not country:
            continue

        prod = str(prod).strip()
        country = str(country).strip()

        if country == "TOTAL":
            # Write SUM formula
            if prod in ranges:
                r = ranges[prod]
                col_letter = get_column_letter(new_date_col)
                first = r["first_row"]
                last = r["last_country_row"]
                total_cell = ws.cell(row=row, column=new_date_col)
                total_cell.value = f"=SUM({col_letter}{first}:{col_letter}{last})"
                total_cell.number_format = "0.0%"
                total_cell.fill = YELLOW_FILL
        else:
            # Country row — find share value
            prod_data = data.get(prod, {})
            countries = prod_data.get("countries", {})
            share = countries.get(country)

            if share is not None:
                cell = ws.cell(row=row, column=new_date_col)
                cell.value = share
                cell.number_format = "0.0%"
            # If country not in current data, leave blank

    # Check for new countries not yet in the sheet
    for prod in TRACKED_PRODUCTS:
        prod_data = data.get(prod, {})
        countries = prod_data.get("countries", {})

        if prod not in ranges:
            continue

        existing_countries = set()
        r = ranges[prod]
        for row in range(r["first_row"], (r["total_row"] or r["last_country_row"]) + 1):
            c = ws.cell(row=row, column=2).value
            if c and str(c).strip() != "TOTAL":
                existing_countries.add(str(c).strip())

        for country in sorted(countries.keys()):
            if country not in existing_countries:
                log.info("New country '%s' for product '%s' — appending row", country, prod)
                # Insert before TOTAL row
                total_row = r["total_row"]
                if total_row:
                    ws.insert_rows(total_row)
                    new_row = total_row
                    # Shift ranges
                    r["total_row"] = total_row + 1
                    r["last_country_row"] = new_row

                    ws.cell(row=new_row, column=1, value=prod)
                    ws.cell(row=new_row, column=2, value=country)

                    quota_cell = ws.cell(row=new_row, column=3, value=prod_data.get("max_quota", 0))
                    quota_cell.number_format = "#,##0"

                    share_cap_cell = ws.cell(row=new_row, column=4, value=prod_data.get("max_share", 0.0))
                    share_cap_cell.number_format = "0.0%"

                    cell = ws.cell(row=new_row, column=new_date_col, value=countries[country])
                    cell.number_format = "0.0%"

                    # Update TOTAL SUM range
                    col_letter = get_column_letter(new_date_col)
                    total_cell = ws.cell(row=r["total_row"], column=new_date_col)
                    total_cell.value = f"=SUM({col_letter}{r['first_row']}:{col_letter}{r['last_country_row']})"
                    total_cell.number_format = "0.0%"
                    total_cell.fill = YELLOW_FILL

    # Rewrite OVER column in new position (one column to the right)
    # The old OVER column is now the new date column — do NOT clear it.
    over_header = ws.cell(row=2, column=new_over_col, value="OVER")
    over_header.font = BOLD_FONT
    over_header.fill = YELLOW_FILL

    # Recalculate OVER for all data rows
    for row in range(3, ws.max_row + 1):
        country = ws.cell(row=row, column=2).value
        if not country or str(country).strip() == "TOTAL":
            ws.cell(row=row, column=new_over_col).value = None
            continue

        max_share = ws.cell(row=row, column=4).value
        if max_share is None or max_share == 0:
            ws.cell(row=row, column=new_over_col).value = None
            continue

        latest_share = ws.cell(row=row, column=new_date_col).value
        if latest_share is not None and isinstance(latest_share, (int, float)):
            ws.cell(row=row, column=new_over_col, value="YES" if latest_share >= max_share else None)
        else:
            ws.cell(row=row, column=new_over_col).value = None

    log.info("Updated sheet '%s' with column '%s'", ws.title, date_header)


def create_b1_sheet(wb: Workbook, b1_data: dict | None, trq_quarter: str, today: date):
    """Create or rewrite the B1 Imports sheet."""
    sheet_name = "B1 Imports"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(title=sheet_name)

    # Determine calendar quarter label
    cal_start_m, cal_end_m = TRQ_TO_B1_CALENDAR_MONTHS[trq_quarter]
    month_names = ["", "January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"]
    year = today.year
    cal_label = f"{month_names[cal_start_m]} 1 to {month_names[cal_end_m]} {30 if cal_end_m in (6, 9, 11) else 31}, {year}"

    # Row 1: Title
    ws.cell(row=1, column=1, value=f"B1 Import Data - Calendar Quarter: {cal_label}")

    # Rows 3-7: Disclaimers
    ws.cell(row=3, column=1,
            value="NOTE: B1 data reflects actual customs entries using calendar quarters (Jan-Mar, Apr-Jun, etc.).")
    ws.cell(row=4, column=1,
            value="TRQ utilization data uses offset quarters (Mar 26-Jun 26, etc.) and tracks permit issuance, not customs entries.")
    ws.cell(row=5, column=1,
            value="The two datasets have ~5-6 day boundary misalignment and measure different aspects of the import process.")
    ws.cell(row=6, column=1,
            value="B1 imports are matched at the 8-digit tariff item level using the HTS codes listed in the 'HTS code covered' sheet.")
    ws.cell(row=7, column=1,
            value="These HTS codes are a selected subset; the official TRQ product scope may include additional tariff items not tracked here.")

    # Row 9: Headers
    headers = ["Product Category", "Country", "Total Tonnes", "Total C$1000", "Avg C$/Tonne"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=9, column=col_idx, value=header)
        cell.font = BOLD_FONT

    if b1_data is None:
        ws.cell(row=11, column=1, value="B1 data not available for the matching calendar quarter.")
        log.info("B1 sheet created (no data)")
        return

    # Write data
    row_num = 10
    for product in TRACKED_PRODUCTS:
        if product not in b1_data:
            continue
        countries = b1_data[product]
        for country in sorted(countries.keys()):
            d = countries[country]
            tonnes = d["tonnes"]
            value = d["value"]
            avg_price = value * 1000 / tonnes if tonnes > 0 else 0.0

            ws.cell(row=row_num, column=1, value=product)
            ws.cell(row=row_num, column=2, value=country)

            t_cell = ws.cell(row=row_num, column=3, value=tonnes)
            t_cell.number_format = "#,##0.00"

            v_cell = ws.cell(row=row_num, column=4, value=value)
            v_cell.number_format = "#,##0.00"

            p_cell = ws.cell(row=row_num, column=5, value=avg_price)
            p_cell.number_format = "#,##0.00"

            row_num += 1

    log.info("B1 sheet created with %d data rows", row_num - 10)


def create_hts_sheet(wb: Workbook):
    """Create the static HTS code reference sheet."""
    sheet_name = "HTS code covered"
    if sheet_name in wb.sheetnames:
        return  # Already exists

    ws = wb.create_sheet(title=sheet_name)

    # Flat products (column A) and Long products (column B)
    flat_products = ["Hot-Rolled Sheet", "Steel Plate", "Cold-Rolled Sheet", "Coated Steel Sheet"]
    long_products = ["Rebar", "Hot-Rolled Bar", "Wire Rod", "Structural Steel"]

    row = 1
    # Write flat products
    for product in flat_products:
        ws.cell(row=row, column=1, value=product)
        row += 1
        for code in HTS_CODE_MAP[product]:
            ws.cell(row=row, column=1, value=code)
            row += 1
        row += 1  # blank row between products

    # Write long products in column B (restart from row 1)
    row = 1
    for product in long_products:
        ws.cell(row=row, column=2, value=product)
        row += 1
        for code in HTS_CODE_MAP[product]:
            ws.cell(row=row, column=2, value=code)
            row += 1
        row += 1

    log.info("HTS code reference sheet created")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    today = date.today()
    quarter = get_current_quarter(today)
    log.info("Today: %s, TRQ Quarter: %s", today, quarter)

    start_date, end_date = get_quarter_date_range(quarter, today)
    log.info("Quarter range: %s to %s", start_date, end_date)

    # Download TRQ CSVs
    fta_csv = download_csv("FTA", quarter)
    nfta_csv = download_csv("NFTA", quarter)

    # Fallback to previous quarter if current not available
    if fta_csv is None or nfta_csv is None:
        prev_quarters = {"Q1": "Q4", "Q2": "Q1", "Q3": "Q2", "Q4": "Q3"}
        prev_q = prev_quarters[quarter]
        log.info("%s data not yet available, trying %s...", quarter, prev_q)
        if fta_csv is None:
            fta_csv = download_csv("FTA", prev_q)
        if nfta_csv is None:
            nfta_csv = download_csv("NFTA", prev_q)
        if fta_csv is None or nfta_csv is None:
            log.error("Could not download TRQ CSV data. Exiting.")
            sys.exit(1)
        quarter = prev_q
        start_date, end_date = get_quarter_date_range(quarter, today)

    # Parse TRQ data
    log.info("Parsing FTA CSV...")
    fta_data = parse_trq_csv(fta_csv)
    log.info("Parsing NFTA CSV...")
    nfta_data = parse_trq_csv(nfta_csv)

    # B1 import data
    b1_data = None
    if should_fetch_b1(quarter, today):
        b1_data = scrape_b1_imports()
        if b1_data is None:
            log.warning("B1 scraping returned no data for tracked products")
    else:
        log.info("B1 data for matching calendar quarter not yet available (TRQ %s, current month %d)",
                 quarter, today.month)

    # Load or create workbook
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if os.path.exists(OUTPUT_FILE):
        log.info("Loading existing workbook: %s", OUTPUT_FILE)
        wb = load_workbook(OUTPUT_FILE)
    else:
        log.info("Creating new workbook")
        wb = Workbook()
        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    # Update or create TRQ sheets
    for trq_type, data, prefix in [("NFTA", nfta_data, "non-FTA"), ("FTA", fta_data, "FTA")]:
        sheet_name = f"{prefix} {quarter}"
        if sheet_name in wb.sheetnames:
            log.info("Updating existing sheet '%s'...", sheet_name)
            update_trq_sheet(wb[sheet_name], data, today)
        else:
            log.info("Creating new sheet '%s'...", sheet_name)
            create_trq_sheet(wb, sheet_name, quarter, trq_type, data, today)

    # B1 sheet
    create_b1_sheet(wb, b1_data, quarter, today)

    # HTS reference sheet
    create_hts_sheet(wb)

    # Save
    wb.save(OUTPUT_FILE)
    log.info("Workbook saved to %s", OUTPUT_FILE)
    log.info("Done.")


if __name__ == "__main__":
    main()
