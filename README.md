# Canada TRQ Weekly Tracker

Automated monitoring of Canadian steel Tariff Rate Quota (TRQ) utilization and B1
import data. A single Python script downloads the government data each week and
writes it into an Excel workbook that matches Laura's template, published via a
GitHub Actions cron and a GitHub Release.

> **New here? Read [docs/known-issues.md](docs/known-issues.md) before touching the
> Excel writer or the B1 scraper** — it records bugs that have already bitten this
> project and how to avoid reintroducing them.

---

## Current status

- **Latest meaningful change (2026-07-06):** hardening pass — TOTAL formulas are
  now rewritten from the final sheet layout every run (completes the Bug 1 fix;
  the shipped workbook's stale Q4 formulas were repaired in the same commit),
  zero-utilization products get a literal 0 instead of a circular `=SUM()`, the
  parser locates the Part A/Part B boundary dynamically and fails loud on a
  structure change, and a pytest suite (`tests/`) runs on every push.
- **Previous milestones:** landing-page discovery replacing constructed URLs
  (Bug 4, commit 1b18dd1); B1 month filtering (Bug 2).
- **Verified correct (against live sources):** TRQ utilization values, per-country
  shares, Max Quota / Max Share, OVER flags, TOTAL formulas across ALL date
  columns (enforced by `tests/wb_invariants.py`), the weekly multi-column update
  path incl. new-country insertions, and B1 aggregation arithmetic.
- The weekly cron has been running and committing since 2026-03-30.

To see exactly what changed last, use `git log --oneline` — that is the source of
truth for "latest version", not this file.

---

## What it produces

`data/canada_trq_tracker.xlsx`, with these sheets:

| Sheet | Contents |
|---|---|
| `non-FTA Q<n>` / `FTA Q<n>` | Per-country quota-share %, one new dated column each week, TOTAL rows (SUM formulas), and an OVER flag when a country reaches its single-country cap. |
| `B1 Imports` | Actual customs import tonnes/value per product+country, for the quarter-aligned months. |
| `HTS code covered` | Static reference: HTS codes mapped to each tracked product. |

Only **8 of 23** TRQ products are tracked (see `TRACKED_PRODUCTS` in the script).

---

## How to run

```bash
pip install -r requirements.txt
python canada_trq_tracker.py
```

- No arguments. It discovers the current TRQ quarter (label, CSV URLs, and date
  range) from the official landing page, downloads the live CSV/HTML, and
  **updates** `data/canada_trq_tracker.xlsx` in place (creating it on first run).
  Re-running the same day is idempotent (it skips an existing date column).
- Windows + POSIX both supported (date formatting branches on `platform.system()`).
- **Caution — local runs mutate the bot-committed workbook.** `git pull` first,
  and don't push a locally-mutated `data/…xlsx` unless that's what you intend:
  the delivery chain publishes whatever the next CI run finds in the repo.
- Tests: `pip install -r requirements-dev.txt && python -m pytest tests/ -q`
  (offline; also run by CI on every push).

## Architecture (two layers — keep them separate)

1. **Parsing layer** — `parse_trq_csv()`, `scrape_b1_imports()`: raw government
   data → plain Python dicts. Verified correct against the sources.
2. **Excel writer layer** — `create_trq_sheet()` / `update_trq_sheet()` /
   `create_b1_sheet()`: dicts → workbook. **Historically where bugs live.** Always
   validate the *saved workbook*, not just the parsed dict.

## Data sources

- TRQ CSVs: **discovered from the official landing page** (`LANDING_URL` in the
  script) — the CSV filename, quarter label, and date range are read from the
  page, never constructed, because the source renames files across program years
  (`TRQ_FTA-Q1.csv` → `TRQ_FTA-Y2Q1.csv`; see known-issues Bug 4). Two format
  variants exist — old `ExecutionTime` header, new `Textbox` header; the parser
  detects and handles both.
- B1 imports: `https://www.eics-scei.gc.ca/report-rapport/b1.htm` (one large HTML
  table; the page currently serves year-to-date, so the scraper filters to the
  quarter-aligned months in `TRQ_TO_B1_CALENDAR_MONTHS`).

Full design notes: [docs/superpowers/specs/2026-03-30-canada-trq-tracker-design.md](docs/superpowers/specs/2026-03-30-canada-trq-tracker-design.md).

## Delivery to the end user

```
Monday cron → GitHub Actions runs the script (cloud, from this repo)
           → commits data/canada_trq_tracker.xlsx AND publishes it to the "latest" Release
           → colleague's download_latest.bat pulls from releases/latest/download/
```

A fix reaches the end user only after it is **pushed** and a run completes
(scheduled Monday, or a manual `workflow_dispatch`). Fixing locally does nothing
for them — and pushing **before the next Monday run** matters.

## Validating correctness by hand

Because the data is values lifted from large government tables, verify against the
source, not just the script's output:

- **TRQ:** download the CSV. Part A (lines 5–27) = 23-product summary; Part B =
  per-country sections positionally aligned to Part A items (zero-utilization
  products have an empty section that must be skipped). Spot-check a few products'
  shares against the saved sheet.
- **B1:** the government prints its own `Summary of HS 10 Code…` rows, which the
  scraper **ignores** — making them a perfect independent cross-check. For the
  month-filtered total, re-sum the detail rows yourself for the allowed months.
- **Encoding:** names like "Türkiye" (U+00FC) are stored correctly; a `�` in a
  terminal is just the console, not corruption — confirm with
  `[hex(ord(c)) for c in value]`.

## Files

| Path | Purpose |
|---|---|
| `canada_trq_tracker.py` | The whole program. |
| `requirements.txt` | Pinned deps: requests, beautifulsoup4, lxml, openpyxl. |
| `requirements-dev.txt` | Test-only deps (pytest). |
| `.github/workflows/weekly_scrape.yml` | Monday cron + Release publishing. |
| `.github/workflows/retry_on_failure.yml` | Auto re-run on a fresh runner when the weekly job fails (IP-block mitigation, capped at 2 retries). |
| `.github/workflows/tests.yml` | Runs the offline test suite on every push/PR. |
| `tests/` | pytest suite: parser (incl. real captured fixtures), writer invariants, discovery. |
| `data/canada_trq_tracker.xlsx` | Output (committed; also published to Release). |
| `Canadian Quota Template.xlsx` | Laura's reference template (format target only — the script never reads it). |
| `guide.html` | End-user usage guide. |
| `download_latest.bat` | End-user one-click downloader (pulls the Release). |
| `docs/known-issues.md` | Bug history & prevention — read before editing. |
| `docs/superpowers/specs/…design.md` | Original design spec (quarter-URL section superseded; see marker inside). |
