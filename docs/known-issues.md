# Known Issues & Lessons Learned — Canada TRQ Tracker

A record of bugs that have bitten this project, so future work (and future AI
sessions) can avoid reintroducing them. **Read this before changing the Excel
writer or the B1 scraper.**

---

## Architecture reminder (read first)

Two layers, kept separate on purpose:

1. **Parsing layer** — `parse_trq_csv()` and `scrape_b1_imports()` turn raw
   government data into plain Python dicts. Verified correct against live sources
   (values match the source tables exactly).
2. **Excel writer layer** — `create_trq_sheet()` / `update_trq_sheet()` /
   `create_b1_sheet()` turn those dicts into the workbook.

> **Most bugs here live in the writer, not the parser.** "Scraped the right value"
> and "wrote the right value to the right cell" are different claims — always
> validate the *saved workbook*, not just the parsed dict.

### Delivery chain (why a local fix isn't enough)

```
Monday cron → GitHub Actions runs the script (cloud, from the repo)
           → commits data/canada_trq_tracker.xlsx AND publishes it to the "latest" Release
           → colleague's download_latest.bat pulls from releases/latest/download/
```

A fix reaches the end user only after it is **committed and pushed** and a run
completes (scheduled Monday, or manual `workflow_dispatch`). Push fixes **before
the next Monday run**, or the old code publishes another file first.

---

## Bug 1 — Weekly update corrupted the TRQ sheets (position drift) — FIXED (commit 13097dc)

**Severity:** critical (silently produced garbage data).

**Symptom:** on an incremental weekly update, in a week where **new countries
appeared**: SUM formulas written *into country data cells*; rows of one product
interleaved into another product's block; the same country appearing twice; TOTAL
`SUM()` ranges pointing at the wrong rows. The **first run** (build-from-scratch
`create_trq_sheet`) was always clean, which hid the bug while only one column
existed.

**Root cause:** `update_trq_sheet` cached product row ranges once, then called
`ws.insert_rows()` in a loop to add new-country rows. `insert_rows` physically
shifts every row below the insertion point, but the cached indices were not
recomputed, openpyxl does not rewrite formula references, and values written
earlier by position got pushed around. Mixing "compute positions once" with
"mutate positions in a loop" is the trap.

**Fix (as shipped):** after each `insert_rows`, the code now (a) clears the new
date column first, and (b) **shifts the row references of all downstream products
down by one** (`update_trq_sheet`, the "Shift ALL subsequent products' row
references" block). This keeps the cached ranges consistent. Verified clean across
10 weekly columns of real data.

> An alternative, arguably more robust design is **read → merge → rewrite**: read
> the whole sheet into a model, merge the new column, and re-render every row from
> scratch so positions can never desync. The shipped fix is the surgical version;
> if this code is ever rewritten, prefer read-merge-rewrite.

**Prevention:**
- Never use `insert_rows` / `delete_rows` against indices cached *before* the
  mutation. Recompute, insert bottom-up, or rebuild from a model.
- After any writer change, dump the saved workbook and check: contiguous product
  blocks, no formulas in data cells, no duplicate countries, each TOTAL `SUM()`
  spans exactly its product's country rows. Don't trust the parsed dict.
- Test the update path against a **multi-column, new-country** scenario, not just a
  fresh first run.

---

## Bug 2 — B1 imports summed the wrong time window & were mislabeled — FIXED (commit e837ab4)

**Severity:** medium (numbers arithmetically correct, but covered the wrong period
and the label was false).

**Symptom:** the B1 page serves **year-to-date** data (e.g. Jan 1 – May 23), not a
single calendar quarter. `scrape_b1_imports()` summed **every** row with no month
filter, while the sheet was labeled "Calendar Quarter: April 1 to June 30". So Q4
import totals wrongly included Jan–Mar (TRQ Q3) volumes.

**Root cause:** the design intended quarter alignment — `TRQ_TO_B1_CALENDAR_MONTHS`
maps each TRQ quarter to its B1 calendar months (`"Q4": (4, 6)`) — but the scraper
never *applied* the filter. (A code-review pass added month *validation* `1..12` to
the 6-cell parser, which is not the same as *filtering* to the quarter.) A constant
that exists but is never read is a strong signal of an unfinished feature.

> Original intent (Laura's letter): match utilization with imports **"over the same
> time period"** as the quota quarter — i.e. align to the quarter, not YTD.

**Fix (as shipped):** `scrape_b1_imports(month_range)` takes the quarter's month
range and skips rows whose `Month` is outside it; `main()` passes
`TRQ_TO_B1_CALENDAR_MONTHS[quarter]`. A note row on the B1 sheet states the months
used and that the quarter may be partially complete. Verified: filtered totals match
an independent Apr–May re-sum of the detail rows exactly (Rebar 35,595.90 t,
Hot-Rolled Sheet 21,110.17 t).

**Prevention:**
- B1 is intentionally *approximate* (permits vs customs; ~5–6 day offset; HTS
  subset) — keep the disclaimer notes; don't force it to match TRQ exactly.
- If a config constant maps inputs to a behavior, confirm it is actually consumed.
- The B1 quarter may always be partial (customs data lags). That is expected.

---

## Validation recipe

See the "Validating correctness by hand" section of [../README.md](../README.md):
download the TRQ CSV and check Part A/Part B alignment; cross-check B1 totals against
the government's own `Summary of HS 10 Code…` rows (which the scraper ignores). For
month-filtered B1, re-sum the detail rows yourself for the allowed months only.
