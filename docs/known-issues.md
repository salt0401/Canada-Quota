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

**Changed 2026-08-02 — the weekly run moved to the MEPS company server.**

```
Monday 13:00 local → Task Scheduler on WIN-RE1UH50A07U
                     runs tools\server-weekly-task.ps1 -Push
                  → commits data/canada_trq_tracker.xlsx to the repo
                  → uploads it to the "latest" Release
                  → colleague's download_latest.bat pulls from releases/latest/download/
```

A fix reaches the end user only after it is **committed, pushed, AND pulled onto
the server**, and a run completes. That last step is new and is the easy one to
forget: GitHub Actions checked out `master` on every run, so pushing deployed
itself; the server runs whatever was last pulled onto it. See
[SERVER_DEPLOYMENT.md](SERVER_DEPLOYMENT.md) → "Updating the code on the server".

Do it **before the next Monday run**, or the old code publishes another file
first.

---

## Bug 1 — Weekly update corrupted the TRQ sheets (position drift) — FIXED COMPLETELY 2026-07-06 (first fix, commit 13097dc, was incomplete)

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

**First fix (commit 13097dc) — INCOMPLETE:** after each `insert_rows`, the code
(a) cleared the new date column first, and (b) shifted the cached row references
of all downstream products. That fixed **value placement**, but not **formulas
already written to the sheet**: `insert_rows` physically moves cells without
rewriting formula text, so every TOTAL `=SUM()` below an insertion point — in
the just-written column AND in every historical column — kept its stale range.
The claim originally recorded here ("verified clean across 10 weekly columns")
was wrong: the verification only checked the newest column, which is always
written last from fresh ranges and therefore always looks correct. A 2026-07-01
full-project review found **131 corrupted TOTAL formulas** live in the two Q4
sheets of the shipped workbook.

**Complete fix (2026-07-06):** `update_trq_sheet` no longer writes TOTAL
formulas during the value pass at all. After all insertions are done,
`_rewrite_total_formulas()` re-derives every product's TOTAL in **every** date
column from the final physical layout (values always sit on the right rows, so
this is a pure formula repair and is idempotent when nothing shifted — it also
heals any corruption left by earlier runs). The committed workbook's 131 stale
Q4 formulas were repaired in the same commit. Zero-country products get a
literal `0.0` instead of a self-referential `=SUM(Fn:Fn)` circular reference
(`_write_total_cell`). Both paths are locked in by `tests/test_writer.py`
against `tests/wb_invariants.py`.

> An alternative, arguably more robust design is **read → merge → rewrite**: read
> the whole sheet into a model, merge the new column, and re-render every row from
> scratch so positions can never desync. The shipped fix is the surgical version;
> if this code is ever rewritten, prefer read-merge-rewrite.

**Prevention:**
- Never use `insert_rows` / `delete_rows` against indices cached *before* the
  mutation. Recompute, insert bottom-up, or rebuild from a model.
- openpyxl **never** updates formula strings on row insert/delete. Any formula
  written before a mutation is suspect; write formulas LAST, from final layout.
- After any writer change, check the saved workbook against
  `tests/wb_invariants.py` (the executable version of this checklist):
  contiguous product blocks, no formulas in data cells, no duplicate countries,
  each TOTAL `SUM()` spans exactly its product's country rows **in every date
  column, not just the newest one**. Don't trust the parsed dict.
- Test the update path against a **multi-column, new-country** scenario, not just a
  fresh first run (`tests/test_writer.py` does this on every push).

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

## Bug 3 — Weekly run failed: source host dropped the GitHub Actions runner's IP — FIXED (networking hardened)

**Severity:** high (no data published that week; the job hard-fails and emails an
alert, so it's loud, not silent — but the colleague gets no update).

**Symptom:** the 2026-06-29 scheduled run took 4m31s (vs. the usual ~25s) and
exited 1 with `Could not download TRQ CSV data. Exiting.` Every attempt logged
`ConnectTimeoutError ... Connection to www.eics-scei.gc.ca timed out
(connect timeout=30)`. The four prior Mondays all succeeded unchanged.

**Root cause:** environmental, not a code bug. `www.eics-scei.gc.ca` sits behind
a firewall that **silently drops TCP connections from some cloud/datacenter
egress IPs** (the GitHub-hosted Azure runners). The tell is that the failure is a
*connect timeout* (the TCP handshake never completes) rather than an HTTP 403/404
— the request never reaches the application layer. Verified during debugging: from
an ordinary connection the host returns HTTP 200 with full data **even with the
default `python-requests` User-Agent**, so it is not a UA/WAF block; the only
variable is the source IP. GitHub's runner IPs rotate, so most runs draw a clean
IP (success) and the occasional run draws a flagged one (connect timeout). The
script was correct to exit 1 — it just couldn't reach the source.

**Fix (as shipped):** networking hardened in `canada_trq_tracker.py` — a shared
`requests.Session` (`SESSION`) with a browser-like User-Agent, a `urllib3 Retry`
adapter (bounded connect/read/status retries with exponential backoff), and split
`(connect, read)` timeouts. Both download sites (`download_csv`, B1 scraper) use
it. This survives transient drops/slow responses and fails a blocked connect fast.

> **Important limitation:** retries inside a single run all use the *same* runner
> IP, so they cannot beat a *hard* block. The only cure for a blocked IP is a
> **fresh egress IP** — i.e. re-run the job on a new runner, which provisions a
> new VM/IP. This is automated: `.github/workflows/retry_on_failure.yml` (commit
> 293c085) re-runs a failed weekly job on a fresh runner via `workflow_run`,
> capped at two automatic retries via `run_attempt`. Manual fallback: "Re-run
> jobs" in the Actions UI, or `gh run rerun <id> --failed`.

### 2026-08-02 update — this bug shaped, and now constrains, the server move

The weekly run moved to the MEPS company server, whose address is **also a
datacenter IP** (IONOS, `212.227.127.169`). Whether this bug applied to it was
therefore the gating question for the whole migration, not a detail — so it was
measured before any code was written. With the production User-Agent, against
the URLs the script actually fetches:

| Target | Result from the company server, 2026-08-02 |
|---|---|
| `international.gc.ca` TRQ landing page | **200** |
| `eics-scei.gc.ca/report-rapport/TRQ_FTA-Y2Q1.csv` | **200** |
| `eics-scei.gc.ca/report-rapport/TRQ_NFTA-Y2Q1.csv` | **200** |
| `eics-scei.gc.ca/report-rapport/b1.htm` | **200** |

Confirmed beyond the status code by a full end-to-end run that parsed real data
from all four.

**But the move removes the cure.** GitHub rotated runner IPs, so a blocked draw
self-healed on re-run — that is the entire premise of `retry_on_failure.yml`.
The server has **one fixed address and no equivalent**. In-run retries still
cannot beat a hard block, and now there is no fresh IP to fall back to
automatically.

The fallback is therefore manual and deliberate: **GitHub → Actions → "Weekly
Canada TRQ Update" → Run workflow**, which is why that workflow is kept with its
schedule commented out rather than deleted. It draws a GitHub runner IP, and
`retry_on_failure.yml` still auto-retries it on a fresh runner if it fails.

⚠️ Only dispatch it once the server task is confirmed not to have run that week —
otherwise the two race on `git push`.

**A pass on 2026-08-02 is not a guarantee.** The block is reputation-based and
IP-specific; the symptom to watch for is the same as ever, a *connect timeout*
rather than an HTTP status. Re-test with
`meps-server-docs/scripts/validate-targets.ps1`.

**Prevention:**
- Don't "fix" this by only adding a User-Agent or more in-run retries — the block
  is at the network layer, below HTTP, and pinned to the run's IP.
- A sudden jump from ~25s to multi-minute runtime is the fingerprint of network
  timeouts/retries, not heavier work. Check runtime deltas first.
- Keep the hard `sys.exit(1)` on download failure — publishing stale/empty data
  would be worse than a loud failure email.

---

## Bug 4 — Program-year rollover renamed the files; scraper silently tracked the closed quarter — FIXED (landing-page discovery)

**Severity:** high (silent — green checkmark, fresh file, *wrong/closed quarter*).

**Symptom:** found by verifying output against the live source on 2026-06-30. The
scraped numbers were correct, but for the just-**closed** Q4 — while the live
current quarter had already moved on. Every weekly run would have kept appending
duplicate frozen-Q4 columns and never tracked the active quarter.

**Root cause:** the TRQ **program year rolled over** and the source renamed the
quarter files. The current quarter became `TRQ_FTA-Y2Q1.csv` (report title "Year
2, Quarter 1: June 28, 2026 to September 29, 2026") — a new `Y2` (year-2) prefix.
The old code built URLs as `TRQ_{FTA,NFTA}-{Q1..Q4}.csv` with no year concept, so
it requested `TRQ_FTA-Q1.csv` (now **404**) and `TRQ_NFTA-Q1.csv` (now **last
year's** 2025 Q1), then the **prev-quarter fallback silently used the closed Q4**.
The new year also changed quotas/caps (Hot-Rolled Sheet 9,523,100 KG/67% →
8,569,800/71%; Steel Plate 30→33%; Cold-Rolled 32→35%) and shifted boundary dates
(Q4 ends Jun 27, not the hard-coded Jun 26), so the stale quarter was wrong on
baselines too.

**Fix (as shipped):** stop constructing URLs/dates. `discover_current_reports()`
reads the official landing page (`LANDING_URL`), pairs each "Quarter N: <dates>"
link with its adjacent `[.CSV]` link by shared filename stem, and picks, for both
FTA and NFTA, the report whose **date range contains today** (most-recent started
quarter on a gap). It returns the real `csv_url`, a `Q1..Q4` label, and the date
strings. `main()` uses those; `download_csv(url)` now takes a full URL.

**Critical design rules baked in (do not regress):**
- **Fail loud, never fall back.** The old silent prev-quarter fallback is *removed*.
  If discovery fails (page unreachable, structure changed, no current report, CSV
  link malformed, or FTA/NFTA disagree on the quarter) → `log.error` + `sys.exit(1)`.
  Silent fallback is exactly what hid this bug. (The auto-retry workflow from Bug 3
  re-runs on a fresh IP, so a transient discovery failure self-heals.)
- **Read the `[.CSV]` href verbatim, but validate it ends in `.csv`.** The gov page
  has at least one malformed link (FTA Q2's `[.CSV]` points to a `.htm`); never
  reconstruct the filename (casing is inconsistent: `TRQ_nFTA-q2` vs `TRQ_FTA-Q4`).
- **The `Q1..Q4` label still drives the TRQ↔B1 join, unchanged.**
  `TRQ_TO_B1_CALENDAR_MONTHS`, `should_fetch_b1`, sheet names, and the B1 sheet all
  use the label exactly as before — that mapping is year-independent (Year-2 Q1
  still overlaps Jul–Sep). Selection is by **date range**, not by matching
  "Quarter N" text (the same label repeats across years).
- **B1 history is preserved on a quarter transition.** The single "B1 Imports"
  sheet is deleted/rewritten each run, so `main()` renames the existing one to
  `B1 Imports {prev_q}` when the new quarter's TRQ sheets don't exist yet —
  otherwise the just-closed quarter's customs data would be silently overwritten
  with "not available" until the new B1 calendar window opens (July for Q1).

**Known remaining limitation (by design, documented):** sheet names are still
year-blind (`f"{prefix} {quarter}"`). At the *next* program-year rollover, next
year's "FTA Q1" will land in the same sheet as this year's and `update_trq_sheet`
would mix two years' columns. Acceptable because the landing-page discovery only
needs to be re-checked roughly yearly anyway — but if this runs unattended past a
second rollover, give the sheet names a year suffix first.

**Prevention:**
- Verify output against the **live source**, not just that the run is green — a
  passing pipeline can silently scrape the wrong thing.
- Anything keyed on a hard-coded period (`Q1..Q4`, fixed dates, constructed
  filenames) is fragile across a year boundary. Discover from the source of truth.

---

## Hardening pass 2026-07-06 — fail-loud extended to the parse & writer layers

Bug 4's lesson ("a green run can publish wrong data") originally produced
fail-loud guards only at the network layer. They now cover the whole pipeline;
all of these end the run with `log.error` + `sys.exit(1)` **before** the CI
commit step can publish anything:

- **Parse shape:** Part A row count ≠ 23; ALL tracked products missing from the
  parse; ≥ `_PART_AB_MISMATCH_LIMIT` products whose Part A/Part B totals
  disagree (a couple of mismatches remain warn-only by design — Bug 2's
  per-product anomaly stance is unchanged, the threshold only catches
  *systematic* desync).
- **Writer:** a TRQ sheet with no OVER header (previously warn-and-skip, which
  silently froze that sheet forever on a green run).
- **Workbook lifecycle:** a missing `data/canada_trq_tracker.xlsx` refuses to
  silently rebuild (all history would be lost and published); restore it or set
  `TRQ_ALLOW_NEW_WORKBOOK=1` for a deliberate init. A corrupt existing file
  also aborts.
- **Post-save gate:** every run saves to `data/canada_trq_tracker.tmp.xlsx`
  (a real `.xlsx` extension — openpyxl refuses to open `*.tmp`), re-opens it,
  and runs `validate_workbook()` (the in-script version of the Bug 1 invariant
  checklist, checking EVERY date column). Only a validated file replaces the
  real one; a failing file is kept for inspection and the run exits 1, so the
  auto-retry fires and the failure emails.

### Environment note — file locks during the atomic save

**Updated 2026-08-02. The original form of this note is now stale in one respect
and newly relevant in another.**

*Then:* the working copy sat in a OneDrive-synced folder, and OneDrive could
transiently lock or sync-shuffle files under `.git/` and lock the xlsx mid-sync.
The note recommended moving development outside OneDrive.

*Now:* **that move happened.** The development clone lives at
`C:\dev\project\02 Monitored\Canada Quota` — verified 2026-08-02 to be outside
the OneDrive root (`C:\Users\<user>\OneDrive - MEPS`). OneDrive is no longer in
the picture for the repository.

The **retry logic in the save path is still load-bearing**, though, for a
different reason. The weekly run now happens on the MEPS company server, where
Windows Defender real-time scanning *and* Acronis Active Protection both watch
file writes with **no exclusions** (a standing owner ruling). The predicted
symptom is identical to the old OneDrive one: an intermittent `PermissionError`
/ `WinError 32` on `os.replace`, on a file this code just wrote, which succeeds
on an immediate re-run. That is why `main()` retries the swap three times, two
seconds apart, and preserves the new data in `data/canada_trq_tracker.tmp.xlsx`
(gitignored) rather than losing it.

**Do not remove that retry as dead code.** Locally it now guards against Excel
having the workbook open; on the server it guards against the antivirus pair.
See [SERVER_DEPLOYMENT.md](SERVER_DEPLOYMENT.md) → "The antivirus failure
signature".

---

## Validation recipe

See the "Validating correctness by hand" section of [../README.md](../README.md):
download the TRQ CSV and check Part A/Part B alignment; cross-check B1 totals against
the government's own `Summary of HS 10 Code…` rows (which the scraper ignores). For
month-filtered B1, re-sum the detail rows yourself for the allowed months only.
