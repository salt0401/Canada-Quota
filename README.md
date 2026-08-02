# Canada TRQ Weekly Tracker

Automated monitoring of Canadian steel Tariff Rate Quota (TRQ) utilization and B1
import data. A single Python script downloads the government data each week and
writes it into an Excel workbook that matches Laura's template, published to a
GitHub Release.

> **New here? Read [docs/known-issues.md](docs/known-issues.md) before touching the
> Excel writer or the B1 scraper** — it records bugs that have already bitten this
> project and how to avoid reintroducing them.
>
> **Operating it? Read [docs/SERVER_DEPLOYMENT.md](docs/SERVER_DEPLOYMENT.md).**
> Since 2026-08 the weekly run happens on the MEPS company server, not on GitHub
> Actions — which means **pushing a fix no longer reaches the pipeline by
> itself**.

---

## Current status

- **Latest meaningful change (2026-08-02): the weekly run moved off GitHub
  Actions onto the MEPS company server** (`C:\DataScienceProject\CanadaQuota`,
  Windows Task Scheduler, Mondays 13:00 local). Nothing about the pipeline, the
  workbook format, the repository or Laura's downloader changed — only the
  machine. Full detail, including everything that was measured rather than
  assumed, in [docs/SERVER_DEPLOYMENT.md](docs/SERVER_DEPLOYMENT.md). Three
  consequences worth knowing up front:
  - **A pushed fix no longer deploys itself.** GitHub Actions checked out
    `master` every run; the server does not. See "Updating the code on the
    server" in that document.
  - **The failure email is gone**, because GitHub is no longer running the job.
    `.github/workflows/data-freshness-watchdog.yml` replaces it and opens an
    issue instead.
  - **`weekly_scrape.yml` is kept, schedule disabled**, as the emergency
    fallback — and as the only escape from Bug 3, since the server has one
    fixed IP and GitHub runners do not.
- **Previous meaningful change (2026-07-06):** hardening pass — TOTAL formulas are
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
- The weekly cron has been running and committing since 2026-03-30 (on GitHub
  Actions until 2026-08-02, on the company server since).

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
  **updates** `data/canada_trq_tracker.xlsx` in place. Re-running the same day
  is idempotent (it skips an existing date column).
- **Fail-loud guards** (`sys.exit(1)`): discovery/download failure, CSV parse
  failure (wrong Part A count, all tracked products missing, widespread Part
  A/B mismatch), a TRQ sheet missing its OVER header, a **missing workbook**
  (set `TRQ_ALLOW_NEW_WORKBOOK=1` only for a deliberate first-time init — a
  silent rebuild would publish a history-less file), and a saved workbook that
  fails `validate_workbook()` — every run re-opens the file it just wrote and
  checks the known-issues invariants before the pipeline can commit it.
- Windows + POSIX both supported (date formatting branches on `platform.system()`).
- **Caution — local runs mutate the bot-committed workbook.** `git pull` first,
  and don't push a locally-mutated `data/…xlsx` unless that's what you intend:
  the server's weekly run starts from whatever is committed, and a pushed
  mutation would also make the server's next push conflict.
- **This is not how the pipeline runs.** The live path is
  `tools\server-weekly-task.ps1 -Push` on the company server, which wraps this
  script with a UTC/local date guard, a working-tree guard, a
  workbook-advanced assertion, the commit/push and the Release upload. Running
  `canada_trq_tracker.py` by hand exercises the scraper and the writer, not the
  delivery chain.
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

Full design notes: [docs/design-2026-03-30.md](docs/design-2026-03-30.md).

## Delivery to the end user

```
Monday 13:00 local → Task Scheduler on the MEPS company server
                     runs tools\server-weekly-task.ps1 -Push
                  → commits data/canada_trq_tracker.xlsx to this repo
                  → uploads it to the "latest" Release
                  → colleague's download_latest.bat pulls from releases/latest/download/
```

⚠️ **Pushing a fix is no longer enough.** GitHub Actions checked out `master` on
every run, so a push deployed itself. The server runs whatever was last pulled
onto it, and the weekly task deliberately does not pull at the start of a run
(the push credential lives on that box; an auto-pulling task would let a leaked
token run code as `SYSTEM` on the host serving `api.mepsinternational.com`).

To ship a change, push it **and then** deploy it:

```bash
ssh -i ~/.ssh/meps_vps_ed25519 Administrator@212.227.127.169
cd C:\DataScienceProject\CanadaQuota
git status && git pull --rebase origin master
.\venv\Scripts\python.exe -m pytest tests -q     # expect 51 passed
```

Do it **before the next Monday run**, or the old code publishes another file
first. See [docs/SERVER_DEPLOYMENT.md](docs/SERVER_DEPLOYMENT.md).

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
| `tools/server-weekly-task.ps1` | **The live entry point.** Task Scheduler runs this weekly on the company server. Publishing is opt-in (`-Push`). |
| `tools/publish_release.py` | Uploads the workbook to the `latest` Release (REST; replaces the GitHub Action). |
| `tools/check_freshness.py` | Reads the newest dated column out of the workbook. Used by the task and the watchdog. |
| `tools/set-github-token.ps1`, `tools/git-askpass.cmd` | Store and supply the push credential without it touching a command line or `.git/config`. |
| `tools/assert-inert.ps1` | Asserts the server copy is inert (or, with `-PostCutover`, correctly live). |
| `.github/workflows/weekly_scrape.yml` | **Schedule disabled 2026-08.** Kept as the emergency fallback via `workflow_dispatch`. |
| `.github/workflows/data-freshness-watchdog.yml` | External heartbeat — alerts if the workbook stops advancing. Replaces GitHub's failure email. |
| `.github/workflows/retry_on_failure.yml` | Auto re-run on a fresh runner when a **dispatched** fallback run fails (IP-block mitigation, capped at 2 retries). |
| `.github/workflows/tests.yml` | Runs the offline test suite on every push/PR. |
| `tests/` | pytest suite: parser (incl. real captured fixtures), writer invariants, discovery. |
| `data/canada_trq_tracker.xlsx` | Output (committed; also published to Release). |
| `enduser/guide.html` | End-user usage guide. |
| `enduser/download_latest.bat` | End-user one-click downloader (pulls the Release). |
| `docs/known-issues.md` | Bug history & prevention — read before editing. |
| `docs/design-2026-03-30.md` | Original design spec (quarter-URL section superseded; see marker inside). |
| `docs/reference/Canadian Quota Template.xlsx` | Laura's reference template (format target only — the script never reads it). |
