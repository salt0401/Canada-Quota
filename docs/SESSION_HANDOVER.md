# Session handover — Canada TRQ Tracker

**There is exactly one of these files, and it is overwritten.** It is not a log
archive. When you finish a working session, replace its contents with the state
you are leaving behind; do not add `SESSION_HANDOVER_2.md` or a dated copy. The
value of this file is that it is the *current* answer, and a folder of stale
session notes is worse than none.

Written 2026-10-20, from a session running on the server itself, which fixed
the four-week outage described below (known-issues Bug 5).

---

## If you are a fresh session running ON the server, read this part

> ### ⚠ `C:\DataScienceProject\CanadaQuota` IS PRODUCTION. It is not a checkout.
>
> It is simultaneously the git working copy you would edit and the deployment
> the scheduled task runs from. There is no separate build or release step.
>
> **`tools\server-weekly-task.ps1` refuses to run if the working tree is dirty
> outside `data/`.** That guard is deliberate — it stops a hand-edited clone
> being published to a public repository unattended. But it means that **any
> uncommitted edit you leave behind on a Sunday night stops Monday's run.**
>
> **Do not weaken that guard to make development comfortable.** Develop in the
> second clone instead: **`C:\DataScienceProject\_work\CanadaQuota`** (created
> 2026-10-20). Its `origin` is the production folder, not GitHub, and it has no
> `venv` of its own. Run tests with production's interpreter by full path:
> `C:\DataScienceProject\CanadaQuota\venv\Scripts\python.exe -m pytest tests -q`.
> Deploy a commit made there with
> `git -C C:\DataScienceProject\CanadaQuota pull --ff-only C:\DataScienceProject\_work\CanadaQuota master`.

**Anything that reaches production's `master` goes public on the next `-Push`
run**, because the task ends with `git push origin master`. Deploying and
publishing to GitHub are the same act here.

Everything else about the environment is documented in the **`meps-server-docs`**
repository, cloned at `C:\DataScienceProject\meps-server-docs`. This project's
docs deliberately do not duplicate it.

---

## Current state

| | |
|---|---|
| **Status** | ⚠ **Weekly runs failed 2026-09-28, 10-05, 10-12, 10-19** (Bug 5). Fix verified in the work clone 2026-10-20 |
| **Is the fix live?** | Only once production's `git log` contains the commit that added Bug 5 to `known-issues.md`. Check `git -C C:\DataScienceProject\CanadaQuota log --oneline -5` |
| Location | `C:\DataScienceProject\CanadaQuota` |
| Interpreter | `venv\Scripts\python.exe` → Python **3.12.10** |
| Task | `MEPS Canada TRQ Weekly Update`, **Mondays 13:00 local**, as `SYSTEM` |
| Entry point | `tools\server-weekly-task.ps1 -Push` |
| Logs | `logs\server_<YYYYMMDD>.log` (180-day retention, gitignored) |
| Credential | `C:\DataScienceProject\_secrets\canadaquota-github.token` — **no expiry, by owner decision** |
| Last good unattended run | `f666e2c Weekly TRQ update 2026-09-21` |

**Evidence for the fix, gathered 2026-10-20 in the work clone:**

- `pytest tests -q` → **53 passed** (the 51 pre-existing tests, golden-fixture
  regression included, plus 2 new Bug 5 tests). Both new tests failed on the
  unfixed parser with the production error before the fix.
- The live `TRQ_FTA-Y2Q2`, `TRQ_NFTA-Y2Q2` and `TRQ_FTA-Y2Q1` CSVs parse to all 8
  tracked products, with no Part A/B mismatch warnings. Country shares sum to each
  product's total within 0.02 percentage points.
- A full `canada_trq_tracker.py` run exited 0. It created `non-FTA Q2` / `FTA Q2`,
  archived `B1 Imports Q1`, and passed `validate_workbook`. Spot-checked cells
  match the raw CSV (FTA Q2 Hot-Rolled Sheet: Japan 0.3998, Korea 0.5738). That
  test-run workbook was **not** committed: it is in the work clone's
  `git stash`, and the first real run produces the same thing.

**Not yet observed:** an unattended `-Push` run with the fix. The proof will
be a `Weekly TRQ update <date>` commit by `meps-server-canadaquota` after
`f666e2c`.

---

## Open items

| Item | State |
|---|---|
| **Four missing weekly columns** | 2026-09-28 (last week of Y2Q1) and Q2's 10-05, 10-12 and 10-19. The source serves only current data, so they cannot be re-fetched. Q2 sheets will start at the first successful run. This session does not know whether Laura's side has been told about the gap |
| **Did anything alert during the outage?** | Unverified. The newest column is 2026-09-21, so against its 7-day limit the GitHub freshness watchdog should have opened an issue from about 2026-09-29. This session had no GitHub API access to check. If no issue exists, the watchdog is broken, and that is a bigger problem than Bug 5. The server's own health check (`C:\DataScienceProject\_monitoring\status.txt`) did report the failure |
| **Push token has no expiry** | Owner's decision. It dies with the issuing GitHub account's access, surfacing as a **401 on push**. Nothing else will notice |
| **Sheet names are year-blind** | known-issues Bug 4's documented limitation. Must get a year suffix before the next program-year rollover (Year-3 Q1, late June 2027) |
| **Colleagues' feature requests** | Still untouched |

---

## Things that will cost you a cycle if you rediscover them

1. **`python` and `py` on this box are 3.13, not this project's 3.12.** Three
   interpreters coexist. Always call `venv\Scripts\python.exe` by full path.
2. **Line endings are mixed in this repo** — some files CRLF (`README.md`, the
   live-captured fixture), most LF. `.gitattributes` pins `* -text` and the
   clones set `core.autocrlf false`, so nothing renormalises. A diff far larger
   than what you wrote is line endings. `git diff --stat` catches it.
3. **PowerShell deployment scripts must be pure ASCII.** Windows PowerShell 5.1
   reads a BOM-less file as ANSI. Parse-check before running:
   `[System.Management.Automation.Language.Parser]::ParseFile(...)`.
4. **Never capture an `Invoke-Native`-style helper's exit code by assignment if
   the helper also logs with `Write-Output`.** You bind an array, and `-ne 0` on
   an array is a filter. This project publishes the code in `$script:NativeExit`.
5. **Windows gives Python `cp1252`, not UTF-8.** Set `PYTHONUTF8=1` and
   `PYTHONIOENCODING=utf-8` for any manual run (the task script already does).
   Country names like `Türkiye` break otherwise.
6. **The workbook IS the history.** `canada_trq_tracker.py` refuses to rebuild a
   missing one. Do not reach for `TRQ_ALLOW_NEW_WORKBOOK`. Restore from git or
   the release instead.
7. **The CSV is parsed by records, not lines** (Bug 5). If you touch
   `parse_trq_csv()`, keep it that way: the title row now contains quoted line
   breaks.

---

## Where the authoritative documents are

| Question | Read |
|---|---|
| How the server run works, how to triage it, how to redeploy | `docs/SERVER_DEPLOYMENT.md` |
| Bugs that have already bitten this project | `docs/known-issues.md` — **read before touching the parser, the Excel writer or the B1 scraper** |
| What the program does and produces | `README.md` |
| The server itself | `meps-server-docs` |
| What actually changed and why | `git log` |

---

## Suggested next step

If the fix is not yet in production's `git log`, deploying it before the next
Monday 13:00 run is the whole job. Then confirm the pipeline:

```powershell
cd C:\DataScienceProject\CanadaQuota
git status                                        # must be clean
git log --oneline -3                              # newest should be a recent "Weekly TRQ update"
.\venv\Scripts\python.exe -m pytest tests -q      # expect 53 passed
.\venv\Scripts\python.exe tools\check_freshness.py
Get-ScheduledTaskInfo -TaskName "MEPS Canada TRQ Weekly Update"
```

`LastTaskResult` of `0` and a newest column within the last 7 days means
everything is working.
