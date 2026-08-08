# Session handover — Canada TRQ Tracker

**There is exactly one of these files, and it is overwritten.** It is not a log
archive. When you finish a working session, replace its contents with the state
you are leaving behind; do not add `SESSION_HANDOVER_2.md` or a dated copy. The
value of this file is that it is the *current* answer, and a folder of stale
session notes is worse than none.

Written 2026-08-08, from the laptop session that moved this project off GitHub
Actions and onto the MEPS company server.

---

## If you are a fresh session running ON the server, read this part

The previous work was done **from a laptop over SSH**. You are running **on the
box itself**, and that changes one thing more than any other:

> ### ⚠ `C:\DataScienceProject\CanadaQuota` IS PRODUCTION. It is not a checkout.
>
> It is simultaneously the git working copy you would edit and the deployment
> the scheduled task runs from. There is no separate build or release step.
>
> **`tools\server-weekly-task.ps1` refuses to run if the working tree is dirty
> outside `data/`.** That guard is deliberate — it stops a hand-edited clone
> being published to a public repository unattended. But it means that **any
> uncommitted edit you leave behind on a Sunday night stops Monday's run.** The
> job fails, the week's data does not publish, and the watchdog raises an issue
> on Tuesday.
>
> **Do not weaken that guard to make development comfortable.** Do one of:
>
> 1. **Recommended — develop in a second clone.** Make
>    `C:\DataScienceProject\_work\CanadaQuota`, work there, push to GitHub, then
>    deploy to the production clone with `git pull`. Production stays clean by
>    construction, and you get to test before the live copy sees your change.
> 2. If you do edit in place, **commit or `git stash` before you stop**, and
>    check `git status` is clean.

Everything else about the environment is documented in the **`meps-server-docs`**
repository (private). If it is not cloned on the server yet, clone it — it is the
reference for the machine itself, and this project's docs deliberately do not
duplicate it.

---

## Current state — verified, with evidence

| | |
|---|---|
| **Status** | ✅ **LIVE on the company server since 2026-08-02.** GitHub Actions no longer runs it |
| Location | `C:\DataScienceProject\CanadaQuota` — sibling of `MEPSWebsScrap`, `EUQuota`, `ConferenceMonitoring`, `StainlessAlloySurcharges` |
| Interpreter | `venv\Scripts\python.exe` → Python **3.12.10** |
| Task | `MEPS Canada TRQ Weekly Update`, **Mondays 13:00 local**, as `SYSTEM` |
| Entry point | `tools\server-weekly-task.ps1 -Push` |
| Logs | `logs\server_<YYYYMMDD>.log` (180-day retention, gitignored) |
| Credential | `C:\DataScienceProject\_secrets\canadaquota-github.token` — **no expiry, by owner decision** |
| Last verified commit | `497fd10` |

**The migration is proven on a real unattended run, not just a manual one.**
Commit `497fd10 Weekly TRQ update 2026-08-03`, authored by
`meps-server-canadaquota <canadaquota@meps.local>` at **2026-08-03 13:00:30
+0100** — the scheduled task firing on time and publishing by itself. A commit
from that identity only exists at the end of a successful `-Push` run, so its
presence is the evidence.

Other things confirmed by measurement rather than assumption:

- **51 tests pass on the server** (Python 3.12.10), matching the laptop baseline.
- **All four government source URLs return 200 from the server's IONOS IP.** This
  was the gating question for the whole migration — see `known-issues.md` Bug 3,
  which records that host silently dropping datacenter egress addresses.
- **Windows produces byte-identical date-column headers** to the Linux runner it
  replaced (`%B %#d %Y` vs `%B %-d %Y`), verified on six cases including
  single-digit days.
- **The freshness watchdog's schedule works.** Runs #2–#7 all fired on `schedule`
  between 2026-08-02 and 2026-08-07. Note its *first* slot was silently skipped,
  and delivery runs **46–125 minutes late** against the 15:00 UTC cron — normal
  for GitHub's best-effort scheduler, but do not expect punctuality.
- **The delivery chain was verified from outside**, by downloading
  `releases/latest/download/canada_trq_tracker.xlsx` — the exact URL the
  colleague's `.bat` uses — and confirming it carried the new column while the
  closed-quarter sheets kept all 15 of theirs.

---

## Open items

| Item | State |
|---|---|
| **Watchdog run #6 was `cancelled`** (2026-08-06), not failed | Job never started; #7 succeeded the next day. Cause not established — most likely GitHub infrastructure, possibly the `concurrency` group. **The real point: a cancelled run is a silent gap in the heartbeat.** Nothing checked that day and nothing said so. Worth a look if it recurs |
| **SSH from the previous laptop times out** | **Already tracked in `meps-server-docs`, not a new finding** — recorded there as broken since 2026-08-05, re-verified down 2026-08-07, same egress `31.14.249.144`. A *timeout* means dropped upstream, i.e. the source IP is no longer permitted. The agreed fix is **Tailscale** (approved 2026-08-07; `scripts/install-tailscale.ps1` in that repo), which retires the IONOS source-IP allowlist for maintenance access. **Irrelevant while you work on the box itself**, but it is why this handover ends with git evidence rather than a live SSH check |
| **Push token has no expiry** | Owner's decision. Nothing will prompt a renewal. It is only as long-lived as the issuing GitHub account, so it dies when that account loses access — surfacing as a **401 on push**. The watchdog is the only thing that will notice |
| **EU Quota has the same `Invoke-Native` bug this project had** | Confirmed firing on every run of that live daily job (`fatal: no rebase in progress` in its log). Not fixed — different project, deliberately left alone. The fix pattern is `tools/server-weekly-task.ps1` at commit `56104ca` |
| **Colleagues' feature requests** | Untouched by design. The owner's instruction was to move the project *completely and unchanged* first, then improve it in the new environment. That precondition is now met |

---

## Things that will cost you a cycle if you rediscover them

1. **`python` and `py` on this box are 3.13, not this project's 3.12.** Three
   interpreters coexist. Always call `venv\Scripts\python.exe` by full path. A
   script relying on bare `python` runs under the wrong interpreter, silently.
2. **Line endings are mixed in this repo** — some blobs CRLF, some LF.
   `.gitattributes` pins `* -text` and the clone sets `core.autocrlf false`, so
   nothing renormalises. **Use an editor that preserves a file's existing
   endings.** A tool that rewrites LF as CRLF turns a three-paragraph edit into a
   627-line diff; that happened during this migration. `git diff --stat` catches
   it — a diff far larger than what you wrote is line endings, not your edit.
3. **PowerShell deployment scripts must be pure ASCII.** Windows PowerShell 5.1
   reads a BOM-less file as ANSI, so one non-ASCII character breaks parsing with
   errors pointing at unrelated lines. Parse-check before running:
   `[System.Management.Automation.Language.Parser]::ParseFile(...)`.
4. **Never capture an `Invoke-Native`-style helper's exit code by assignment if
   the helper also logs with `Write-Output`.** You bind an *array* of log lines
   plus the code, and `-ne 0` on an array is a filter, not a comparison — a
   non-empty result is truthy, so a **successful** `git push` tests as rejected.
   This project publishes the code in `$script:NativeExit` instead. The bug only
   affects commands that write output; `git diff --cached --quiet` is silent and
   was always fine.
5. **Windows gives Python `cp1252`, not UTF-8.** The task script sets
   `PYTHONUTF8=1` and `PYTHONIOENCODING=utf-8` because the run logs country names
   like `Türkiye`. Do not remove them.
6. **The workbook IS the history.** `canada_trq_tracker.py` refuses to rebuild a
   missing one. If you see the `TRQ_ALLOW_NEW_WORKBOOK` escape hatch, do not
   reach for it — restore from git or the release instead.

---

## What this migration changed, and what it deliberately did not

**Unchanged:** the scraper, the parser, the Excel writer, the workbook format,
the repository, the `latest` release, and every copy of `download_latest.bat`
already on a colleague's desktop. The pipeline was moved, not modified — that was
the explicit instruction, so that any failure would be attributable to the change
of host and nothing else.

**Changed, and worth knowing:**

- **A pushed fix no longer deploys itself.** GitHub Actions checked out `master`
  every run. The server runs what was last pulled onto it, and the weekly task
  deliberately does **not** pull at the start of a run — the push credential
  lives on that box, and an auto-pulling task would turn a leaked token into code
  execution as `SYSTEM` on the host serving `api.mepsinternational.com`.
- **GitHub's failure email is gone**, because GitHub no longer runs the job.
  `.github/workflows/data-freshness-watchdog.yml` replaces it.
- **The Bug 3 escape hatch is gone.** GitHub rotated runner IPs so a blocked draw
  self-healed; the server has one fixed address. `weekly_scrape.yml` is kept with
  its schedule commented out precisely so `workflow_dispatch` can draw a
  different IP if the source ever refuses the server. **Never dispatch it while
  the server task is armed and has not already failed for that week** — the two
  race on `git push`, and the task refuses to auto-resolve a binary conflict.

---

## Where the authoritative documents are

| Question | Read |
|---|---|
| How the server run works, how to triage it, how to redeploy | `docs/SERVER_DEPLOYMENT.md` |
| Bugs that have already bitten this project | `docs/known-issues.md` — **read before touching the Excel writer or the B1 scraper** |
| What the program does and produces | `README.md` |
| The server itself — access, firewalls, other workloads, constraints | the separate **`meps-server-docs`** repo |
| What actually changed and why | `git log` — it is the source of truth, and the commit messages carry the reasoning |

---

## Suggested next step

The move is complete and the environment is stable, so the precondition the owner
set ("move completely and perfectly first, then modify") is satisfied. The
colleagues' requested improvements are the natural next piece of work.

Before starting any of them, confirm the pipeline is still healthy:

```powershell
cd C:\DataScienceProject\CanadaQuota
git status                                        # must be clean
git log --oneline -3                              # newest should be a recent "Weekly TRQ update"
.\venv\Scripts\python.exe -m pytest tests -q      # expect 51 passed
.\venv\Scripts\python.exe tools\check_freshness.py
Get-ScheduledTaskInfo -TaskName "MEPS Canada TRQ Weekly Update"
```

`LastTaskResult` of `0` and a newest column within the last 7 days means
everything is working.
