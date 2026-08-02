# Server deployment — the weekly run on the MEPS company server

Since August 2026 the weekly Canada TRQ scrape runs on the **MEPS company
server**, not on GitHub Actions. This document is the operational reference for
that: what is installed where, how a run works, and what to do when one fails.

> **Status — 2026-08-02: DEPLOYED AND VERIFIED, NOT YET CUT OVER.**
> The code, interpreter and dependencies are on the server, the test suite passes
> there (**51 passed**, matching the laptop baseline) and a full inert end-to-end
> run succeeded in 16 s. **Nothing is scheduled and nothing can push**: no
> scheduled task references this project, `server-weekly-task.ps1` requires
> `-Push`, and the token file does not exist yet. **GitHub Actions remains the
> live pipeline** until cutover. Update this banner when that changes.

For anything about the server *itself* — access, firewalls, other workloads,
constraints — read the separate **`meps-server-docs`** repository. This file
covers only what is specific to this project.

---

## At a glance

| | |
|---|---|
| **Host** | `WIN-RE1UH50A07U` · `212.227.127.169` · Windows Server 2019 |
| **Location** | `C:\DataScienceProject\CanadaQuota` — sibling of `MEPSWebsScrap` and `EUQuota`, per the server's layout convention |
| **Interpreter** | `C:\DataScienceProject\CanadaQuota\venv\Scripts\python.exe` (Python **3.12.10**) |
| **Scheduled task** | `MEPS Canada TRQ Weekly Update`, **Mondays 13:00 local**, runs as `SYSTEM` |
| **Entry point** | `tools\server-weekly-task.ps1 -Push` |
| **Run log** | `C:\DataScienceProject\CanadaQuota\logs\server_<YYYYMMDD>.log` (180-day retention) |
| **Credential** | `C:\DataScienceProject\_secrets\canadaquota-github.token` — outside every repo |
| **Publishes to** | The same GitHub repo and the same `latest` release as before. **Laura's `download_latest.bat` is unchanged.** |
| **Watched by** | `.github/workflows/data-freshness-watchdog.yml`, running on GitHub |

> ⚠️ **`python` on this server is NOT this project's Python.** Three interpreters
> coexist and bare `python` / `py` both resolve to **3.13.1**. Always use the
> full venv path above. A script relying on bare `python` runs under the wrong
> interpreter, silently.

---

## Why the weekly run moved

The company server is a **standing requirement** for MEPS data projects, not a
cost optimisation. Worth stating plainly because the obvious assumption is
wrong on two counts: this job ran **weekly**, not daily, and
`salt0401/Canada-Quota` is a **public** repository, so its Actions minutes were
unlimited and free. Nothing was being billed. The move is about where MEPS work
is hosted, and it puts this project alongside the steel-news pipeline and the
EU/UK quota tracker under one parent folder.

What did **not** change: the pipeline, the workbook format, the repository, the
`latest` release, and every copy of `download_latest.bat` already on a
colleague's Desktop. Only the machine that runs the scrape is different.

### Why the change of host was not a formality

`docs/known-issues.md` **Bug 3** records that `www.eics-scei.gc.ca` sits behind a
firewall that **silently drops TCP connections from some cloud/datacenter egress
IPs**. The company server is *also* a datacenter address (IONOS). Had that host
refused it, the migration would have been impossible, not merely awkward — so
reachability was tested from the server before any code was written. It passed;
see [Target reachability](#target-reachability-measured-from-this-host-2026-08-02).

---

## What a run does

```
Mon 13:00 local  Task Scheduler
  |
  +-- tools\server-weekly-task.ps1 -Push
       |
       1. Preflight: venv, git working copy, the history workbook, token
       2. Date guard: local date == UTC date, or refuse to publish
       3. Working-tree guard: clean, or nothing but the workbook
       4. venv\Scripts\python.exe canada_trq_tracker.py
       |     -> discovers the current quarter from the landing page
       |     -> downloads both quarter CSVs + the B1 customs table
       |     -> adds one dated column to each TRQ sheet, rewrites TOTALs
       |     -> validates the SAVED workbook, then swaps it in atomically
       5. tools\check_freshness.py --expect-date <today>
       6. git commit + push  (data/canada_trq_tracker.xlsx only)
       7. tools\publish_release.py -> uploads the workbook to the 'latest' release
```

A run takes **~16 seconds** (measured 2026-08-02; the B1 page alone is 6 MB /
18,354 rows). The old GitHub Actions job took ~25 s.

**Step 7 runs after step 6, deliberately** — preserving the order of the
workflow it replaces. The commit is the durable record; the release is the
delivery surface. If the upload fails after a successful push, colleagues keep
receiving last week's file while the repo already holds this week's, which is
recoverable by re-running step 7 alone.

---

## The guards, and why each exists

**1. Publishing is opt-in (`-Push`).** Without the flag the script scrapes and
updates the local workbook but commits nothing, pushes nothing and uploads
nothing. This is what makes a bring-up or debugging run safe. Discard its output
with:

```
git checkout -- data/canada_trq_tracker.xlsx
```

**2. The date guard.** Task Scheduler fires on *local* time, and this server runs
`GMT Standard Time` — UTC in winter, **UTC+1 in summer**. `canada_trq_tracker.py`
heads the new column with `date.today()`, which is local, whereas the GitHub
runner it replaced was always UTC. Between 00:00 and 01:00 local in summer the
two disagree. The designed 13:00 slot is nowhere near that window, so this guard
is really for **manual** runs.

> **If you trigger a run by hand, gate on the server's clock, not your own.**
> During the EU Quota cutover a manual run was fired "just after midnight UTC"
> according to a laptop that turned out to be **6 minutes fast**; the server,
> verified accurate to 3 seconds, was still on the previous UTC day. Check first:
>
> ```powershell
> [datetime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss")
> ```
>
> (`Get-Date -UFormat %s` is not a UTC epoch in PowerShell 5.1 — it derives from
> local time and reads an hour high under BST.)

**3. The working-tree guard.** A dirty `data/canada_trq_tracker.xlsx` is either
the scratch output of an inert run or the residue of a run that died before
committing; either way it is discarded and regenerated from the live source. A
run that reached the commit leaves a **clean** tree, so this can never throw away
something that was published or is about to be. Anything else dirty means
somebody edited the server clone by hand, and the run refuses rather than
publishing it.

> Worth knowing before it surprises you: because this guard runs *first*, a
> manual re-run straight after an inert run **rebuilds** the column instead of
> taking `update_trq_sheet`'s same-day skip path — the tree was dirty, so the
> workbook got reset before the scrape. The skip only appears when the tree is
> clean, i.e. after a successful `-Push` run. Both paths produce the same data;
> only the log differs. (Observed during bring-up, 2026-08-02.)

**4. The freshness assertion.** "The script exited 0" and "the workbook
advanced" are different claims — the whole of `known-issues.md` is about that
gap. `check_freshness.py --expect-date` reads the newest dated column back out
of the saved file, and fails if the update path silently did nothing. It holds
on a fresh run and on a same-day re-run alike, because the script skips a date
column it already has.

**5. No blind conflict resolution.** The task pushes first and only pulls if the
push is rejected. If the rebase then conflicts, it **aborts and fails** rather
than picking a side: the workbook is binary, and `-X theirs` on binary history
would silently discard a week. A rejection is information — most likely both
pipelines are armed.

**6. The script's own fail-loud guards** (`sys.exit(1)`) are unchanged:
discovery/download failure, parse-shape change, a TRQ sheet missing its OVER
header, a missing workbook, and a saved workbook that fails `validate_workbook()`.
Those are covered by `known-issues.md`, not this file.

---

## Updating the code on the server

**This is the biggest behavioural change of the move, and it is easy to get
caught by.** On GitHub Actions, `actions/checkout` fetched the latest `master`
on every run, so pushing a fix was enough. **The server does not do that.** The
clone is whatever was last deployed to it, and the weekly task deliberately does
**not** pull at the start of a run.

That is a security decision, not an oversight. The push credential lives on this
box; if it ever leaked, an auto-pulling task would turn "can write to a public
repo" into "can execute code as `SYSTEM` on the host that serves
`api.mepsinternational.com`". Deployment stays deliberate.

To ship a code change:

```bash
ssh -i ~/.ssh/meps_vps_ed25519 Administrator@212.227.127.169
cd C:\DataScienceProject\CanadaQuota
git status                      # must be clean
git pull --rebase origin master
.\venv\Scripts\python.exe -m pytest tests -q     # expect 51 passed
```

If dependencies changed, also
`.\venv\Scripts\python.exe -m pip install -r requirements.txt -r requirements-dev.txt`.

*(One incidental exception: if a weekly push is rejected, the retry path pulls
before pushing again, which can bring code down as a side effect. It is logged
loudly when it happens.)*

---

## The credential

A **fine-grained** GitHub PAT, scoped to `salt0401/Canada-Quota` only, with
`Contents: Read and write`. Nothing else.

It is stored at `C:\DataScienceProject\_secrets\canadaquota-github.token`,
**outside every git working copy**, readable only by `SYSTEM` and
`Administrators`. It reaches git through `GIT_ASKPASS` at run time, so it never
appears on a command line (visible in the process list on a shared machine) and
never lands in `.git/config`.

Set or rotate it with:

```
ssh -i ~/.ssh/meps_vps_ed25519 Administrator@212.227.127.169
powershell -ExecutionPolicy Bypass -File C:\DataScienceProject\CanadaQuota\tools\set-github-token.ps1
```

**Why fine-grained and single-repo matters here.** This host is internet-facing,
is over a year behind on patches, has SQL Server exposed on 1433, and is backed
up by Acronis to storage MEPS does not control. A credential placed here should
be assumed to be *reachable*. Scoped as above, the worst case is someone writing
to one repository whose entire contents are already public.

**Set an expiry and diary the renewal.** An expired token surfaces as a push
failure, which the watchdog catches within a day.

### Two encoding traps

The token file must be **UTF-8 with no BOM** and have **no trailing newline**.
`git-askpass.cmd` emits it with `type`, so a BOM would be prepended to the
credential and GitHub would reject it. Windows PowerShell 5.1 writes a BOM by
default from both `Set-Content` and `Out-File` — `set-github-token.ps1` uses
.NET directly to avoid this, and verifies the result.

### Git Credential Manager must not be reached

GCM is the default credential helper on Git for Windows and opens a **GUI
prompt** when it has no cached credential. A scheduled task running as `SYSTEM`
has no desktop to show it on, so the push would **hang indefinitely** rather than
fail — and a hung unattended job is strictly worse than a failed one.

The task script therefore passes `-c credential.helper=` on every git
invocation. Note that this cannot be done with `git config credential.helper ""`
from PowerShell: an empty-string element is dropped when an array is splatted to
a native command, which silently turns the write into a *read* and leaves the
manager active.

---

## Monitoring

Nothing on this server notices a scheduled task that never fires. Windows Task
Scheduler has no alerting, and a job that does not start cannot report that it
did not start.

**This is a capability the move removed, not one it never had:** while the
scrape ran on GitHub Actions, GitHub emailed the repository owner whenever the
scheduled workflow failed.

`.github/workflows/data-freshness-watchdog.yml` replaces it. It runs on
**GitHub**, not here — a watchdog hosted on the machine it watches is not a
watchdog — daily at 15:00 UTC, and asserts one fact: how old is the newest dated
column in the committed workbook? That single assertion covers every failure
mode, because the task not firing, discovery failing, the source blocking this
IP, the fail-loud guards aborting and the push failing all end the same way: the
workbook does not advance.

On a healthy pipeline the newest column is 0 days old on Monday afternoon and 6
by Sunday. The limit is **7**, so a missed Monday is tolerated at Monday 15:00
UTC (in case the run was merely late) and alerts on Tuesday. It opens an issue
titled **"Weekly TRQ update has not published"**, comments while the problem
persists, and closes it on recovery.

---

## Triage

```bash
ssh -i ~/.ssh/meps_vps_ed25519 Administrator@212.227.127.169 "Get-ScheduledTaskInfo -TaskName 'MEPS Canada TRQ Weekly Update'"
```

`LastTaskResult` of `0` means the script ran and succeeded. Then read the log:

```bash
ssh -i ~/.ssh/meps_vps_ed25519 Administrator@212.227.127.169 "Get-Content C:\DataScienceProject\CanadaQuota\logs\server_$(date -u +%Y%m%d).log"
```

| Symptom in the log | Cause | Fix |
|---|---|---|
| `venv interpreter not found` | The venv was deleted or the folder moved | Rebuild: `C:\Python312\python.exe -m venv venv` then `pip install -r requirements.txt -r requirements-dev.txt` |
| `is missing. That file IS the history` | The workbook was deleted | Restore with `git checkout -- data/`, or from the `latest` release. **Never** set `TRQ_ALLOW_NEW_WORKBOOK` to work around it |
| `-Push was requested but the token file ... does not exist` | Credential missing or rotated away | Re-run `set-github-token.ps1` |
| `Local date ... and UTC date ... disagree` | The trigger drifted into the pre-01:00 window — **or you triggered it manually from a workstation whose clock is ahead of the server's** | Move the trigger back to 13:00 local. For a manual run, check the server's own clock first |
| `The server clone has changes outside ...` | Somebody hand-edited the working copy | Inspect with `git status` / `git diff` and resolve deliberately |
| `newest column is ..., expected ...` | The scrape ran but the workbook did not advance | Read the Python lines above it in the same log; do not re-run blindly |
| `Could not download TRQ CSV data` / connect timeout | **Bug 3 — the source is refusing this IP** | See below. Use the GitHub Actions fallback to get the week out, then investigate |
| `Push rejected` recurring weekly | Both pipelines are armed — the cutover trap | Confirm `weekly_scrape.yml` has no active `schedule:` |
| `git push failed` (401) | Token expired or lost `Contents: write` | Re-issue the token |
| Task never ran at all | Task disabled, or the box rebooted mid-window | `Get-ScheduledTask`; re-enable |
| Intermittent `WinError 32` on a rename | Antivirus holding a handle — see below | Re-run; if it recurs, request exclusions |

### ⚠️ If the source starts refusing this server's IP

This is the one risk the move **increases**. GitHub rotates runner IPs, so a
blocked draw self-heals on re-run — which is exactly what
`.github/workflows/retry_on_failure.yml` automated. The company server has
**one fixed IONOS address and no such escape hatch.**

The fallback is the workflow that used to own this job:

> GitHub → Actions → **Weekly Canada TRQ Update** → **Run workflow**

That publishes from a GitHub runner exactly as before, on a different egress IP,
and `retry_on_failure.yml` still auto-retries it on a fresh runner if it fails.

⚠️ **Only do this once the server task is confirmed not to have run**, or the two
race on `git push`.

### The antivirus failure signature

Windows Defender real-time scanning **and** Acronis Active Protection both watch
file writes on this host, with **no exclusions** (a standing owner decision —
SQL Server runs here without exclusions too). The predicted symptom is an
**intermittent** `PermissionError` / `WinError 32` on `os.replace`, on a file
this code just wrote, which **succeeds on an immediate re-run**.

`canada_trq_tracker.py` already tolerates this: the atomic swap retries three
times, two seconds apart, and on final failure keeps the new data in
`data/canada_trq_tracker.tmp.xlsx` (gitignored) rather than losing it. If you
see it recur, request a path exclusion for `C:\DataScienceProject` in both
products before suspecting the Python.

---

## Re-deploying from scratch

```powershell
# on a machine with the repo
git bundle create CanadaQuota.bundle --all
# note the sha256, copy it over, verify on the server, then:
git clone CanadaQuota.bundle C:\DataScienceProject\CanadaQuota
cd C:\DataScienceProject\CanadaQuota
git remote set-url origin https://github.com/salt0401/Canada-Quota.git
git config core.autocrlf false
git config user.name  meps-server-canadaquota
git config user.email canadaquota@meps.local
git config credential.https://github.com.username x-access-token
C:\Python312\python.exe -m venv venv
venv\Scripts\python.exe -m pip install -r requirements.txt -r requirements-dev.txt
venv\Scripts\python.exe -m pytest tests -q          # expect 51 passed
```

Use `git bundle`, not a working-tree copy: git performs the checkout so the
server's own rules apply, only tracked content travels, and no GitHub credential
is needed to obtain the code.

**`core.autocrlf false` is not optional.** This repository's blobs have **mixed**
line endings — `canada_trq_tracker.py`, `README.md` and the test fixtures are
stored CRLF, while `docs/known-issues.md`, `enduser/guide.html` and the workflow
YAML are stored LF. Nothing ever enforced a convention. Git for Windows defaults
`core.autocrlf` to `true`, which converts CRLF→LF on `git add`, so the
CRLF-stored files would silently renormalise the moment anything touched their
mtime — leaving the working tree permanently "modified", and the weekly task
refuses to publish from a dirty tree. The repository also ships `* -text` in
`.gitattributes`, which pins each blob's bytes as they are; this setting is the
belt to that braces, because config does not survive a re-clone by someone who
forgets.

> Same trap on the authoring side: use an editor that preserves a file's
> existing endings. A tool that rewrites LF as CRLF turns a paragraph edit into
> a several-hundred-line diff. `git diff --stat` catches it — an edit that
> reports far more changed lines than you wrote is this.

**Install nothing.** Python 3.12.10 (`C:\Python312`) and Git 2.55 were already on
this box from the earlier deployments, and this project pins 3.12. The
deployment adds a folder, a venv, a task and a credential file — no new
software, which is why it needed no installation notice.

Keep the bundle in `C:\DataScienceProject\_installers\` so the deployment is
reproducible.

---

## Deployment record

| | |
|---|---|
| **Deployed** | 2026-08-02 |
| **Commit** | `6767fbb` |
| **Transport** | `git bundle` (517,138 bytes), SHA-256 `9e8e6bc1cd110b44f8b3b916c5e4304a7142c3d5fa72bb7455a0c6c15a2d19f2`, verified at both ends |
| **Interpreter** | Python 3.12.10 (`C:\Python312`, pre-existing, deliberately off PATH) |
| **Dependencies** | `requests 2.34.2`, `beautifulsoup4 4.15.0`, `lxml 6.1.1`, `openpyxl 3.1.5`, `pytest 9.1.1` — all pinned |
| **Test suite on the server** | **51 passed**, identical to the laptop baseline, and the first run of this codebase on 3.12.10 rather than 3.14.2 |
| **PATH** | Verified unchanged after deployment: `python` and `py` still resolve to 3.13.1 |
| **Disk** | 161.8 GB free before; the deployment is ~60 MB including the venv |
| **First inert end-to-end run** | 16 s. Both TRQ sheets updated, B1 table parsed (6.0 MB / 18,354 rows → 8 tracked products, 69 rows), `Workbook validation passed`, output discarded afterwards |

### Windows portability, verified rather than assumed

Two things differ between a Linux GitHub runner and this host, and both feed
straight into the workbook:

| Check | Result on the server |
|---|---|
| `format_date_header` — Windows takes the `%B %#d %Y` branch, Linux `%B %-d %Y` | **Byte-identical** on all six cases tested, including single-digit days (`August 3 2026`) and the exact headers already in the shipped workbook (`March 30 2026`, `July 27 2026`) |
| `LC_TIME` | `(None, None)` — the C locale, so `%B` yields English month names despite the box's `en-GB` culture |
| Default text encoding | **`cp1252`**, not UTF-8. The task script sets `PYTHONUTF8=1` and `PYTHONIOENCODING=utf-8` for exactly this reason: the run logs country names such as `Türkiye`, and a name outside cp1252 would raise `UnicodeEncodeError` from the logging handler and kill a run for a reason unrelated to the data |

The first of those is the one that would have been expensive to discover later:
a mismatched date branch changes the column header format mid-history, in a file
whose whole value is a consistent time series.

### Target reachability, measured from this host 2026-08-02

Doc 00 of `meps-server-docs` is emphatic that "the server has internet" is not
validation, because bot-walls judge per site: from this same IP, `gmk.center`
returns 200 while `mining.com` returns 403. Tested with this project's
production User-Agent, against the URLs it actually fetches:

| Target | TCP:443 | HTTP |
|---|---|---|
| `international.gc.ca` TRQ landing page (`LANDING_URL`, the discovery source) | open | **200** |
| `eics-scei.gc.ca/report-rapport/TRQ_FTA-Y2Q1.csv` | open | **200** |
| `eics-scei.gc.ca/report-rapport/TRQ_NFTA-Y2Q1.csv` | open | **200** |
| `eics-scei.gc.ca/report-rapport/b1.htm` | open | **200** |
| `api.github.com` | open | **200** |
| `pypi.org/simple` | open | **200** |

Confirmed beyond the status code by the inert end-to-end run, which parsed real
data from all four government URLs.

**A pass today is not a permanent guarantee** — Bug 3's block is
reputation-based and IP-specific. Re-test with
`meps-server-docs/scripts/validate-targets.ps1` if scrapes start failing from
this host but succeed elsewhere.
