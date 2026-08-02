# server-weekly-task.ps1
#
# Weekly Canada TRQ scrape and publish, for Windows Task Scheduler on the MEPS
# company server. Replaces .github/workflows/weekly_scrape.yml, which ran this
# on a GitHub runner until cutover. See docs/SERVER_DEPLOYMENT.md.
#
# PURE ASCII BY DESIGN. Windows PowerShell 5.1 reads a BOM-less file as ANSI,
# so a single non-ASCII character breaks string parsing with errors that point
# at unrelated lines. Parse-check before first use --
# meps-server-docs/docs/10-scripting-gotchas.md section 4.
#
# SAFETY: publishing is opt-in. Without -Push this scrapes and updates the
# local workbook, but commits nothing, pushes nothing and uploads nothing.
# That is what makes a bring-up run safe while GitHub Actions is still live --
# two hosts publishing the same week would race on git push.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File server-weekly-task.ps1
#   powershell -ExecutionPolicy Bypass -File server-weekly-task.ps1 -Push

param(
    [switch]$Push,
    [string]$ProjectRoot      = "C:\DataScienceProject\CanadaQuota",
    [string]$TokenFile        = "C:\DataScienceProject\_secrets\canadaquota-github.token",
    [string]$Branch           = "master",
    [int]   $LogRetentionDays = 180
)

$ErrorActionPreference = "Stop"

$DataFile     = "data/canada_trq_tracker.xlsx"
$DataFileWin  = Join-Path $ProjectRoot "data\canada_trq_tracker.xlsx"

# ---------------------------------------------------------------- logging ---
# Logs live in logs\ at the project root, NOT under data\. data\ is the
# committed output directory and the publish step stages from it; a log written
# there is one careless `git add data/` away from being published to a public
# repository every week. logs\ is gitignored.

$logDir = Join-Path $ProjectRoot "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logFile = Join-Path $logDir ("server_{0}.log" -f (Get-Date).ToUniversalTime().ToString("yyyyMMdd"))

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $stamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
    $line = "{0}Z [{1}] {2}" -f $stamp, $Level, $Message
    Write-Output $line
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Fail {
    param([string]$Message)
    Write-Log $Message "ERROR"
    Write-Log "=== RUN FAILED ===" "ERROR"
    exit 1
}

# Judge native commands by exit code, never by stderr. Git, pip and curl all
# write ordinary progress to stderr, and PowerShell 5.1 with
# ErrorActionPreference=Stop treats that as fatal even on exit code 0 --
# meps-server-docs/docs/10-scripting-gotchas.md section 1.
#
# The exit code is published in $script:NativeExit rather than RETURNED, and
# that is not a style choice. Write-Log emits with Write-Output, so anything
# this function logs joins its own pipeline: `$code = Invoke-Native ...` would
# bind an ARRAY of every logged line plus the code, and `if ($code -ne 0)` on an
# array is a filter, not a comparison -- a non-empty result is truthy, so a
# SUCCESSFUL git push would test as failed. Piping to Out-Null instead is no
# better: it silently discards the log lines from the console, which is the
# view a human is watching during a manual run.
function Invoke-Native {
    param([string]$Exe, [string[]]$Arguments, [string]$What, [switch]$AllowFailure)
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    $output = & $Exe @Arguments 2>&1
    $script:NativeExit = $LASTEXITCODE
    $ErrorActionPreference = $prev
    foreach ($line in $output) { Write-Log ("    " + $line) }
    if ($script:NativeExit -ne 0 -and -not $AllowFailure) {
        Fail ("{0} failed with exit code {1}" -f $What, $script:NativeExit)
    }
}

Write-Log "=== Canada TRQ weekly update starting ==="
Write-Log ("Push mode: {0}" -f $(if ($Push) { "LIVE (will commit, push and upload)" } else { "INERT (local only)" }))

# ------------------------------------------------------------- preflight ---

$venvPython = Join-Path $ProjectRoot "venv\Scripts\python.exe"
if (-not (Test-Path $venvPython)) {
    Fail "venv interpreter not found at $venvPython. Bare 'python' on this box is 3.13, not this project's 3.12 -- always use the full path."
}
if (-not (Test-Path (Join-Path $ProjectRoot ".git"))) {
    Fail "$ProjectRoot is not a git working copy."
}
# The workbook is the accumulated history, and canada_trq_tracker.py refuses to
# rebuild a missing one (TRQ_ALLOW_NEW_WORKBOOK) precisely because a silent
# rebuild would publish a history-less file. Catch it here too, with a clearer
# message than a Python traceback in a log nobody is watching.
if (-not (Test-Path $DataFileWin)) {
    Fail "$DataFileWin is missing. That file IS the history. Restore it with 'git checkout -- data/' or from the 'latest' GitHub release; do NOT set TRQ_ALLOW_NEW_WORKBOOK."
}
if ($Push -and -not (Test-Path $TokenFile)) {
    Fail "-Push was requested but the token file $TokenFile does not exist. Run tools\set-github-token.ps1 first."
}

Set-Location $ProjectRoot

# UTF-8 is not optional. The data contains country names like 'Turkiye' (with a
# u-umlaut) and the script logs them; the GitHub runner this replaces was UTF-8
# by default, whereas Windows gives Python the ANSI codepage. A name outside
# cp1252 would raise UnicodeEncodeError from the logging handler and kill the
# run for a reason that has nothing to do with the data.
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

# The timezone trap. This server runs GMT Standard Time -- UTC in winter, UTC+1
# in summer -- and Task Scheduler fires on LOCAL time. canada_trq_tracker.py
# stamps the new column with date.today(), which is LOCAL, whereas the GitHub
# runner it replaces was always UTC. Between 00:00 and 01:00 local in summer the
# two disagree, and the column would be headed a day ahead of the commit that
# carries it.
#
# The designed slot (Monday 13:00 local) is nowhere near that window, so this
# guard exists for MANUAL runs. It blocks publication, not the run: an inert run
# filing a scratch column under the wrong date is discarded anyway.
$utcDate   = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd")
$localDate = (Get-Date).ToString("yyyy-MM-dd")
Write-Log ("Date check -- local {0} / UTC {1}" -f $localDate, $utcDate)
if ($utcDate -ne $localDate) {
    if ($Push) {
        Fail "Local date ($localDate) and UTC date ($utcDate) disagree, and -Push was requested. The trigger has drifted into the pre-01:00 window where local time and UTC fall on different days. Move the trigger back to 13:00 local rather than publishing under an ambiguous date. If you fired this by hand, check the SERVER's clock, not your workstation's: [datetime]::UtcNow"
    }
    Write-Log "Local and UTC dates disagree. Harmless for an inert run -- the column will be headed $localDate and should be discarded afterwards -- but this run could NOT have published." "WARN"
}

# Working-tree state. A dirty workbook is the expected residue of a run that
# died between writing the file and committing it; the data in it was never
# validated by the publish path and the source only serves current values, so
# discarding it costs nothing that is not already lost. Anything ELSE dirty
# means someone hand-edited the server clone, which must not be published.
$dirty = @(& git status --porcelain | Where-Object { $_.Trim() -ne "" })
if ($dirty.Count -gt 0) {
    $others = @($dirty | Where-Object { $_ -notmatch [regex]::Escape($DataFile) })
    if ($others.Count -gt 0) {
        Write-Log "Unexpected local modifications in the server clone:" "ERROR"
        foreach ($d in $others) { Write-Log ("    " + $d) "ERROR" }
        Fail "The server clone has changes outside $DataFile. Someone edited this working copy. Inspect and resolve by hand; refusing to run."
    }
    Write-Log "Discarding leftover uncommitted changes to $DataFile (residue of an earlier failed run; it will be regenerated from the live source)." "WARN"
    Invoke-Native "git" @("checkout", "--", $DataFile) "git checkout data file"
}

# ---------------------------------------------------------------- scrape ---

Write-Log "Running canada_trq_tracker.py ..."
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Invoke-Native $venvPython @("canada_trq_tracker.py") "canada_trq_tracker.py"
$sw.Stop()
Write-Log ("Scrape and workbook update completed in {0:N0}s" -f $sw.Elapsed.TotalSeconds)

# The script exiting 0 and the workbook actually advancing are different claims.
# canada_trq_tracker.py is idempotent -- a second run the same day SKIPS the
# existing column -- so this asserts the newest column is today's, which holds
# both on a fresh run and on a same-day re-run, and fails if the update path
# silently did nothing.
Write-Log "Asserting the workbook now carries a column for $localDate ..."
Invoke-Native $venvPython @("tools\check_freshness.py", "--expect-date", $localDate) "workbook freshness check"

if (-not $Push) {
    Write-Log "INERT run complete. Nothing was committed, pushed or uploaded."
    Write-Log "Discard the local workbook change with: git checkout -- $DataFile"
    Write-Log "=== RUN OK (inert) ==="
    exit 0
}

# ------------------------------------------------------- publish upstream ---

# Identity is set per-repository so nothing global on this shared box changes.
Invoke-Native "git" @("config", "user.name", "meps-server-canadaquota") "git config user.name"
Invoke-Native "git" @("config", "user.email", "canadaquota@meps.local") "git config user.email"

# Credentials reach git through GIT_ASKPASS reading the token file, so the
# token never appears on a command line (visible in the process list on a
# shared machine) and never lands in .git/config. GIT_TERMINAL_PROMPT=0 makes a
# credential problem fail immediately instead of hanging an unattended task.
$env:CANADAQUOTA_TOKEN_FILE = $TokenFile
$env:GIT_ASKPASS = Join-Path $ProjectRoot "tools\git-askpass.cmd"
$env:GIT_TERMINAL_PROMPT = "0"

# Git Credential Manager is the default helper on Git for Windows and opens a
# GUI prompt when it has no cached credential. Under a scheduled task running
# as SYSTEM there is no desktop to show it on, so the task would hang forever
# instead of failing, and a hung unattended job is strictly worse than a failed
# one. "-c credential.helper=" resets the helper list for that one invocation.
#
# It is written as a single "-c credential.helper=" pair rather than
# `git config credential.helper ""` because PowerShell drops an empty-string
# element when splatting an array to a native command -- which silently turns
# the write into a read and leaves the manager active.
$GitNoHelper = @("-c", "credential.helper=")

# Stage the workbook by exact path, not `git add data/`. The workflow this
# replaces used the directory form, which was safe on a throwaway runner but is
# not here: anything that ever lands in data\ on a long-lived machine would be
# published to a public repository.
Invoke-Native "git" @("add", "--", $DataFile) "git add"

Invoke-Native "git" @("diff", "--cached", "--quiet") "git diff --cached" -AllowFailure
if ($script:NativeExit -eq 0) {
    Write-Log "Workbook is byte-identical to the last commit; nothing to commit." "WARN"
} else {
    Invoke-Native "git" @("commit", "-m", "Weekly TRQ update $utcDate") "git commit"

    # Push first, WITHOUT pulling. After cutover this server is the only writer,
    # so a rejection is information: most likely both pipelines are live (the
    # cutover trap) or a human pushed. Only then pull and retry once, so an
    # ordinary docs commit on GitHub does not block the week's data.
    #
    # NOTE: that retry pull brings down remote CODE changes as well as history.
    # It is the ONLY path by which this clone updates itself; the task never
    # pulls at the start of a run, deliberately. See docs/SERVER_DEPLOYMENT.md
    # ("Updating the code on the server") for why, and for the deploy command.
    Invoke-Native "git" ($GitNoHelper + @("push", "origin", $Branch)) "git push" -AllowFailure
    if ($script:NativeExit -ne 0) {
        Write-Log "Push rejected -- the remote has moved. Pulling once and retrying. If this recurs weekly, check that the GitHub Actions schedule is really disabled: two pipelines publishing the same week is the cutover trap." "WARN"
        Invoke-Native "git" ($GitNoHelper + @("pull", "--rebase", "origin", $Branch)) "git pull --rebase" -AllowFailure
        if ($script:NativeExit -ne 0) {
            Invoke-Native "git" @("rebase", "--abort") "git rebase --abort" -AllowFailure
            Fail "Rebase onto origin/$Branch conflicted. The workbook is binary, so this is NOT auto-resolved: picking a side blindly would silently discard a week of history. Resolve by hand over SSH."
        }
        Invoke-Native "git" ($GitNoHelper + @("push", "origin", $Branch)) "git push (retry)"
    }
    Write-Log "Pushed: Weekly TRQ update $utcDate"
}

# The release asset is what download_latest.bat actually fetches, so it is
# uploaded on EVERY run -- including the no-change case above, exactly as the
# workflow step it replaces did. A commit that never reaches the release is
# invisible to the colleague; the release is the delivery surface.
Write-Log "Uploading the workbook to the 'latest' release..."
Invoke-Native $venvPython @("tools\publish_release.py", "--token-file", $TokenFile) "release asset upload"

Remove-Item Env:\CANADAQUOTA_TOKEN_FILE, Env:\GIT_ASKPASS -ErrorAction SilentlyContinue

# ------------------------------------------------------------- retention ---

$cutoff = (Get-Date).AddDays(-$LogRetentionDays)
Get-ChildItem -Path $logDir -Filter "server_*.log" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt $cutoff } |
    ForEach-Object {
        Write-Log ("Pruning old log {0}" -f $_.Name)
        Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
    }

Write-Log "=== RUN OK (live) ==="
exit 0
