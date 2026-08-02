# assert-inert.ps1
#
# Proves what state the server copy of this project is actually in.
#
# "It is not live yet" and "it is live now" are both beliefs. This turns each
# into an assertion that fails loudly. Run it after deploying, after any change
# to the deployment, and any time you are about to do something on the server
# while the other pipeline might still be armed -- two hosts publishing the same
# week race on git push.
#
# Default mode asserts the deployment is INERT (pre-cutover).
# Use -PostCutover to assert the opposite: that it is correctly LIVE.
#
# Exit code 0 = every guard held. Non-zero = the number of guards that failed.
#
# Pure ASCII by design.

param(
    [string]$ProjectRoot = "C:\DataScienceProject\CanadaQuota",
    [string]$TokenFile   = "C:\DataScienceProject\_secrets\canadaquota-github.token",
    [string]$TaskName    = "MEPS Canada TRQ Weekly Update",
    [switch]$PostCutover
)

$ErrorActionPreference = "Continue"
$fail = 0

function Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = "")
    if ($Ok) { "PASS  $Label $Detail" }
    else     { "FAIL  $Label $Detail"; $script:fail++ }
}

$mode = if ($PostCutover) { "LIVE (post-cutover)" } else { "INERT (pre-cutover)" }
"Asserting the deployment is: $mode"
""

# --- the scheduled task ---------------------------------------------------
$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($PostCutover) {
    Check "scheduled task exists" ($null -ne $task)
    if ($task) {
        $trigger = ($task.Triggers | ForEach-Object { $_.StartBoundary }) -join ", "
        $days    = ($task.Triggers | ForEach-Object { $_.DaysOfWeek }) -join ", "
        $principal = $task.Principal.UserId
        Check "task runs as SYSTEM" ($principal -match "SYSTEM") "(is: $principal)"
        Check "task is enabled" ($task.State -ne "Disabled") "(state: $($task.State))"
        "      trigger: $trigger  daysOfWeek: $days"
        $info = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($info) { "      last run: $($info.LastRunTime)  result: $($info.LastTaskResult)  next: $($info.NextRunTime)" }
    }
} else {
    $any = Get-ScheduledTask | Where-Object {
        $_.TaskName -eq $TaskName -or
        ($_.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -match "CanadaQuota"
    }
    Check "no scheduled task references this project" ($null -eq $any)
}

# --- the credential -------------------------------------------------------
$tokenExists = Test-Path $TokenFile
if ($PostCutover) {
    Check "push token is present" $tokenExists
    if ($tokenExists) {
        $bytes = [System.IO.File]::ReadAllBytes($TokenFile)
        $bom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
        Check "token has no BOM" (-not $bom)
        Check "token has no trailing newline" ($bytes[$bytes.Length-1] -ne 10 -and $bytes[$bytes.Length-1] -ne 13)
        $acl = (& icacls.exe $TokenFile) -join " "
        Check "token not readable by Users" ($acl -notmatch "BUILTIN\\Users|Everyone")
    }
} else {
    Check "no push token exists (so -Push cannot succeed)" (-not $tokenExists)
}

# --- the repository -------------------------------------------------------
$origin = & git -C $ProjectRoot remote get-url origin
Check "origin points at GitHub" ($origin -eq "https://github.com/salt0401/Canada-Quota.git") "(is: $origin)"

$dirty = & git -C $ProjectRoot status --porcelain
Check "working tree is clean" ([string]::IsNullOrWhiteSpace($dirty -join "")) "($(($dirty | Measure-Object).Count) changed)"

# The workbook IS the accumulated history; canada_trq_tracker.py refuses to
# rebuild a missing one for exactly that reason.
Check "the history workbook exists" (Test-Path (Join-Path $ProjectRoot "data\canada_trq_tracker.xlsx"))

# --- the script's own guards ---------------------------------------------
$taskScript = Join-Path $ProjectRoot "tools\server-weekly-task.ps1"
$src = Get-Content $taskScript -Raw
Check "task script declares the -Push switch" ($src.Contains('[switch]$Push'))
Check "task script returns early without -Push" ($src.Contains('if (-not $Push)'))
Check "task script carries the UTC/local date guard" ($src.Contains('Local date ($localDate) and UTC date ($utcDate) disagree'))
Check "task script stages the workbook by exact path" ($src.Contains('"add", "--", $DataFile'))

# --- the OTHER pipeline ---------------------------------------------------
# The cutover trap: if the GitHub Actions schedule is still armed, both hosts
# publish the same week and race on push. Assert it from this clone's own copy
# of the workflow. (This reflects the last pull -- confirm on GitHub too.)
$wf = Join-Path $ProjectRoot ".github\workflows\weekly_scrape.yml"
if (Test-Path $wf) {
    $lines = Get-Content $wf
    $activeSchedule = @($lines | Where-Object { $_ -match '^\s*schedule:' -or $_ -match '^\s*-\s*cron:' })
    if ($PostCutover) {
        Check "GitHub Actions weekly schedule is disabled" ($activeSchedule.Count -eq 0) "($($activeSchedule.Count) active line(s))"
    } else {
        "INFO  GitHub Actions weekly schedule active lines: $($activeSchedule.Count) (expected while it is still the live pipeline)"
    }
    Check "workflow_dispatch fallback is retained" (@($lines | Where-Object { $_ -match '^\s*workflow_dispatch:' }).Count -ge 1)
} else {
    Check "weekly_scrape.yml is present in the clone" $false
}

""
if ($fail -eq 0) {
    if ($PostCutover) { "ALL GUARDS ASSERTED - the deployment is correctly LIVE." }
    else              { "ALL GUARDS ASSERTED - the server copy is INERT and cannot publish." }
} else {
    "$fail GUARD(S) FAILED - do not proceed until this is understood."
}
exit $fail
