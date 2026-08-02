# register-scheduled-task.ps1
#
# Arms the weekly run on the MEPS company server. THIS IS THE CUTOVER STEP --
# after it succeeds, this machine is the live pipeline.
#
# Run it ONLY once all of these are true:
#   1. tools\set-github-token.ps1 has stored the push credential,
#   2. an inert run has succeeded on this server,
#   3. the GitHub Actions `schedule:` in .github/workflows/weekly_scrape.yml is
#      commented out AND that change has been pulled onto this clone.
#
# (3) is the cutover trap: two pipelines publishing the same week race on
# `git push`. This script refuses to proceed if it can still see an active
# schedule in the workflow file.
#
#   ssh -i ~/.ssh/meps_vps_ed25519 Administrator@212.227.127.169
#   powershell -ExecutionPolicy Bypass -File C:\DataScienceProject\CanadaQuota\tools\register-scheduled-task.ps1
#
# ROLLBACK (returns the server to inert; GitHub Actions can then be re-armed):
#   Unregister-ScheduledTask -TaskName "MEPS Canada TRQ Weekly Update" -Confirm:$false
#
# WHY MONDAY 13:00 LOCAL. It replaces a GitHub cron of `0 12 * * 1`, i.e. 12:00
# UTC year-round. A 13:00 LOCAL trigger is 12:00 UTC in summer (BST) and 13:00
# UTC in winter (GMT) -- so it is never EARLIER than the run it replaces, in
# either season, and the source data can only ever be fresher. It also lands at
# ~08:00 US Eastern in both seasons, which is what the end-user guide promises,
# and it is hours clear of every existing job on this box: MEPS Currency API
# (06:00), MEPS EU Quota Daily Update (06:40), MEPS SteelNews Scrape (07:15) and
# MEPSR Backup (23:30).
#
# Pure ASCII by design.

param(
    [string]$ProjectRoot = "C:\DataScienceProject\CanadaQuota",
    [string]$TokenFile   = "C:\DataScienceProject\_secrets\canadaquota-github.token",
    [string]$TaskName    = "MEPS Canada TRQ Weekly Update",
    [string]$At          = "13:00",
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$script = Join-Path $ProjectRoot "tools\server-weekly-task.ps1"

"== preflight =="
if (-not (Test-Path $script))    { throw "task script not found: $script" }
if (-not (Test-Path $TokenFile)) {
    throw "Push credential missing: $TokenFile. Run tools\set-github-token.ps1 first -- arming a task that cannot push would produce a weekly failure, not a weekly publish."
}
if (-not (Test-Path (Join-Path $ProjectRoot "venv\Scripts\python.exe"))) {
    throw "venv interpreter missing. Deploy properly before arming a schedule."
}

# The cutover trap, asserted rather than assumed.
$wf = Join-Path $ProjectRoot ".github\workflows\weekly_scrape.yml"
if (Test-Path $wf) {
    $active = @(Get-Content $wf | Where-Object { $_ -match '^\s*schedule:' -or $_ -match '^\s*-\s*cron:' })
    if ($active.Count -gt 0 -and -not $Force) {
        throw "weekly_scrape.yml in this clone still has an ACTIVE schedule ($($active.Count) line(s)). Both pipelines would publish the same week and race on git push. Comment the schedule out, push it, `git pull` here, then re-run. (-Force overrides, deliberately.)"
    }
} else {
    throw "weekly_scrape.yml not found in this clone -- cannot confirm the old pipeline is disabled."
}

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing -and -not $Force) {
    throw "A task named '$TaskName' already exists. Inspect it, or re-run with -Force to replace it."
}

"  script     : $script"
"  token      : present"
"  old cron   : disabled in this clone"
"  server UTC : $([datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss'))"
"  server loc : $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"

"== registering =="
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument ("-NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"{0}`" -Push" -f $script) `
    -WorkingDirectory $ProjectRoot

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At $At

# SYSTEM with a ServiceAccount logon: no stored password, and it matches the
# convention already set by the other two MEPS data projects on this box.
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# StartWhenAvailable: if the box is busy or was off at 13:00, run as soon as it
# can rather than skipping the week. A late run simply heads its column with the
# day it actually ran, which is correct -- one column per week either way.
# IgnoreNew: never let two instances overlap on a 2-physical-core box.
# ExecutionTimeLimit: a run takes ~16 s; an hour is a generous ceiling that
# still stops a hung job living forever.
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 15)

if ($existing) { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false }

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Weekly Canada steel TRQ scrape and publish. Replaces the GitHub Actions job. See C:\DataScienceProject\CanadaQuota\docs\SERVER_DEPLOYMENT.md" | Out-Null

"== registered =="
$t = Get-ScheduledTask -TaskName $TaskName
$i = Get-ScheduledTaskInfo -TaskName $TaskName
"  name    : $($t.TaskName)"
"  state   : $($t.State)"
"  runs as : $($t.Principal.UserId) ($($t.Principal.LogonType))"
"  trigger : $($t.Triggers[0].StartBoundary)  days: $($t.Triggers[0].DaysOfWeek)"
"  next run: $($i.NextRunTime)"
""
"Now assert it: tools\assert-inert.ps1 -PostCutover"
