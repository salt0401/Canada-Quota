# set-github-token.ps1
#
# One-time setup: store the GitHub push token the weekly task uses, on the
# company server, with the right encoding and the right permissions.
#
# Run this INTERACTIVELY over SSH or RDP. It prompts for the token, so the
# token never appears in a command line, a shell history, a script, or a
# transcript.
#
#   ssh -i ~/.ssh/meps_vps_ed25519 Administrator@212.227.127.169
#   powershell -ExecutionPolicy Bypass -File C:\DataScienceProject\CanadaQuota\tools\set-github-token.ps1
#
# !! LOG IN FIRST, THEN RUN IT. Do NOT pass the script as a remote command:
#
#   ssh ... "powershell -File ...\set-github-token.ps1"     <-- HANGS
#
# `ssh host "command"` allocates no pseudo-terminal, and Read-Host
# -AsSecureString needs a real console to read hidden keystrokes. The prompt
# prints, then the process blocks forever with nowhere to read from. Observed
# 2026-08-02: three attempts left five stuck powershell.exe processes and wrote
# nothing. `ssh -t` forces a PTY and usually works, but an interactive login is
# the reliable route.
#
# -FromStdin is the fallback for when no console is available at all. It reads
# the token from standard input instead of prompting, so it works as a remote
# command or with input redirected from a file:
#
#   ssh ... "powershell -ExecutionPolicy Bypass -File ...\set-github-token.ps1 -FromStdin"
#   <paste the token, press Enter, then Ctrl-D>
#
# The trade-off is that the token is ECHOED in your local terminal, so it lands
# in that terminal's scrollback. It still never reaches a command line or the
# process list, which is the property that matters on a shared machine. Clear
# your scrollback afterwards if that matters to you.
#
# The token must be a FINE-GRAINED personal access token, scoped to
# salt0401/Canada-Quota only, with Repository permissions ->
# "Contents: Read and write". Nothing else. This host is internet-facing, is
# over a year behind on patches and is backed up to storage MEPS does not
# control, so a credential placed here should be assumed to be reachable.
# Scoped as above the worst case is someone writing to one repository whose
# entire contents are already public.
#
# Two encoding details that are easy to get wrong and fail confusingly:
#
#   * NO BOM. git-askpass.cmd emits the file with `type`, so a UTF-8 BOM would
#     be prepended to the token and GitHub would reject it as a bad credential.
#     Set-Content and Out-File both write a BOM by default in Windows
#     PowerShell 5.1, so this uses .NET directly. (The same trap makes sshd
#     silently reject every key in administrators_authorized_keys -- see
#     meps-server-docs/docs/08-openssh-deployment-channel.md.)
#   * NO trailing newline. Written with WriteAllText, not WriteAllLines.
#
# Pure ASCII by design. Windows PowerShell 5.1 reads a BOM-less file as ANSI,
# so one non-ASCII character breaks parsing with errors pointing at unrelated
# lines -- see meps-server-docs/docs/10-scripting-gotchas.md section 4.

param(
    [string]$TokenFile = "C:\DataScienceProject\_secrets\canadaquota-github.token",
    [switch]$FromStdin
)

$ErrorActionPreference = "Stop"

$dir = Split-Path $TokenFile -Parent
if (-not (Test-Path $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    "Created $dir"
}

if ($FromStdin) {
    "Reading the token from standard input (it will be echoed by your terminal)."
    "Paste it, press Enter, then Ctrl-D (Ctrl-Z then Enter from a Windows console)."
    ""
    $token = [Console]::In.ReadToEnd()
    if ($null -ne $token) { $token = $token.Trim() }
} else {
    # Fail fast with a useful message rather than blocking forever. Read-Host
    # -AsSecureString needs a real console; run as a remote ssh command there is
    # none, and the process hangs with the prompt already printed, which looks
    # like the paste did not register.
    if ([Console]::IsInputRedirected) {
        throw "No interactive console: standard input is redirected. You are probably running this as a remote ssh command (ssh host `"powershell -File ...`"), which allocates no pseudo-terminal. Log in first and run it at the remote prompt, or re-run with -FromStdin. See this script's header."
    }
    ""
    "Paste the fine-grained GitHub token (scoped to salt0401/Canada-Quota,"
    "Repository permissions -> Contents: Read and write), then press Enter."
    "Input is hidden."
    ""
    $secure = Read-Host -AsSecureString "Token"
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        $token = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr).Trim()
    } finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

if (-not $token) { throw "No token entered." }
if ($token.Length -lt 20) { throw "That does not look like a GitHub token (too short)." }

# UTF-8 with NO BOM, and no trailing newline.
[System.IO.File]::WriteAllText($TokenFile, $token, (New-Object System.Text.UTF8Encoding($false)))

# Readable only by the account the scheduled task runs as, plus admins.
# /inheritance:r drops the inherited grants first, or Users would still read it.
& icacls.exe $TokenFile /inheritance:r /grant "SYSTEM:R" /grant "Administrators:F" | Out-Null
if ($LASTEXITCODE -ne 0) { throw "icacls failed with exit code $LASTEXITCODE" }

# Verify what landed WITHOUT printing the token.
$bytes = [System.IO.File]::ReadAllBytes($TokenFile)
$hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
$endsClean = ($bytes[$bytes.Length - 1] -ne 10 -and $bytes[$bytes.Length - 1] -ne 13)
$sha = (Get-FileHash -Path $TokenFile -Algorithm SHA256).Hash.Substring(0, 12)

""
"Stored: $TokenFile"
"  bytes           : $($bytes.Length)"
"  BOM             : $hasBom   (must be False)"
"  ends cleanly    : $endsClean   (must be True - no trailing newline)"
"  sha256 prefix   : $sha   (fingerprint only; safe to quote when confirming)"
""
"Permissions:"
& icacls.exe $TokenFile | ForEach-Object { "  $_" }
""
if ($hasBom -or -not $endsClean) {
    throw "The token file is malformed. Re-run this script."
}
"Token stored correctly."
""
"Next: prove the credential works WITHOUT publishing anything -"
"  cd C:\DataScienceProject\CanadaQuota"
"  `$env:CANADAQUOTA_TOKEN_FILE = '$TokenFile'"
"  `$env:GIT_ASKPASS = 'C:\DataScienceProject\CanadaQuota\tools\git-askpass.cmd'"
"  git -c credential.helper= ls-remote origin master"
"then register the scheduled task."
