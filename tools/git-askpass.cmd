@echo off
REM GIT_ASKPASS helper for the server's weekly push.
REM
REM Git invokes this whenever it needs a credential. The remote URL carries the
REM username (x-access-token), so the only thing ever asked for is the password,
REM and the answer is the token in CANADAQUOTA_TOKEN_FILE.
REM
REM Why this exists rather than putting the token in the remote URL or on a
REM command line: arguments are visible in the process list to every account on
REM this shared machine, and a token written into .git/config would be swept up
REM by the Acronis backup along with the rest of the folder.
REM
REM The token file must be UTF-8 with NO BOM and no trailing newline -- `type`
REM emits the bytes verbatim, so a BOM would be prepended to the credential and
REM GitHub would reject it. tools\set-github-token.ps1 writes and verifies that.
type "%CANADAQUOTA_TOKEN_FILE%"
