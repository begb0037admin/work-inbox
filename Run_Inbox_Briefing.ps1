$ErrorActionPreference = 'Stop'

Write-Host 'Downloading latest script from GitHub...'
$t = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$scriptPath = Join-Path $PSScriptRoot 'fetch_inbox.py'
Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/fetch_inbox.py?t=$t" -OutFile $scriptPath

# Connector-only, standing rule (Kevin, 15 Sep 2026): COM and IMAP are retired
# as live options on every machine. Never leave these unset -- fetch_inbox.py
# defaults MAIL_BACKEND/CAL_BACKEND to 'com' when unset, which is exactly the
# fallback Kevin does not want running anywhere, ever again.
$env:MAIL_BACKEND  = 'connector'
$env:CAL_BACKEND   = 'connector'
$env:TEAMS_BACKEND = 'connector'

Write-Host 'Running inbox briefing...'
python $scriptPath

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nDone. Refresh your dashboard."
} else {
    Write-Host "`nERROR - check output above."
    Read-Host 'Press Enter to close'
}
