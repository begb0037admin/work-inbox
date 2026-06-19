$ErrorActionPreference = 'Stop'

Write-Host 'Downloading latest script from GitHub...'
$t = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$scriptPath = Join-Path $PSScriptRoot 'fetch_inbox.py'
Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/fetch_inbox.py?t=$t" -OutFile $scriptPath

Write-Host 'Running inbox briefing...'
python $scriptPath

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nDone. Refresh your dashboard."
} else {
    Write-Host "`nERROR - check output above."
    Read-Host 'Press Enter to close'
}
