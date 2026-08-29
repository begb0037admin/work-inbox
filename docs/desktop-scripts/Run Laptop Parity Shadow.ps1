# Run Laptop Parity Shadow.ps1
# work-inbox laptop migration, Phase 4 -- the parallel/shadow run.
#
# Runs parity_vs_briefing.py on the weekday cadence: a fresh IMAP mail capture
# (MAIL_BACKEND=imap WI_MAIL_PARALLEL=1, which exits after the raw dump -- NO
# calendar, NO AI, NO push) + a diff against the live desktop briefing.json.
#
# Writes ONLY:
#   %USERPROFILE%\work-inbox\data\parallel\imap_inbox_raw.json / imap_sent_raw.json
#   %USERPROFILE%\work-inbox\data\parallel\parity_vs_briefing_<ts>.json
#   %USERPROFILE%\work-inbox\logs\parity_shadow.log   (appended, timestamped)
# Never pushes, never touches data\briefing.json, never opens classic Outlook.
#
# Live copy lives at  %USERPROFILE%\work-inbox\Run Laptop Parity Shadow.ps1
# (this docs/desktop-scripts/ copy is the reference). Registered by
# Register-LaptopParityShadow.ps1.

$ErrorActionPreference = 'Continue'
$root = Join-Path $env:USERPROFILE 'work-inbox'
$logdir = Join-Path $root 'logs'
New-Item -ItemType Directory -Force -Path $logdir | Out-Null
$log = Join-Path $logdir 'parity_shadow.log'

function Log($m) { "$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))  $m" | Tee-Object -FilePath $log -Append }

Log "=== Run Laptop Parity Shadow start (user $env:USERDOMAIN\$env:USERNAME) ==="
Set-Location $root

# subscription billing hygiene even though the parity script does not call claude
$env:ANTHROPIC_API_KEY = $null

# always run the current scripts from main (cache-busted raw pull)
$t = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
foreach ($f in 'fetch_inbox.py','imap_mail.py','diff_mail_pull.py','parity_vs_briefing.py') {
  try {
    Invoke-WebRequest -UseBasicParsing "https://raw.githubusercontent.com/begb0037admin/work-inbox/main/$f`?t=$t" -OutFile (Join-Path $root $f)
  } catch {
    Log "WARN: could not refresh $f ($($_.Exception.Message)) -- using the local copy"
  }
}

Log "running: python parity_vs_briefing.py"
& python (Join-Path $root 'parity_vs_briefing.py') *>> $log
$rc = $LASTEXITCODE
Log "parity_vs_briefing.py exit $rc  (0 = 0 real flags, 1 = real flags to review, 2 = setup/fetch issue)"
Log "=== Run Laptop Parity Shadow end ==="
exit 0   # a shadow run must never surface a Task Scheduler failure
