<#
Run Laptop Bridge Briefing.ps1
==============================
work-inbox LAPTOP BRIDGE -- REAL mail-only briefings from Kevin's Oxford laptop
(101L-DE013193 / begb0037.AD-OAK, user ad-oak\begb0037) for as long as the admin
DESKTOP's M365 device-registration is broken (0x8004dec5) and Outlook COM there
is dead. Kevin + Max are repairing the desktop separately.

This is NOT the Phase 4 parity shadow ("Run Laptop Parity Shadow.ps1"). This one
pushes for real:
    IMAP+OAuth2 mail pull  ->  claude -p triage (kevin@ isolated config)
    ->  Phase 4  data/briefing.json  -> GitHub
    ->  Phase 5  command-centre task-suggestion sync
    ->  (best effort) needs_reply.json + drafted_replies.json publishers

Calendar: NONE. CAL_BACKEND=com by default, but there is no classic Outlook on
the laptop, so fetch_inbox.py degrades the calendar phases to empty + a warning
(handled path, not a crash). The bridge briefing simply has no calendar section.
Accepted for the bridge. Pass -CalBackend connector to skip the COM calendar
attempt entirely (identical empty-calendar result, never touches COM).

Mirrors the live desktop "Run Inbox Briefing.bat" environment, minus Outlook COM
and minus the hope@ overflow config (single account on the laptop for now -- a
Pro-cap hit degrades that one run; acceptable for a short bridge).

PARAMS
  -CoreOnly          run only fetch_inbox.py (skip the two downstream publishers).
                     Use this for the first supervised run.
  -CalBackend        com (default) | connector

LIVE COPY   %USERPROFILE%\work-inbox\Run Laptop Bridge Briefing.ps1
REFERENCE   work-inbox/docs/desktop-scripts/Run Laptop Bridge Briefing.ps1
REGISTERED  Register-LaptopBridgeBriefing.ps1  ->  task "Work Inbox Bridge Briefing"

REVERT / END OF BRIDGE
  1. Unregister-ScheduledTask -TaskName 'Work Inbox Bridge Briefing' -Confirm:$false   (on the laptop)
  2. Enable-ScheduledTask   -TaskName 'Work Inbox Briefing'                            (on the admin desktop)
#>
param(
  [switch]$CoreOnly,
  [ValidateSet('com','connector')] [string]$CalBackend = 'com'
)

$ErrorActionPreference = 'Continue'
$root   = Join-Path $env:USERPROFILE 'work-inbox'
$tools  = Join-Path $root 'tools'
$logdir = Join-Path $root 'logs'
New-Item -ItemType Directory -Force -Path $root, $tools, $logdir | Out-Null

$stamp  = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')
$log    = Join-Path $logdir "bridge_briefing_$stamp.log"
$latest = Join-Path $logdir 'bridge_briefing_last_run.log'

function Log($m) {
  $line = "$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))  $m"
  Write-Host $line
  Add-Content -LiteralPath $log -Value $line
}

Log "=== Laptop Bridge Briefing START  (user $env:USERDOMAIN\$env:USERNAME  host $env:COMPUTERNAME) ==="
Log "params: CoreOnly=$CoreOnly  CalBackend=$CalBackend  log=$log"
Set-Location $root

# --- isolated Claude Code config: kevin@ ONLY (no hope@ failover on the laptop yet) ---
$kevinCfg = 'C:\WorkInboxAI\kevin'
if (-not (Test-Path (Join-Path $kevinCfg '.credentials.json'))) {
  Log "FATAL: $kevinCfg\.credentials.json not found -- the kevin@ isolated Claude Code config is not logged in."
  Log "FIX:   `$env:CLAUDE_CONFIG_DIR='$kevinCfg'; claude /login   (sign in as kevin@lelitte.co.uk), then re-run."
  Copy-Item $log $latest -Force
  exit 3
}

$env:AI_BACKEND                    = 'claude_code'
$env:ANTHROPIC_API_KEY            = ''            # force subscription billing (matches desktop .bat)
$env:WI_CLAUDE_CONFIG_DIR         = $kevinCfg     # -> claude -p gets CLAUDE_CONFIG_DIR=C:\WorkInboxAI\kevin
$env:WI_CLAUDE_CONFIG_DIR_FALLBACK = ''           # explicit: single account, no hope@ overflow
$env:MAIL_BACKEND                 = 'imap'
$env:CAL_BACKEND                  = $CalBackend
$env:WI_MAIL_PARALLEL            = ''             # explicit: this is a REAL run, not a parallel capture
$env:PYTHONUTF8                  = '1'

# --- refresh pipeline scripts from main (cache-busted raw pull, same mechanism the desktop uses) ---
$t    = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$base = 'https://raw.githubusercontent.com/begb0037admin/work-inbox/main'
foreach ($f in 'fetch_inbox.py','imap_mail.py','reauth_imap.py') {
  try {
    Invoke-WebRequest -UseBasicParsing "$base/$f`?t=$t" -OutFile (Join-Path $root $f)
    Log "refreshed $f from main"
  } catch {
    Log "WARN: could not refresh $f ($($_.Exception.Message)) -- using the local copy"
  }
}

# --- CORE: fetch_inbox.py  (Phase 1 IMAP -> combined claude -p triage -> Phase 4 push -> Phase 5 CC sync) ---
Log "running: python -u fetch_inbox.py   [MAIL_BACKEND=imap  CAL_BACKEND=$CalBackend  AI_BACKEND=claude_code  cfg=$kevinCfg]"
& python -u (Join-Path $root 'fetch_inbox.py') 2>&1 | Tee-Object -FilePath $log -Append
$rc = $LASTEXITCODE
Log "fetch_inbox.py exit $rc"

if ($rc -ne 0) {
  Log "CORE FAILED (exit $rc). NOT running publishers."
  if ($rc -eq 1) { Log "exit 1 is most likely an expired IMAP token -> run 'Re-auth Work Inbox IMAP.bat' (one browser click), then re-run this." }
  Copy-Item $log $latest -Force
  Log "=== Laptop Bridge Briefing END (core failed) ==="
  exit $rc
}

if ($CoreOnly) {
  Log "CoreOnly set -- skipping the needs_reply / drafted_replies publishers."
  Copy-Item $log $latest -Force
  Log "=== Laptop Bridge Briefing END (core OK, CoreOnly) ==="
  exit 0
}

# --- BEST-EFFORT downstream publishers. A failure here NEVER fails the run: the
#     briefing itself already succeeded and is the primary deliverable. ---
function Get-PipelineScript($url, $dest, $marker) {
  try {
    $tt = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    Invoke-WebRequest -UseBasicParsing "$url`?t=$tt" -OutFile "$dest.download"
    if ((Get-Item "$dest.download").Length -lt 400) { throw 'downloaded file too small' }
    if ($marker -and -not (Select-String -Quiet -LiteralPath "$dest.download" -Pattern $marker)) {
      throw "downloaded file missing marker /$marker/"
    }
    Move-Item -Force "$dest.download" $dest
    return $true
  } catch {
    Log "WARN: could not fetch $(Split-Path $dest -Leaf) ($($_.Exception.Message)) -- publisher will be skipped"
    if (Test-Path "$dest.download") { Remove-Item "$dest.download" -Force }
    return $false
  }
}

$tb = "$base/tools"
$okDeps = (Get-PipelineScript "$tb/style_corpus_common.py"  (Join-Path $tools 'style_corpus_common.py')  '^def recipient_tier') `
      -and (Get-PipelineScript "$tb/phase_failure_notify.py" (Join-Path $tools 'phase_failure_notify.py') '^def notify_phase_failure')

if ($okDeps -and (Get-PipelineScript "$tb/publish_needs_reply.py" (Join-Path $tools 'publish_needs_reply.py') '^def run\(token')) {
  Log "running: python -u tools\publish_needs_reply.py"
  Push-Location $tools
  & python -u (Join-Path $tools 'publish_needs_reply.py') 2>&1 | Tee-Object -FilePath $log -Append
  Log "publish_needs_reply.py exit $LASTEXITCODE"
  Pop-Location
} else {
  Log "SKIP publish_needs_reply.py (dependency/download issue) -- non-fatal"
}

if ($okDeps -and (Get-PipelineScript "$tb/publish_drafted_replies.py" (Join-Path $tools 'publish_drafted_replies.py') '^def run\(token')) {
  Log "running: python -u tools\publish_drafted_replies.py"
  Push-Location $tools
  & python -u (Join-Path $tools 'publish_drafted_replies.py') 2>&1 | Tee-Object -FilePath $log -Append
  Log "publish_drafted_replies.py exit $LASTEXITCODE"
  Pop-Location
} else {
  Log "SKIP publish_drafted_replies.py (dependency/download issue) -- non-fatal"
}

Copy-Item $log $latest -Force
Log "=== Laptop Bridge Briefing END (core OK) ==="
exit 0
