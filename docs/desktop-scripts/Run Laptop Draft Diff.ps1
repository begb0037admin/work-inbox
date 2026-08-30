<#
Run Laptop Draft Diff.ps1
=========================
work-inbox LAPTOP MIGRATION -- the ongoing draft/final diff capture
(tools/draft_final_diff_capture.py) moved off the admin DESKTOP / off Outlook
COM onto Kevin's Oxford laptop (101L-DE013193 / begb0037.AD-OAK, user
ad-oak\begb0037), reading the server-side Drafts + Sent Items folders over
IMAP+OAuth2 -- the same MSAL broker-silent auth as "Run Laptop Bridge Briefing.ps1".

  MAIL_BACKEND=imap  ->  draft_diff_imap.py reads Drafts + Sent over IMAP
                         (no win32com, no classic Outlook)
  ->  ConversationID-equivalent correlation via the Thread-Index header
  ->  whole-pair redaction (style_corpus_common.py), same as on the desktop
  AI_BACKEND=claude_code  ->  edit_type/note enrichment via headless `claude -p`
                         (subscription auth, CLAUDE_CONFIG_DIR=C:\WorkInboxAI\kevin,
                         ANTHROPIC_API_KEY stripped) -- drains
                         pending_classification.json in the same run. A `claude -p`
                         failure just re-stages the pair (never hard-fails the run).
  ->  LOCAL-ONLY staging dir  <MyDocuments>\CorpusStaging\draft_watch_imap\
      (its own ledger -- the key scheme differs from the COM ledger, so a first
       run here re-baselines and produces ZERO pairs, by design)

This wrapper WRITES a tiny run-status file to GitHub via Push-LaptopRunStatus.ps1:
  data/laptop_status/draftdiff_status.json   (counts + exit code ONLY, never
  any email content) -- so the desktop toast watcher can surface success AND
  failure while the pipeline is on the laptop. Nothing else is pushed.

PARAMS
  -NoAI          correlation + redaction only; skip the `claude -p` enrichment.
                 Pairs stage in pending_classification.json for a later run.
  -Cadence      Bridge (default) | Full  -- cosmetic here; drives Register-LaptopDraftDiff.ps1.
  -NoStatusPush skip the GitHub status-file push (local run only).

LIVE COPY   %USERPROFILE%\work-inbox\Run Laptop Draft Diff.ps1
REFERENCE   work-inbox/docs/desktop-scripts/Run Laptop Draft Diff.ps1
REGISTERED  Register-LaptopDraftDiff.ps1  ->  task "Work Inbox Laptop Draft Diff"
NEEDS ALSO  %USERPROFILE%\work-inbox\Push-LaptopRunStatus.ps1   (pull it alongside this file)

END OF MIGRATION (desktop fixed, if ever reverting)
  1. laptop:  Unregister-ScheduledTask -TaskName 'Work Inbox Laptop Draft Diff' -Confirm:$false
  2. desktop: Enable-ScheduledTask -TaskName 'Draft Diff Capture'
#>
param(
  [switch]$NoAI,
  [ValidateSet('Bridge','Full')] [string]$Cadence = 'Bridge',
  [switch]$NoStatusPush
)

$ErrorActionPreference = 'Continue'
$repo   = 'begb0037admin/work-inbox'
$root   = Join-Path $env:USERPROFILE 'work-inbox'
$logdir = Join-Path $root 'logs'
$stage  = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CorpusStaging\draft_watch_imap'
New-Item -ItemType Directory -Force -Path $root, $logdir, $stage | Out-Null

$stamp    = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')
$log      = Join-Path $logdir "draft_diff_$stamp.log"
$latest   = Join-Path $logdir 'draft_diff_last_run.log'
$statsOut = Join-Path $env:TEMP "wi_draft_diff_stats_$stamp.json"
$pusher   = Join-Path $PSScriptRoot 'Push-LaptopRunStatus.ps1'

function Log($m) {
  $line = "$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))  $m"
  Write-Host $line
  Add-Content -LiteralPath $log -Value $line
}

function Publish-Status([int]$code) {
  if ($NoStatusPush) { return }
  try {
    if (Test-Path $pusher) {
      $a = @{ Kind = 'draftdiff'; ExitCode = $code }
      if (Test-Path $statsOut) { $a['StatsFile'] = $statsOut }
      & $pusher @a 2>&1 | ForEach-Object { Log "status: $_" }
    } else {
      Log "status: Push-LaptopRunStatus.ps1 not found at $pusher -- skipped"
    }
  } catch { Log "status: publish failed (non-fatal): $($_.Exception.Message)" }
}

Log "=== Laptop Draft Diff START  (user $env:USERDOMAIN\$env:USERNAME  host $env:COMPUTERNAME) ==="
Log "params: NoAI=$NoAI  Cadence=$Cadence  NoStatusPush=$NoStatusPush"
Log "staging dir (local only): $stage"
Set-Location $root

# --- isolated Claude Code config: kevin@ (only needed when the enrichment runs) ---
$kevinCfg = 'C:\WorkInboxAI\kevin'
if (-not $NoAI) {
  if (-not (Test-Path (Join-Path $kevinCfg '.credentials.json'))) {
    Log "FATAL: $kevinCfg\.credentials.json not found -- the kevin@ isolated Claude Code config is not logged in."
    Log "FIX:   `$env:CLAUDE_CONFIG_DIR='$kevinCfg'; claude /login   (sign in as kevin@lelitte.co.uk), then re-run."
    Log "  (or run with -NoAI to capture correlated+redacted pairs now and enrich later.)"
    Copy-Item $log $latest -Force
    Publish-Status 3
    exit 3
  }
}

# --- refresh the scripts this needs from main (cache-busted raw pull) ---
$t    = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$base = "https://raw.githubusercontent.com/$repo/main"
$need = @{
  'draft_final_diff_capture.py' = "$base/tools/draft_final_diff_capture.py"
  'style_corpus_common.py'      = "$base/tools/style_corpus_common.py"
  'draft_diff_imap.py'          = "$base/draft_diff_imap.py"
  'imap_mail.py'                = "$base/imap_mail.py"
  'reauth_imap.py'              = "$base/reauth_imap.py"
}
foreach ($name in $need.Keys) {
  try {
    Invoke-WebRequest -UseBasicParsing "$($need[$name])?t=$t" -OutFile (Join-Path $root $name)
    Log "refreshed $name from main"
  } catch {
    Log "WARN: could not refresh $name ($($_.Exception.Message)) -- using the local copy if present"
  }
}

# --- environment: mirror the bridge briefing minus Outlook ---
$env:MAIL_BACKEND      = 'imap'
$env:PYTHONUTF8        = '1'
$env:ANTHROPIC_API_KEY = ''            # force subscription billing (belt-and-braces; the subprocess also strips it)
if ($NoAI) {
  $aiArgs = @('--no-ai')
  $env:AI_BACKEND = 'api'              # irrelevant with --no-ai; keep explicit
  Log "NoAI: correlation + redaction only; pairs stage in pending_classification.json"
} else {
  $aiArgs = @()
  $env:AI_BACKEND          = 'claude_code'
  $env:WI_CLAUDE_CONFIG_DIR = $kevinCfg
  Log "AI_BACKEND=claude_code  cfg=$kevinCfg  -- edit_type/note enrichment via claude -p this run"
}

$ledger = Join-Path $stage 'ledger.json'
Log "running: python -u draft_final_diff_capture.py --mail-backend imap $($aiArgs -join ' ') --ledger-path <stage>\ledger.json --out-dir <stage> --stats-out <temp>"
& python -u (Join-Path $root 'draft_final_diff_capture.py') `
    --mail-backend imap @aiArgs `
    --ledger-path $ledger `
    --out-dir $stage `
    --stats-out $statsOut 2>&1 | Tee-Object -FilePath $log -Append
$rc = $LASTEXITCODE
Log "draft_final_diff_capture.py exit $rc"

Copy-Item $log $latest -Force
Publish-Status $rc
if (Test-Path $statsOut) { Remove-Item $statsOut -Force -ErrorAction SilentlyContinue }

if ($rc -ne 0) {
  Log "=== Laptop Draft Diff END (FAILED, exit $rc) ==="
  if ($rc -eq 1) {
    Log "exit 1 = a phase raised. Common cause: expired IMAP token -> the log shows 'silent token refresh failed'; fix with:  cd `"$root`"; python reauth_imap.py   (one browser click) then re-run."
  }
  exit $rc
}
Log "=== Laptop Draft Diff END (OK) ==="
exit 0
