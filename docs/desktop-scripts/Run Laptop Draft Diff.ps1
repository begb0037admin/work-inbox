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
  ->  LOCAL-ONLY staging dir  %USERPROFILE%\Documents\CorpusStaging\draft_watch_imap\
      (its own ledger -- the key scheme differs from the COM ledger, so a first
       run here re-baselines and produces ZERO pairs, by design)

AI classification: OFF by default. The script's edit_type/note step uses the
Anthropic SDK directly and needs ANTHROPIC_API_KEY; the laptop deliberately has
none (the 27 Aug cutover stopped the ~GBP 36/mo API charge). With -WithAI the
wrapper requires ANTHROPIC_API_KEY in the environment and drops --no-ai.
Without it, correlated + redacted pairs accumulate in pending_classification.json
for a later keyed drain -- the capture (the time-critical part: the draft text
before it vanishes) is fully preserved; only the enrichment waits.

This wrapper WRITES a tiny run-status file to GitHub via Push-LaptopRunStatus.ps1:
  data/laptop_status/draftdiff_status.json   (counts + exit code ONLY, never
  any email content) -- so the desktop toast watcher can surface success AND
  failure while the pipeline is on the laptop. Nothing else is pushed.

PARAMS
  -WithAI        run edit_type/note classification (requires ANTHROPIC_API_KEY).
  -Cadence      Bridge (default) | Full   -- only affects the reference text in
                Register-LaptopDraftDiff.ps1; this wrapper itself is cadence-agnostic.
  -NoStatusPush skip the GitHub status-file push (local run only).

LIVE COPY   %USERPROFILE%\work-inbox\Run Laptop Draft Diff.ps1
REFERENCE   work-inbox/docs/desktop-scripts/Run Laptop Draft Diff.ps1
REGISTERED  Register-LaptopDraftDiff.ps1  ->  task "Work Inbox Laptop Draft Diff"

END OF MIGRATION (desktop fixed, if ever reverting)
  1. laptop:  Unregister-ScheduledTask -TaskName 'Work Inbox Laptop Draft Diff' -Confirm:$false
  2. desktop: Enable-ScheduledTask -TaskName 'Draft Diff Capture'
#>
param(
  [switch]$WithAI,
  [ValidateSet('Bridge','Full')] [string]$Cadence = 'Bridge',
  [switch]$NoStatusPush
)

$ErrorActionPreference = 'Continue'
$repo   = 'begb0037admin/work-inbox'
$root   = Join-Path $env:USERPROFILE 'work-inbox'
$logdir = Join-Path $root 'logs'
$stage  = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CorpusStaging\draft_watch_imap'
New-Item -ItemType Directory -Force -Path $root, $logdir, $stage | Out-Null

$stamp  = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')
$log    = Join-Path $logdir "draft_diff_$stamp.log"
$latest = Join-Path $logdir 'draft_diff_last_run.log'
$statsOut = Join-Path $env:TEMP "wi_draft_diff_stats_$stamp.json"

function Log($m) {
  $line = "$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))  $m"
  Write-Host $line
  Add-Content -LiteralPath $log -Value $line
}

Log "=== Laptop Draft Diff START  (user $env:USERDOMAIN\$env:USERNAME  host $env:COMPUTERNAME) ==="
Log "params: WithAI=$WithAI  Cadence=$Cadence  NoStatusPush=$NoStatusPush"
Log "staging dir (local only): $stage"
Set-Location $root

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

# --- environment: mirror the bridge briefing minus Outlook, minus AI unless -WithAI ---
$env:MAIL_BACKEND = 'imap'
$env:PYTHONUTF8   = '1'
$aiArgs = @('--no-ai')
if ($WithAI) {
  if (-not $env:ANTHROPIC_API_KEY) {
    Log "FATAL: -WithAI given but ANTHROPIC_API_KEY is not set in this session. Aborting (would degrade to --no-ai silently otherwise)."
    Copy-Item $log $latest -Force
    if (-not $NoStatusPush) { & (Join-Path $PSScriptRoot 'Push-LaptopRunStatus.ps1') -Kind draftdiff -ExitCode 3 -Note 'WithAI requested, no ANTHROPIC_API_KEY' 2>&1 | Tee-Object -FilePath $log -Append }
    exit 3
  }
  $aiArgs = @()
  Log "WithAI: ANTHROPIC_API_KEY present -- edit_type/note classification enabled this run."
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

# --- publish the run-status file to GitHub (counts + exit code only) ---
if (-not $NoStatusPush) {
  $pushArgs = @('-Kind', 'draftdiff', '-ExitCode', $rc)
  if (Test-Path $statsOut) { $pushArgs += @('-StatsFile', $statsOut) }
  try {
    & (Join-Path $PSScriptRoot 'Push-LaptopRunStatus.ps1') @pushArgs 2>&1 | Tee-Object -FilePath $log -Append
  } catch {
    Log "WARN: status push failed (non-fatal): $($_.Exception.Message)"
  }
}
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
