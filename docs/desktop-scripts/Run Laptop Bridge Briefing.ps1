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

Calendar: NONE by default. CAL_BACKEND=com, but there is no classic Outlook on
the laptop, so fetch_inbox.py degrades the calendar phases to empty + a warning
(handled path, not a crash). The bridge briefing simply has no calendar section.
Accepted for the bridge.

LANE B (connector calendar) GLUE -- added 1 Sept 2026, NOT LIVE YET:
  -CalBackend connector wires in lane_b_cal_guard.py --run BEFORE fetch_inbox.py:
    exit 0 (clean)      -> proceed with CAL_BACKEND=connector for real.
    exit 1 (persistent HALT -- a real calendar change during the read window, or
            a write tool was seen) -> Disable-ScheduledTask on THIS task + a
            best-effort BurntToast + fall back to CAL_BACKEND=com for this cycle
            only (degrades to empty+warning on this laptop, same as always --
            the briefing still ships). No auto re-enable; Kevin investigates.
    exit 3 (transient -- connector unavailable / could not verify this cycle)
            -> log + fall back to CAL_BACKEND=com for this cycle; task STAYS
            enabled; try again next cadence.
  CODEX_HOME for the guard/Call-1: see $LaneBCodexHome below -- ONE clearly
  commented variable, currently blank (uses whatever codex is already logged
  into on this session -- the Oxford Edu account, at the time this was wired).
  Flip it to the dedicated Lane B login once Kevin does the personal-ChatGPT
  switch and tells us the CODEX_HOME to point at.
  **THE LIVE SCHEDULED TASK STAYS ON CAL_BACKEND=com.** This wiring is dormant
  until Kevin gives the explicit go-ahead to register/re-register the task with
  -CalBackend connector -- same cutover discipline as the mail IMAP migration.
  Passing -CalBackend connector by hand (this script only, not the live task)
  is how Kevin/the coordinator proves it end to end before that go-ahead.

Mirrors the live desktop "Run Inbox Briefing.bat" environment, minus Outlook COM
and minus the hope@ overflow config (single account on the laptop for now -- a
Pro-cap hit degrades that one run; acceptable for a short bridge).

PARAMS
  -CoreOnly          run only fetch_inbox.py (skip the two downstream publishers).
                     Use this for the first supervised run.
  -CalBackend        com (default, LIVE) | connector (Lane B, NOT live -- manual
                     proof/testing only until Kevin's cutover go-ahead)

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

# ============================================================================
# LANE B -- single override points (edit these two lines only to change identity
# or task name; nothing else in this script should need touching for that).
#   $LaneBCodexHome = '' (blank)         -> use whatever codex is already logged
#                                            into on this session (Edu, as of
#                                            1 Sept 2026 -- Kevin has not yet
#                                            done the personal-account switch).
#   $LaneBCodexHome = 'C:\WorkInboxAI\codex-laneb'  -> ONE-LINE FLIP to the
#                                            dedicated Lane B login once Kevin
#                                            has run `codex login` into it on
#                                            the personal ChatGPT account.
# ============================================================================
$LaneBCodexHome = ''
$TaskName       = 'Work Inbox Bridge Briefing'   # this task's own name, for Disable-ScheduledTask on a guard HALT

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

# Publish a tiny GitHub run-status file (counts / exit code only, no email
# content) so the desktop toast watcher can surface a FAILED laptop run --
# and, since 2 Sept 2026, a Lane B calendar-guard HALT even on a run that
# otherwise succeeds (the guard falls back to CAL_BACKEND=com and the
# briefing still ships, so $code alone would report "ok" and hide a real
# safety trip -- see $LaneBGuardResult below).
# Best-effort: never changes the exit code, never throws.
function Publish-Status([int]$code) {
  try {
    $pusher = Join-Path $PSScriptRoot 'Push-LaptopRunStatus.ps1'
    if (Test-Path $pusher) {
      & $pusher -Kind briefing -ExitCode $code -LaneBGuard $LaneBGuardResult -LaneBGuardDetail $LaneBGuardDetail 2>&1 | ForEach-Object { Log "status: $_" }
    } else {
      Log "status: Push-LaptopRunStatus.ps1 not found next to this wrapper -- skipped"
    }
  } catch { Log "status: publish failed (non-fatal): $($_.Exception.Message)" }
}

# Lane B calendar-guard outcome for THIS run, surfaced to Publish-Status above so
# the cross-machine (desktop) toast watcher sees a HALT even when the briefing
# itself still succeeds via the CAL_BACKEND=com fallback. Values: 'not-run'
# (CalBackend never passed -CalBackend connector -- the live task today),
# 'clean', 'halted' (persistent HALT -- task disabled, THIS is the one that must
# be hard to miss), 'transient' (connector unavailable this cycle, not actionable),
# 'unexpected-<n>'.
$LaneBGuardResult = 'not-run'
$LaneBGuardDetail = ''

Log "=== Laptop Bridge Briefing START  (user $env:USERDOMAIN\$env:USERNAME  host $env:COMPUTERNAME) ==="
Log "params: CoreOnly=$CoreOnly  CalBackend=$CalBackend  log=$log"
Set-Location $root

# --- isolated Claude Code config: kevin@ ONLY (no hope@ failover on the laptop yet) ---
$kevinCfg = 'C:\WorkInboxAI\kevin'
if (-not (Test-Path (Join-Path $kevinCfg '.credentials.json'))) {
  Log "FATAL: $kevinCfg\.credentials.json not found -- the kevin@ isolated Claude Code config is not logged in."
  Log "FIX:   `$env:CLAUDE_CONFIG_DIR='$kevinCfg'; claude /login   (sign in as kevin@lelitte.co.uk), then re-run."
  Copy-Item $log $latest -Force
  Publish-Status 3
  exit 3
}

$env:AI_BACKEND                    = 'claude_code'
$env:ANTHROPIC_API_KEY            = ''            # force subscription billing (matches desktop .bat)
$env:WI_CLAUDE_CONFIG_DIR         = $kevinCfg     # -> claude -p gets CLAUDE_CONFIG_DIR=C:\WorkInboxAI\kevin
$env:WI_CLAUDE_CONFIG_DIR_FALLBACK = ''           # explicit: single account, no hope@ overflow
$env:MAIL_BACKEND                 = 'imap'
$env:CAL_BACKEND                  = $CalBackend
$env:WI_BRIDGE_ALLOW_EMPTY_CALENDAR = '1'         # no calendar source on the laptop -> empty calendar/absences must not veto the Phase 4 push
$env:WI_MAIL_PARALLEL            = ''             # explicit: this is a REAL run, not a parallel capture
$env:PYTHONUTF8                  = '1'

# --- refresh pipeline scripts from main (cache-busted raw pull, same mechanism the desktop uses) ---
$t    = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$base = 'https://raw.githubusercontent.com/begb0037admin/work-inbox/main'
foreach ($f in 'fetch_inbox.py','imap_mail.py','reauth_imap.py','normalise_pull.py','lane_b_call1.py','lane_b_cal_guard.py') {
  try {
    Invoke-WebRequest -UseBasicParsing "$base/$f`?t=$t" -OutFile (Join-Path $root $f)
    Log "refreshed $f from main"
  } catch {
    Log "WARN: could not refresh $f ($($_.Exception.Message)) -- using the local copy"
  }
}

# --- LANE B calendar HALT guard -- only when CalBackend=connector was explicitly
#     passed (the live task does not pass this; see the header note). ---
if ($CalBackend -eq 'connector') {
  if ($LaneBCodexHome -ne '') {
    $env:CODEX_HOME           = $LaneBCodexHome
    $env:WI_LANE_B_CODEX_HOME = $LaneBCodexHome
    Log "Lane B: CODEX_HOME override -> $LaneBCodexHome"
  } else {
    Log "Lane B: `$LaneBCodexHome not set -- using whatever codex is already logged into on this session"
  }
  Log "running: python lane_b_cal_guard.py --run"
  & python -u (Join-Path $root 'lane_b_cal_guard.py') --run 2>&1 | Tee-Object -FilePath $log -Append
  $guardRc = $LASTEXITCODE
  Log "lane_b_cal_guard.py exit $guardRc"
  switch ($guardRc) {
    0 {
      Log "Lane B guard CLEAN -- proceeding with CAL_BACKEND=connector"
      $LaneBGuardResult = 'clean'
    }
    1 {
      Log "Lane B guard PERSISTENT HALT (a real calendar change during the read window, or a write tool was seen) -- disabling '$TaskName' and falling back to CAL_BACKEND=com for THIS cycle only"
      $CalBackend = 'com'
      $LaneBGuardResult = 'halted'
      $LaneBGuardDetail = "task '$TaskName' disabled; calendar fell back to COM this cycle. See data\lane_b\ and data\codex_runs\GUARD_TRIPPED_* on the laptop."
      try {
        Disable-ScheduledTask -TaskName $TaskName -ErrorAction Stop | Out-Null
        Log "Disabled scheduled task '$TaskName' -- re-enable manually after investigating: Enable-ScheduledTask -TaskName '$TaskName'"
      } catch {
        Log "WARN: could not Disable-ScheduledTask '$TaskName' ($($_.Exception.Message)) -- disable it manually"
      }
      # LOCAL toast (laptop, only seen if someone is logged into this session).
      # Cross-machine (desktop) notification of the SAME event happens via
      # Publish-Status below -> data/laptop_status/briefing_status.json ->
      # Watch-BridgeBriefing.ps1 on the desktop -- see that script for the
      # matching toast. Both are required; neither replaces the other.
      try {
        Import-Module BurntToast -ErrorAction Stop
        New-BurntToastNotification -Text 'Work Inbox - Lane B calendar guard HALTED', "Task '$TaskName' disabled. $LaneBGuardDetail"
        Log "local BurntToast fired"
      } catch {
        Log "WARN: BurntToast unavailable/failed ($($_.Exception.Message)) -- LOCAL toast skipped; the HALT + task-disable above are still real, and the desktop toast (via Publish-Status) is independent of this and still fires"
      }
    }
    3 {
      Log "Lane B guard TRANSIENT (connector unavailable / could not verify this cycle) -- falling back to CAL_BACKEND=com for THIS cycle only; task stays enabled, will retry next cadence"
      $CalBackend = 'com'
      $LaneBGuardResult = 'transient'
      $LaneBGuardDetail = 'connector unavailable this cycle; not actionable, no toast.'
    }
    default {
      Log "Lane B guard unexpected exit $guardRc -- treating conservatively: falling back to CAL_BACKEND=com for THIS cycle only; task stays enabled"
      $CalBackend = 'com'
      $LaneBGuardResult = "unexpected-$guardRc"
      $LaneBGuardDetail = "lane_b_cal_guard.py exited $guardRc (not 0/1/3) -- treated conservatively, not disabled."
    }
  }
  $env:CAL_BACKEND = $CalBackend   # re-assert in case the guard downgraded it above
}

# --- CORE: fetch_inbox.py  (Phase 1 IMAP -> combined claude -p triage -> Phase 4 push -> Phase 5 CC sync) ---
Log "running: python -u fetch_inbox.py   [MAIL_BACKEND=imap  CAL_BACKEND=$CalBackend  AI_BACKEND=claude_code  cfg=$kevinCfg]"
& python -u (Join-Path $root 'fetch_inbox.py') 2>&1 | Tee-Object -FilePath $log -Append
$rc = $LASTEXITCODE
Log "fetch_inbox.py exit $rc"

if ($rc -ne 0) {
  Log "CORE FAILED (exit $rc). NOT running publishers. Check the log above for the failing phase."
  if ($rc -eq 1) {
    Log "exit 1 = a phase raised. Common causes: (a) expired IMAP token -> the log shows 'IMAP mail sign-in expired'; fix with:  cd `"$root`"; python reauth_imap.py   (one browser click) then re-run.  (b) a Phase 4 safe-write veto -> the log shows 'Safe write blocked briefing update: ...'."
  }
  Copy-Item $log $latest -Force
  Publish-Status $rc
  Log "=== Laptop Bridge Briefing END (core failed) ==="
  exit $rc
}

if ($CoreOnly) {
  Log "CoreOnly set -- skipping the needs_reply / drafted_replies publishers."
  Copy-Item $log $latest -Force
  Publish-Status 0
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
Publish-Status 0
Log "=== Laptop Bridge Briefing END (core OK) ==="
exit 0
