<#
Run Laptop Bridge Briefing.ps1
==============================
work-inbox LAPTOP BRIDGE -- REAL mail-only briefings from Kevin's Oxford laptop
(101L-DE013193 / begb0037.AD-OAK, user ad-oak\begb0037) for as long as the admin
DESKTOP's M365 device-registration is broken (0x8004dec5) and Outlook COM there
is dead. Kevin + Max are repairing the desktop separately.

This is NOT the Phase 4 parity shadow ("Run Laptop Parity Shadow.ps1"). This one
pushes for real:
    Microsoft 365 connector mail pull  ->  claude -p triage (kevin@ isolated config)
    ->  Phase 4  data/briefing.json  -> GitHub
    ->  Phase 5  command-centre task-suggestion sync
    ->  (best effort) needs_reply.json + drafted_replies.json publishers

Calendar: NONE by default. CAL_BACKEND=com, but classic Outlook is not running /
connected on the laptop, so fetch_inbox.py degrades the calendar phases to empty
and a warning (handled path, not a crash). The bridge briefing simply has no
calendar section. Accepted for the bridge; do not generalise this no-op to other
machines without checking their Outlook installation and session state.

LANE B (connector calendar + Teams) GLUE -- calendar added 1 Sept 2026, Teams
wiring added 2 Sept 2026 evening (Drew) -- NEITHER LIVE YET:
  -CalBackend connector and/or -TeamsBackend connector wire in
  lane_b_cal_guard.py --run --domain <calendar|teams|both> BEFORE fetch_inbox.py
  (ONE guard call covering whichever domain(s) were requested -- see below for
  why Teams reuses this same gate rather than getting its own):
    exit 0 (clean)      -> proceed with CAL_BACKEND=connector / TEAMS_BACKEND=connector
            for whichever domain(s) were requested, for real.
    exit 1 (persistent HALT -- a real calendar change during the read window, or
            a write tool was seen) -> Disable-ScheduledTask on THIS task + a
            best-effort BurntToast + fall back to CAL_BACKEND=com AND
            TEAMS_BACKEND=off for this cycle only (degrades to empty+warning on
            this laptop, same as always -- the briefing still ships). No auto
            re-enable; Kevin investigates.
    exit 3 (transient -- connector unavailable / could not verify this cycle)
            -> log + fall back to CAL_BACKEND=com AND TEAMS_BACKEND=off for this
            cycle; task STAYS enabled; try again next cadence.
    exit 5 (MODEL POLICY VIOLATION -- deterministic model/effort code/config bug)
            -> log + local BurntToast + fall back to CAL_BACKEND=com AND
            TEAMS_BACKEND=off for this cycle; task STAYS enabled, but this needs
            investigation rather than another silent transient retry.
    exit -1 (external/native termination surfaced by PowerShell as 0xFFFFFFFF)
            -> log the abnormal termination + fall back to CAL_BACKEND=com AND
            TEAMS_BACKEND=off for this cycle; task STAYS enabled and retries next
            cadence. This is distinct from an ordinary connector-transient exit.
  TEAMS HAS NO SEPARATE GUARD: fetch_inbox.py's own TEAMS_BACKEND comment block
  documents that lane_b_call1.py's re-contamination guard already covers the
  microsoft_teams.* tool namespace exactly like it covers calendar's -- Teams
  is raw-digest-only (v1, no AI triage/judgment on its content, see
  docs/LANE_B_TEAMS_CAL_DESIGN.md), so the same single kill-switch that halts
  on any unexpected/write tool call is the deliberate, sufficient safety
  mechanism for both domains. No separate Teams-specific HALT logic is planned.
  CODEX_HOME for the guard/Call-1: lane_b_call1.py reads the ordered
  lane_b_identities.json ring and logs each identity's result. Profiles that
  are missing or unauthenticated are skipped. Every domain starts again at
  ring position 1 on its next call and next run.
  **THE LIVE SCHEDULED TASK STAYS ON CAL_BACKEND=com.** This wiring is dormant
  until Kevin gives the explicit go-ahead to register/re-register the task with
  -CalBackend connector -- same cutover discipline as the connector mail path.
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
  -TeamsBackend      off (default, LIVE) | connector (Lane B, NOT live -- manual
                     proof/testing only until Kevin's cutover go-ahead). Independent
                     of -CalBackend (Teams has no COM/classic-Outlook equivalent to
                     fall back to -- it has only ever been connector-or-nothing).
  -MailBackend       connector (default; IMAP retired on this laptop). Runs
                     `lane_b_call1.py --domain mail` directly (NOT through
                     lane_b_cal_guard.py -- that guard's pre/post snapshot-diff is
                     calendar-specific and Kevin explicitly declined a kill-switch
                     rework for mail; the SAME verb-based re-contamination guard
                     already live for calendar/Teams, inside lane_b_call1.py
                      itself, is the sole mechanism here too). Exit 0 -> proceed with
                      MAIL_BACKEND=connector. Exit 1 (HALT -- a write/off-scope tool
                      call was observed) -> disable THIS task + local BurntToast +
                      abort this cycle. Exit 5 (MODEL POLICY VIOLATION) -> log +
                      local BurntToast + abort this cycle, task stays enabled, but
                      investigate the deterministic code/config bug rather than
                      treating it as ordinary connector flakiness. Any other exit
                      (2/3 -- usage error / codex run failed) -> abort this cycle,
                      task stays enabled, and retry next cadence. There is no
                      IMAP fallback on this retired laptop. KNOWN GAP, disclosed
                      not hidden: unlike calendar/Teams' guard, a mail HALT is NOT
                      yet threaded through Push-LaptopRunStatus.ps1/Watch-
                      BridgeBriefing.ps1, so there is no CROSS-MACHINE desktop toast
                      for it yet -- only the local laptop toast + this run's own log.
                      Follow-up, not a safety gap (the guard/disable still fire).

LIVE COPY   %USERPROFILE%\work-inbox\Run Laptop Bridge Briefing.ps1
REFERENCE   work-inbox/docs/desktop-scripts/Run Laptop Bridge Briefing.ps1
REGISTERED  Register-LaptopBridgeBriefing.ps1  ->  task "Work Inbox Bridge Briefing"

REVERT / END OF BRIDGE
  1. Unregister-ScheduledTask -TaskName 'Work Inbox Bridge Briefing' -Confirm:$false   (on the laptop)
  2. Enable-ScheduledTask   -TaskName 'Work Inbox Briefing'                            (on the admin desktop)
#>
param(
  [switch]$CoreOnly,
  [ValidateSet('com','connector')] [string]$CalBackend = 'com',
  [ValidateSet('off','connector')] [string]$TeamsBackend = 'off',
  [ValidateSet('connector')] [string]$MailBackend = 'connector'
)

# ============================================================================
# LANE B identity ring -- the ordered entries in lane_b_identities.json are
# the only source of identity/path configuration. lane_b_call1.py skips a
# missing or unauthenticated profile and advances to the next entry on any
# connector failure. This wrapper clears ambient CODEX_HOME values so they
# cannot override the checked-in ring.
# ============================================================================
$TaskName       = 'Work Inbox Bridge Briefing'   # this task's own name, for Disable-ScheduledTask on a guard HALT

$ErrorActionPreference = 'Continue'
$root   = Join-Path $env:USERPROFILE 'work-inbox'
$tools  = Join-Path $root 'tools'
$logdir = Join-Path $root 'logs'
New-Item -ItemType Directory -Force -Path $root, $tools, $logdir | Out-Null

# One lock shared by both Oxford Windows identities.  The two profile
# checkouts are separate directories, so a lock under $root would not stop a
# scheduled AD-OAK run colliding with an SSH/manual begb0037-a run.  Holding an
# exclusive FileStream makes the check atomic; a stale lock file is reused on
# the next run once its old process has released the handle.
$lockPath = 'C:\Users\Public\bridge_briefing.lock'
$script:BridgeLockStream = $null
try {
  $script:BridgeLockStream = [System.IO.File]::Open(
    $lockPath,
    [System.IO.FileMode]::OpenOrCreate,
    [System.IO.FileAccess]::ReadWrite,
    [System.IO.FileShare]::None)
  $script:BridgeLockStream.SetLength(0)
  $lockText = "PID=$PID`nUSER=$env:USERDOMAIN\$env:USERNAME`nSTART=$([DateTime]::Now.ToString('o'))`n"
  $lockBytes = [Text.Encoding]::UTF8.GetBytes($lockText)
  $script:BridgeLockStream.Write($lockBytes, 0, $lockBytes.Length)
  $script:BridgeLockStream.Flush()
  Write-Host "BRIDGE LOCK ACQUIRED -- $lockPath"
} catch [System.IO.IOException] {
  Write-Host "BRIDGE ALREADY RUNNING -- another Work Inbox pipeline holds $lockPath" -ForegroundColor Yellow
  Write-Host "Wait for that run to finish; no duplicate was started."
  exit 2
} catch {
  Write-Host "BRIDGE LOCK ERROR -- could not establish $lockPath ($($_.Exception.Message))" -ForegroundColor Red
  Write-Host "No pipeline was started."
  exit 2
}
Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
  if ($script:BridgeLockStream) {
    $script:BridgeLockStream.Dispose()
    $script:BridgeLockStream = $null
  }
} | Out-Null

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
      $pusherArgs = @{
        Kind = 'briefing'; ExitCode = $code
        LaneBGuard = $LaneBGuardResult; LaneBGuardDetail = $LaneBGuardDetail
      }
      if ($LaneBDomainsSummary -and $LaneBDomainsSummary.Count -gt 0) { $pusherArgs.LaneBDomains = $LaneBDomainsSummary }
      & $pusher @pusherArgs 2>&1 | ForEach-Object { Log "status: $_" }
    } else {
      Log "status: Push-LaptopRunStatus.ps1 not found next to this wrapper -- skipped"
    }
  } catch { Log "status: publish failed (non-fatal): $($_.Exception.Message)" }
}

# Lane B guard outcome (calendar and/or Teams -- ONE guard, see the header note)
# for THIS run, surfaced to Publish-Status above so the cross-machine (desktop)
# toast watcher sees a HALT even when the briefing itself still succeeds via the
# CAL_BACKEND=com / TEAMS_BACKEND=off fallback. Values: 'not-run' (neither
# -CalBackend connector nor -TeamsBackend connector was passed -- the live task
# today), 'clean', 'halted' (persistent HALT -- task disabled, THIS is the one
# that must be hard to miss), 'transient' (connector unavailable this cycle, not
# actionable), 'unexpected-<n>'.
$LaneBGuardResult = 'not-run'
$LaneBGuardDetail = ''
# Per-domain status/count/served_by/primary_failover_identical (added 3 Sept
# 2026, regression-fix verification -- see Push-LaptopRunStatus.ps1's own
# header note). Populated below, after the guard call, from the freshest
# data\lane_b\*_lane_b.json run log -- pure counts/status, never content.
$LaneBDomainsSummary = $null

Log "=== Laptop Bridge Briefing START  (user $env:USERDOMAIN\$env:USERNAME  host $env:COMPUTERNAME) ==="
Log "params: CoreOnly=$CoreOnly  CalBackend=$CalBackend  TeamsBackend=$TeamsBackend  MailBackend=$MailBackend  log=$log"
Set-Location $root

# --- isolated Claude Code config for the briefing triage phase ---
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
$env:MAIL_BACKEND                 = $MailBackend
$env:CAL_BACKEND                  = $CalBackend
$env:TEAMS_BACKEND                = $TeamsBackend
$env:WI_BRIDGE_ALLOW_EMPTY_CALENDAR = '1'         # no calendar source on the laptop -> empty calendar/absences must not veto the Phase 4 push
$env:WI_MAIL_PARALLEL            = ''             # explicit: this is a REAL run, not a parallel capture
$env:WI_TRIAGE_V2                = '1'            # explicit pin, 9 Sep 2026 (Kevin, "turn it on"). fetch_inbox.py's own
                                                   # os.environ.get("WI_TRIAGE_V2","1") default is already ON as of HANDOVER
                                                   # section M (live-proven: needs=25, suppressedCount=17 on the 07:23 run)
                                                   # -- this line makes that explicit on the laptop task rather than relying
                                                   # solely on the code default, so a future default change can't silently
                                                   # flip this task's behaviour too. To roll back: set this to '' (or delete
                                                   # the line) AND see HANDOVER.md restore point at fetch_inbox.py ~line 364.
$env:PYTHONUTF8                  = '1'
$connectorBudgetSeconds = 1200
$env:WI_LANE_B_RUN_BUDGET_S = [string]$connectorBudgetSeconds
$env:WI_LANE_B_RUN_DEADLINE_EPOCH = [string]([DateTimeOffset]::UtcNow.ToUnixTimeSeconds() + $connectorBudgetSeconds)

# --- refresh pipeline scripts from main (cache-busted raw pull, same mechanism the desktop uses) ---
$t    = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$base = 'https://raw.githubusercontent.com/begb0037admin/work-inbox/main'
foreach ($f in 'fetch_inbox.py','normalise_pull.py','lane_b_call1.py','lane_b_cal_guard.py','codex_model_policy.py','lane_b_identities.json') {
  # codex_model_policy.py added 10 Sep 2026 (Priority 4, touchpoint-3 Codex
  # review finding) -- lane_b_call1.py now imports it; without refreshing it
  # here alongside lane_b_call1.py, a refreshed lane_b_call1.py could import
  # a stale or missing copy and fail loudly at import time (exit 2) on the
  # very next run. hris-dashboard's fetch_osm_report_connector.py imports
  # this same module from this sibling work-inbox clone on the same laptop,
  # so refreshing it here also keeps that script current -- no separate
  # refresh step needed in hris-dashboard's own wrapper for this file.
  try {
    Invoke-WebRequest -UseBasicParsing "$base/$f`?t=$t" -TimeoutSec 30 -OutFile (Join-Path $root $f)
    Log "refreshed $f from main"
  } catch {
    Log "WARN: could not refresh $f ($($_.Exception.Message)) -- using the local copy"
  }
}

# --- self-refresh the wrapper + status pusher from docs/desktop-scripts/ on
#     main, GUARDED (min-size + marker) so a truncated/failed pull can never
#     replace a working copy. These two are NOT in the loop above because they
#     live under docs/desktop-scripts/, not repo root. The wrapper refresh takes
#     effect on the NEXT run (this process is already parsed); the
#     Push-LaptopRunStatus.ps1 refresh takes effect THIS run (it is invoked at
#     the exit points below). Added 3 Sept 2026 -- previously both files could
#     only be updated on the laptop by a manual SSH/RDP file copy, so repo edits
#     silently never reached the live task (root cause of the stale
#     Push-LaptopRunStatus.ps1 '-LaneBGuard' param error). ---
foreach ($sf in @(
    @{ Name = 'Run Laptop Bridge Briefing.ps1'; Marker = 'Laptop Bridge Briefing START' }
    @{ Name = 'Push-LaptopRunStatus.ps1';       Marker = 'LaneBDomains' }
)) {
  $dl = (Join-Path $root $sf.Name) + '.download'
  try {
    Invoke-WebRequest -UseBasicParsing "$base/docs/desktop-scripts/$($sf.Name)`?t=$t" -TimeoutSec 30 -OutFile $dl
    if ((Get-Item $dl).Length -lt 1000) { throw 'downloaded file too small' }
    if (-not (Select-String -Quiet -LiteralPath $dl -Pattern $sf.Marker)) { throw "missing marker /$($sf.Marker)/" }
    Move-Item -Force $dl (Join-Path $root $sf.Name)
    Log "refreshed $($sf.Name) from main (docs/desktop-scripts, guarded)"
  } catch {
    Log "WARN: could not refresh $($sf.Name) ($($_.Exception.Message)) -- keeping local copy"
    if (Test-Path $dl) { Remove-Item $dl -Force }
  }
}

# --- LANE B calendar/Teams HALT guard -- only when CalBackend=connector and/or
#     TeamsBackend=connector was explicitly passed (the live task does not pass
#     either yet; see the header note). ONE guard call covers whichever domain(s)
#     were requested -- Teams has no separate guard, see header note for why. ---
$laneBDomain = $null
if ($CalBackend -eq 'connector' -and $TeamsBackend -eq 'connector') { $laneBDomain = 'both' }
elseif ($CalBackend -eq 'connector') { $laneBDomain = 'calendar' }
elseif ($TeamsBackend -eq 'connector') { $laneBDomain = 'teams' }

if ($laneBDomain) {
  # Regression fix, 3 Sept 2026 (root cause of the 2 Sept 20:58 live incident --
  # see HANDOVER-LATEST-2026-09-02-teams-regression.md): ALWAYS start from a
  # clean CODEX_HOME / WI_LANE_B_CODEX_HOME / WI_LANE_B_CODEX_HOME_FAILOVER
  # environment before every Lane B invocation, exactly like every documented
  # manual interactive test this week has done by hand (`Remove-Item
  # Env:\CODEX_HOME`, `Remove-Item Env:\WI_LANE_B_CODEX_HOME`). Scheduled-task
  # processes inherit the FULL ambient user/system environment, which on this
  # host still carries CODEX_HOME=C:\WorkInboxAI\codex-laneb left over from
  # 1 Sept's now-superseded personal-only testing phase. Without this clear,
  # lane_b_call1.py's PRIMARY_CODEX_HOME resolution (WI_LANE_B_CODEX_HOME wins,
  # else inherited CODEX_HOME, else the true OS default) silently picks up that
  # leftover value -- which happens to be the EXACT SAME literal path as
  # FAILOVER_CODEX_HOME's own hardcoded default, collapsing primary and
  # failover into one identity. Automatic failover then can't fire (nothing
  # distinct to fail over to), "primary" silently runs as the old personal
  # test account instead of Edu, and a struggling connector call on that
  # account produces exactly the "only one attempt, identity=primary,
  # codex_failed" shape seen in the incident -- for BOTH domains, since both
  # go through the same misidentified identity. This clear runs unconditionally
  # BEFORE the escape-hatch check below, so the escape hatch (when actually
  # used) still applies cleanly on top of a known-clean starting environment.
  Remove-Item Env:\CODEX_HOME -ErrorAction SilentlyContinue
  Remove-Item Env:\WI_LANE_B_CODEX_HOME -ErrorAction SilentlyContinue
  Remove-Item Env:\WI_LANE_B_CODEX_HOME_FAILOVER -ErrorAction SilentlyContinue
  Log "Lane B: cleared any ambient CODEX_HOME/WI_LANE_B_CODEX_HOME(_FAILOVER) from this process's environment before resolving identity (regression fix, 3 Sept 2026)"

  Log "Lane B: lane_b_call1.py will load lane_b_identities.json and start each domain at ring position 1"
  Log "running: python lane_b_cal_guard.py --run --domain $laneBDomain"

  & python -u (Join-Path $root 'lane_b_cal_guard.py') --run --domain $laneBDomain 2>&1 | Tee-Object -FilePath $log -Append
  $guardRc = $LASTEXITCODE
  Log "lane_b_cal_guard.py exit $guardRc"

  # Per-domain status/count/served_by/primary_failover_identical, for the
  # cross-machine run-status push (regression-fix verification, 3 Sept 2026 --
  # see Push-LaptopRunStatus.ps1's header note). lane_b_call1.py writes
  # data\lane_b\<ts>_lane_b.json UNCONDITIONALLY on every invocation
  # (success, halt, or codex_failed) via its own _write_run_log() -- pick the
  # most recently written one rather than relying on lane_b_normalised.json
  # (which is left at its last-good state, not overwritten, on a bad cycle).
  # Best-effort only: never blocks or fails the run.
  try {
    $laneBLogDir = Join-Path $root 'data\lane_b'
    $latestLog = Get-ChildItem -LiteralPath $laneBLogDir -Filter '*_lane_b.json' -ErrorAction Stop |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestLog) {
      $logJson = Get-Content -Raw -LiteralPath $latestLog.FullName | ConvertFrom-Json
      $summary = @{}
      foreach ($d in $logJson.per_domain.PSObject.Properties.Name) {
        $pd = $logJson.per_domain.$d
        $summary[$d] = @{
          status = $pd.status; count = $pd.count; served_by = $pd.served_by
          primary_failover_identical = $pd.primary_failover_identical
        }
      }
      $LaneBDomainsSummary = $summary
      Log "Lane B per-domain summary ($($latestLog.Name)): $($summary | ConvertTo-Json -Depth 4 -Compress)"
    } else {
      Log "WARN: no data\lane_b\*_lane_b.json found to summarise (unexpected -- lane_b_call1.py should always write one)"
    }
  } catch {
    Log "WARN: could not build Lane B per-domain summary (non-fatal): $($_.Exception.Message)"
  }

  switch ($guardRc) {
    0 {
      Log "Lane B guard CLEAN -- proceeding with CAL_BACKEND=$CalBackend TEAMS_BACKEND=$TeamsBackend"
      $LaneBGuardResult = 'clean'
    }
    1 {
      Log "Lane B guard PERSISTENT HALT (a real calendar change during the read window, or a write tool was seen) -- disabling '$TaskName' and falling back to CAL_BACKEND=com / TEAMS_BACKEND=off for THIS cycle only"
      $CalBackend = 'com'
      $TeamsBackend = 'off'
      $LaneBGuardResult = 'halted'
      $LaneBGuardDetail = "task '$TaskName' disabled; calendar/Teams fell back to COM/off this cycle. See data\lane_b\ and data\codex_runs\GUARD_TRIPPED_* on the laptop."
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
        New-BurntToastNotification -Text 'Work Inbox - Lane B guard HALTED', $LaneBGuardDetail
        Log "local BurntToast fired"
      } catch {
        Log "WARN: BurntToast unavailable/failed ($($_.Exception.Message)) -- LOCAL toast skipped; the HALT + task-disable above are still real, and the desktop toast (via Publish-Status) is independent of this and still fires"
      }
    }
    3 {
      Log "Lane B guard TRANSIENT (connector unavailable / could not verify this cycle) -- falling back to CAL_BACKEND=com / TEAMS_BACKEND=off for THIS cycle only; task stays enabled, will retry next cadence"
      $CalBackend = 'com'
      $TeamsBackend = 'off'
      $LaneBGuardResult = 'transient'
      $LaneBGuardDetail = 'connector unavailable this cycle; not actionable, no toast.'
    }
    5 {
      # Added 10 Sep 2026 (Priority 4, touchpoint-3 Codex review finding):
      # MUST be distinguished from `default` below, which the pre-existing
      # code already treats identically to "unexpected, retry silently" --
      # exactly the masking this exit code exists to prevent. Not a security
      # HALT (task stays enabled), but genuinely not a normal transient
      # failure either -- surfaced with a toast so it doesn't sit silent.
      Log "Lane B guard MODEL POLICY VIOLATION (a code/config bug in codex_model_policy usage, NOT connector unavailability) -- falling back to CAL_BACKEND=com / TEAMS_BACKEND=off for THIS cycle only; task stays enabled but this needs investigation, not just a retry"
      $CalBackend = 'com'
      $TeamsBackend = 'off'
      $LaneBGuardResult = 'policy-violation'
      $LaneBGuardDetail = 'MODEL POLICY VIOLATION -- code/config bug in codex_model_policy usage. See lane_b_call1.py / hris-dashboard fetch_osm_report_connector.py HANDOVER.md. Not disabled (not a mailbox-safety issue), but will keep failing every cycle until fixed.'
      try {
        Import-Module BurntToast -ErrorAction Stop
        New-BurntToastNotification -Text 'Work Inbox - Lane B MODEL POLICY VIOLATION', $LaneBGuardDetail
        Log "local BurntToast fired (model policy violation)"
      } catch {
        Log "WARN: BurntToast unavailable/failed ($($_.Exception.Message)) -- LOCAL toast skipped; the violation is still real and logged above"
      }
    }
    -1 {
      # PowerShell exposes a terminated native child returning 0xFFFFFFFF as
      # -1. lane_b_cal_guard.py does not intentionally return this value;
      # surface it separately so external termination is not mislabeled as an
      # ordinary unexpected connector result.
      Log "Lane B guard EXTERNALLY TERMINATED (exit -1 / 0xFFFFFFFF) -- falling back to CAL_BACKEND=com / TEAMS_BACKEND=off for THIS cycle only; task stays enabled, will retry next cadence"
      $CalBackend = 'com'
      $TeamsBackend = 'off'
      $LaneBGuardResult = 'external-termination'
      $LaneBGuardDetail = 'lane_b_cal_guard.py surfaced exit -1 (0xFFFFFFFF), which is not an intentional guard return; likely external/native termination. Fell back conservatively and left the task enabled.'
    }
    default {
      Log "Lane B guard unexpected exit $guardRc -- treating conservatively: falling back to CAL_BACKEND=com / TEAMS_BACKEND=off for THIS cycle only; task stays enabled"
      $CalBackend = 'com'
      $TeamsBackend = 'off'
      $LaneBGuardResult = "unexpected-$guardRc"
      $LaneBGuardDetail = "lane_b_cal_guard.py exited $guardRc (not 0/1/3) -- treated conservatively, not disabled."
    }
  }
  $env:CAL_BACKEND   = $CalBackend    # re-assert in case the guard downgraded it above
  $env:TEAMS_BACKEND = $TeamsBackend  # re-assert in case the guard downgraded it above
}

# --- LANE B MAIL guard -- added 9 Sept 2026, HANDOVER.md section Q (Kevin's fresh
#     explicit risk acceptance). The laptop is now connector-only, so this runs
#     `lane_b_call1.py --domain mail` DIRECTLY, not through lane_b_cal_guard.py --
#     that guard's pre/post snapshot-diff is calendar-specific and Kevin explicitly
#     declined a kill-switch rework for mail; the SAME verb-based re-contamination
#     guard already live for calendar/Teams (inside lane_b_call1.py itself) is the
#     sole safety mechanism here too. Mail's fetch_mail_domain() always targets
#     the configured identity ring; no ambient CODEX_HOME is used. --
$LaneBMailGuardResult = 'not-run'
$LaneBMailGuardDetail = ''
if ($MailBackend -eq 'connector') {
  Log "running: python lane_b_call1.py --domain mail"
  & python -u (Join-Path $root 'lane_b_call1.py') --domain mail 2>&1 | Tee-Object -FilePath $log -Append
  $mailGuardRc = $LASTEXITCODE
  Log "lane_b_call1.py --domain mail exit $mailGuardRc"
  switch ($mailGuardRc) {
    0 {
      Log "Lane B mail guard CLEAN (or a sub-domain was merely unavailable this cycle, not halted) -- proceeding with MAIL_BACKEND=connector"
      $LaneBMailGuardResult = 'clean'
    }
    1 {
      Log "Lane B mail guard HALT (a write/off-scope tool call was observed) -- disabling '$TaskName' and aborting; no IMAP fallback exists on this retired laptop"
      $LaneBMailGuardResult = 'halted'
      $LaneBMailGuardDetail = "task '$TaskName' disabled; connector run aborted. See data\lane_b\ and data\codex_runs\GUARD_TRIPPED_* on the laptop."
      try {
        Disable-ScheduledTask -TaskName $TaskName -ErrorAction Stop | Out-Null
        Log "Disabled scheduled task '$TaskName' -- re-enable manually after investigating: Enable-ScheduledTask -TaskName '$TaskName'"
      } catch {
        Log "WARN: could not Disable-ScheduledTask '$TaskName' ($($_.Exception.Message)) -- disable it manually"
      }
      try {
        Import-Module BurntToast -ErrorAction Stop
        New-BurntToastNotification -Text 'Work Inbox - Lane B MAIL guard HALTED', $LaneBMailGuardDetail
        Log "local BurntToast fired (mail guard HALT)"
      } catch {
        Log "WARN: BurntToast unavailable/failed ($($_.Exception.Message)) -- LOCAL toast skipped; the HALT + task-disable above are still real. KNOWN GAP (disclosed, not a safety gap): unlike calendar/Teams, this is not yet threaded through Push-LaptopRunStatus.ps1, so there is no cross-machine desktop toast for a mail HALT yet -- this run's own log is authoritative until that follow-up is built."
      }
      Copy-Item $log $latest -Force
      Publish-Status 1
      Log "=== Laptop Bridge Briefing END (mail guard halted) ==="
      exit 1
    }
    5 {
      # Added 10 Sep 2026 (Priority 4, touchpoint-3 Codex review finding):
      # MUST be distinguished from `default` below -- same reasoning as the
      # calendar/Teams switch above. Not a security HALT (task stays
      # enabled), but not ordinary transient flakiness either.
      Log "Lane B mail guard MODEL POLICY VIOLATION (a code/config bug in codex_model_policy usage, NOT connector unavailability) -- aborting this cycle; task stays enabled but this needs investigation, not just a retry"
      $LaneBMailGuardResult = 'policy-violation'
      $LaneBMailGuardDetail = 'MODEL POLICY VIOLATION -- code/config bug in codex_model_policy usage. Not disabled (not a mailbox-safety issue), but will keep failing every cycle until fixed.'
      try {
        Import-Module BurntToast -ErrorAction Stop
        New-BurntToastNotification -Text 'Work Inbox - Lane B mail MODEL POLICY VIOLATION', $LaneBMailGuardDetail
        Log "local BurntToast fired (mail model policy violation)"
      } catch {
        Log "WARN: BurntToast unavailable/failed ($($_.Exception.Message)) -- LOCAL toast skipped; the violation is still real and logged above"
      }
      Copy-Item $log $latest -Force
      Publish-Status 5
      Log "=== Laptop Bridge Briefing END (mail model policy violation) ==="
      exit 5
    }
    default {
      Log "Lane B mail guard non-zero exit $mailGuardRc (2=usage/env error, 3=codex run failed, other=unexpected) -- aborting this cycle; task stays enabled and will retry next cadence"
      $LaneBMailGuardResult = "unexpected-$mailGuardRc"
      $LaneBMailGuardDetail = "lane_b_call1.py --domain mail exited $mailGuardRc -- connector-only run aborted; task stays enabled."
      Copy-Item $log $latest -Force
      Publish-Status $mailGuardRc
      Log "=== Laptop Bridge Briefing END (mail connector unavailable/error) ==="
      exit $mailGuardRc
    }
  }
  $env:MAIL_BACKEND = $MailBackend  # re-assert the connector-only backend
}

# --- CORE: fetch_inbox.py  (Phase 1 mail -> combined claude -p triage -> Phase 4 push -> Phase 5 CC sync) ---
Log "running: python -u fetch_inbox.py   [MAIL_BACKEND=$MailBackend  CAL_BACKEND=$CalBackend  TEAMS_BACKEND=$TeamsBackend  AI_BACKEND=claude_code  cfg=$kevinCfg]"
& python -u (Join-Path $root 'fetch_inbox.py') 2>&1 | Tee-Object -FilePath $log -Append
$rc = $LASTEXITCODE
Log "fetch_inbox.py exit $rc"

if ($rc -ne 0) {
  Log "CORE FAILED (exit $rc). NOT running publishers. Check the log above for the failing phase."
  if ($rc -eq 1) {
    Log "exit 1 = a phase raised. Check the connector/Lane B error above; there is no IMAP re-auth fallback on this retired path. A Phase 4 safe-write veto is also logged as 'Safe write blocked briefing update: ...'."
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
    Invoke-WebRequest -UseBasicParsing "$url`?t=$tt" -TimeoutSec 30 -OutFile "$dest.download"
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

function Invoke-PublisherWithTimeout([string]$scriptPath, [string]$name, [int]$timeoutSeconds) {
  $proc = $null
  try {
    $python = (Get-Command python -ErrorAction Stop).Source
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $python
    $psi.Arguments = '-u "' + $scriptPath + '"'
    $psi.WorkingDirectory = $tools
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    # Close stdin explicitly: a codex/connector child must see EOF, never the
    # scheduler's inherited input stream.
    $proc.StandardInput.Close()
    $outTask = $proc.StandardOutput.ReadToEndAsync()
    $errTask = $proc.StandardError.ReadToEndAsync()
    if (-not $proc.WaitForExit($timeoutSeconds * 1000)) {
      Log "WARN: $name exceeded ${timeoutSeconds}s -- killing its whole process tree and skipping"
      try { & taskkill.exe /T /F /PID $proc.Id | Out-Null } catch { try { $proc.Kill() } catch {} }
      try { [void]$proc.WaitForExit(10000) } catch {}
      return 124
    }
    $stdout = $outTask.GetAwaiter().GetResult()
    $stderr = $errTask.GetAwaiter().GetResult()
    foreach ($line in (($stdout + "`n" + $stderr) -split "`r?`n")) {
      if ($line) { Log "[$name] $line" }
    }
    return $proc.ExitCode
  } catch {
    Log "WARN: $name could not start/finish ($($_.Exception.Message)) -- publisher skipped"
    if ($proc) { try { $proc.Kill() } catch {} }
    return 125
  } finally {
    if ($proc) { $proc.Dispose() }
  }
}

$tb = "$base/tools"
$okDeps = (Get-PipelineScript "$tb/style_corpus_common.py"  (Join-Path $tools 'style_corpus_common.py')  '^def recipient_tier') `
      -and (Get-PipelineScript "$tb/phase_failure_notify.py" (Join-Path $tools 'phase_failure_notify.py') '^def notify_phase_failure')

if ($okDeps -and (Get-PipelineScript "$tb/publish_needs_reply.py" (Join-Path $tools 'publish_needs_reply.py') '^def run\(token')) {
  Log "running: python -u tools\publish_needs_reply.py"
  $publisherRc = Invoke-PublisherWithTimeout (Join-Path $tools 'publish_needs_reply.py') 'publish_needs_reply.py' 180
  Log "publish_needs_reply.py exit $publisherRc"
} else {
  Log "SKIP publish_needs_reply.py (dependency/download issue) -- non-fatal"
}

if ($okDeps -and (Get-PipelineScript "$tb/publish_drafted_replies.py" (Join-Path $tools 'publish_drafted_replies.py') '^def run\(token')) {
  Log "running: python -u tools\publish_drafted_replies.py"
  # Drafted Replies performs one bounded connector subject batch through the
  # Lane B identity ring. Keep the wrapper bound above the connector's own
  # timeout so a slow but valid Oxford-account call is not killed early.
  $publisherRc = Invoke-PublisherWithTimeout (Join-Path $tools 'publish_drafted_replies.py') 'publish_drafted_replies.py' 420
  Log "publish_drafted_replies.py exit $publisherRc"
} else {
  Log "SKIP publish_drafted_replies.py (dependency/download issue) -- non-fatal"
}

Copy-Item $log $latest -Force
Publish-Status 0
Log "=== Laptop Bridge Briefing END (core OK) ==="
exit 0
