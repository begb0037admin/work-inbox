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
  TEAMS HAS NO SEPARATE GUARD: fetch_inbox.py's own TEAMS_BACKEND comment block
  documents that lane_b_call1.py's re-contamination guard already covers the
  microsoft_teams.* tool namespace exactly like it covers calendar's -- Teams
  is raw-digest-only (v1, no AI triage/judgment on its content, see
  docs/LANE_B_TEAMS_CAL_DESIGN.md), so the same single kill-switch that halts
  on any unexpected/write tool call is the deliberate, sufficient safety
  mechanism for both domains. No separate Teams-specific HALT logic is planned.
  CODEX_HOME for the guard/Call-1: UPDATED 2 Sept 2026 evening -- lane_b_call1.py
  now resolves its own primary (Edu, tried first) / failover (personal,
  automatic once Edu's own retry budget is exhausted for ANY reason) identity
  and logs which one served each call. $LaneBCodexHome below is an escape
  hatch only (normally blank) -- see its own comment block for the full story
  and the two wrong ideas rejected en route to this design.
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
  -TeamsBackend      off (default, LIVE) | connector (Lane B, NOT live -- manual
                     proof/testing only until Kevin's cutover go-ahead). Independent
                     of -CalBackend (Teams has no COM/classic-Outlook equivalent to
                     fall back to -- it has only ever been connector-or-nothing).

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
  [ValidateSet('off','connector')] [string]$TeamsBackend = 'off'
)

# ============================================================================
# LANE B identity -- CORRECTED 2 Sept 2026 evening (Kevin, after tonight's Teams
# investigation). Two DIFFERENT wrong ideas were tried and rejected in quick
# succession tonight before landing here, both worth recording so a future
# session doesn't re-propose them:
#   (a) leaving this blank (the state since 1 Sept) -- silently falls through to
#       whatever codex is already logged into on the calling session, which on
#       the laptop's RDP default profile is Edu. That's not "wrong" by itself
#       (Edu IS meant to be tried first -- see below) but it was ACCIDENTAL, not
#       deliberate, and gave no fallback at all when Edu failed (which is
#       exactly what happened to Teams tonight: 4 straight timed-out attempts,
#       no data, ever, on Edu).
#   (b) hardcoding this to the personal account (`C:\WorkInboxAI\codex-laneb`)
#       -- tried first as "the fix" tonight, then Kevin corrected it: the 1 Sept
#       move to personal-only was a TESTING-PHASE workaround for burning
#       through Edu's monthly cap fast during heavy testing, not the permanent
#       architecture. Edu should still be tried first, every time; personal is
#       the safety net, not the new default.
# ACTUAL design: `lane_b_call1.py` itself now has explicit primary(Edu)/
# failover(personal) resolution + automatic failover built in (PRIMARY_CODEX_HOME
# / FAILOVER_CODEX_HOME, WI_LANE_B_CODEX_HOME_FAILOVER env override) -- Edu is
# tried first, its EXISTING retry budget is exhausted, and only then does it
# automatically retry the same call against personal, logging clearly which
# identity actually served each call. This wrapper does NOT need to pick an
# identity or implement any fallback itself any more -- that decision now lives
# in the Python code, not here. This var is kept ONLY as an explicit escape
# hatch to override BOTH calendar's and Teams' primary identity in one place if
# ever needed (e.g. Edu login expires and needs bypassing entirely) -- leave it
# blank for normal operation, which lets lane_b_call1.py's own primary/failover
# logic run as designed.
#   $LaneBCodexHome = '' (blank, NORMAL)      -> lane_b_call1.py resolves its
#                                                 own primary (Edu, explicit
#                                                 deliberate default) and
#                                                 failover (personal) itself.
#   $LaneBCodexHome = '<any path>'            -> escape hatch: forces BOTH
#                                                 CODEX_HOME and
#                                                 WI_LANE_B_CODEX_HOME to this
#                                                 path, which becomes
#                                                 lane_b_call1.py's PRIMARY
#                                                 (its own failover logic still
#                                                 applies on top of that).
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
Log "params: CoreOnly=$CoreOnly  CalBackend=$CalBackend  TeamsBackend=$TeamsBackend  log=$log"
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
$env:TEAMS_BACKEND                = $TeamsBackend
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
    Invoke-WebRequest -UseBasicParsing "$base/docs/desktop-scripts/$($sf.Name)`?t=$t" -OutFile $dl
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

  if ($LaneBCodexHome -ne '') {
    $env:CODEX_HOME           = $LaneBCodexHome
    $env:WI_LANE_B_CODEX_HOME = $LaneBCodexHome
    Log "Lane B: CODEX_HOME escape-hatch override -> $LaneBCodexHome (becomes lane_b_call1.py's PRIMARY; its own failover logic still applies on top)"
  } else {
    Log "Lane B: `$LaneBCodexHome not set (normal) -- lane_b_call1.py resolves its own primary(Edu)/failover(personal) identity and logs which one actually served each call"
  }
  Log "running: python lane_b_cal_guard.py --run --domain $laneBDomain"

  # FAST-FAIL Lane B primary (added 3 Sept 2026, Kevin's explicit requirement).
  # Edu's connector quota is exhausted until 1 Oct -- every primary attempt on
  # Edu is currently either a guaranteed failure or, on the evidence of the
  # 3 Sept 12:51 test run, an intermittent success; either way the default
  # budget (Teams primary: 2 sub-attempts x 360s + 75s quiet-gaps around every
  # call) could burn ~15 min/run before ever reaching failover. Force ONE short
  # primary attempt per domain, no outer retry, a short inter-call quiet gap;
  # switch to personal (failover) immediately if primary doesn't succeed.
  # Failover's own budget (WI_LANE_B_TIMEOUT/WI_LANE_B_RETRIES, unset here) is
  # left at full strength -- personal is proven reliable and keeps the benefit
  # of the doubt. Live-proven 3 Sept: ~29 min (double-failover, pre-change) ->
  # ~Xm (fast-fail, see HANDOVER.md for the exact proving-run number).
  # REVERT THIS WHOLE BLOCK after 1 Oct 2026 when Edu's monthly quota resets --
  # it deliberately overrides the 3 Sept Teams-primary-budget fix (which exists
  # for exactly the case this block short-circuits: a primary genuinely worth
  # waiting on) and is only correct while Edu cannot be trusted to finish in a
  # reasonable time.
  $env:WI_LANE_B_PRIMARY_TIMEOUT            = '45'
  $env:WI_LANE_B_PRIMARY_MAX_ATTEMPTS       = '1'
  $env:WI_LANE_B_PRIMARY_TIMEOUT_TEAMS      = '45'
  $env:WI_LANE_B_PRIMARY_MAX_ATTEMPTS_TEAMS = '1'
  $env:WI_LANE_B_PRIMARY_RETRIES            = '1'
  $env:WI_LANE_B_SNAPSHOT_GAP_S             = '15'
  $env:WI_LANE_B_WARMUP_TIMEOUT             = '45'
  Log "Lane B: FAST-FAIL primary active -- 1 attempt/45s per domain, 15s inter-call gap, immediate personal failover on any primary failure (Edu quota dead until 1 Oct -- REVERT after)"
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

# --- CORE: fetch_inbox.py  (Phase 1 IMAP -> combined claude -p triage -> Phase 4 push -> Phase 5 CC sync) ---
Log "running: python -u fetch_inbox.py   [MAIL_BACKEND=imap  CAL_BACKEND=$CalBackend  TEAMS_BACKEND=$TeamsBackend  AI_BACKEND=claude_code  cfg=$kevinCfg]"
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
