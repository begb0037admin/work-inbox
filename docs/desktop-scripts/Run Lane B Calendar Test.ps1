<#
Run Lane B Calendar Test.ps1
============================
FALLBACK EXECUTION PATH for Lane B (connector calendar), to be pasted into
Kevin's RDP session on the Oxford work laptop 101L-DE013193, running as the
DOMAIN user AD-OAK\begb0037 (the account the "Work Inbox Bridge Briefing"
scheduled task runs as).

IDENTITY (decision 1 Sept ~15:30): Lane B uses a DEDICATED CODEX_HOME
(-CodexHome, default C:\WorkInboxAI\codex-laneb) signed into Kevin's PERSONAL
ChatGPT account (kevin@lelitte.co.uk) -- NOT the Oxford Edu account, whose
500/month credit cap won't sustain Lane B + Kevin's interactive use. Kevin's
interactive codex on this box (whatever CODEX_HOME he uses, typically ~/.codex
= Edu) is completely untouched. One-time setup: `$env:CODEX_HOME='<dir>'; codex
login` as personal, + add the Microsoft Outlook Calendar connector to personal
in ChatGPT settings (Teams already works there).

Why this script exists: SSH into 101L-DE013193 only lands as the local-admin
account begb0037-a, which cannot read into AD-OAK\begb0037's profile, and direct
SSH as the domain user is currently broken. Until that is fixed, this script is
how a Lane B run gets executed in the right context.

WHAT IT DOES
  0. sets CODEX_HOME + WI_LANE_B_CODEX_HOME to -CodexHome; prints codex version +
     the account_id from <CodexHome>\auth.json (FATAL if that login is missing)
  1. pulls the Lane B scripts fresh from raw.githubusercontent main
  2. python lane_b_call1.py --domain calendar   -> data/lane_b/lane_b_normalised.json
     prints meta.lane_b.domains.calendar (status / count / attempts / tool_calls)
     and the calendar events (day / start / all_day / subject)
  3. python lane_b_cal_guard.py --run           -> exit 0 clean / 3 transient / 1 HALT
     (SKIPPED with -Fast)
  4. (with -RealBriefing, implied by -Fast) one MAIL_BACKEND=imap CAL_BACKEND=connector
     AI_BACKEND=claude_code fetch_inbox.py run -- PRODUCES AND PUSHES a real briefing.json.

  -Fast  = the ~6 min one-time proof: step 1 -> step 2 -> (if calendar status ok) step 4.
           SKIPS step 3 (the snapshot HALT guard). Rationale: the "does connector
           calendar reach a real briefing" proof does not need the HALT guard --
           that is a scheduled-cadence safety layer, validated when it is wired into
           the bridge wrapper. lane_b_call1.py's OWN verb-based re-contamination
           guard still runs in step 2 and HALTs on any write tool.
  full run (no -Fast) = steps 1-4, ~12-15 min (codex is ~3 min/call on this box and
           the guard adds 3 more codex calls: its pre-snapshot, its own calendar
           call, its post-snapshot).

USAGE (paste into the RDP PowerShell session, as AD-OAK\begb0037):
    cd $env:USERPROFILE\work-inbox
    # get this script fresh too:
    $u='https://raw.githubusercontent.com/begb0037admin/work-inbox/main/docs/desktop-scripts/Run%20Lane%20B%20Calendar%20Test.ps1'
    Invoke-WebRequest -UseBasicParsing "$u`?t=$([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())" -OutFile '.\Run Lane B Calendar Test.ps1'
    # FAST one-time proof (~6 min, no guard):
    powershell -NoProfile -ExecutionPolicy Bypass -File '.\Run Lane B Calendar Test.ps1' -Fast
    # full check incl. the snapshot guard (~12-15 min):
    powershell -NoProfile -ExecutionPolicy Bypass -File '.\Run Lane B Calendar Test.ps1' -RealBriefing

PARAMS
  -RunDir       default $env:USERPROFILE\work-inbox
  -Fast         one-time proof, skips step 3, implies -RealBriefing (~6 min)
  -RealBriefing also run step 4 (produces + PUSHES a real briefing.json)
  -CodexHome    dedicated Lane B CODEX_HOME (default C:\WorkInboxAI\codex-laneb),
                signed into the PERSONAL ChatGPT account. FATAL if its auth.json
                is missing (prints the one-time `codex login` steps).
  -CodexBin     full path to codex.cmd/.exe/.ps1 if a bare 'codex' is not on PATH
                for this session (sets $env:WI_CODEX_BIN)
  -Retries      WI_LANE_B_RETRIES for the connector-flake retry (default 3;
                pass -Retries 1 to fail fast instead of retrying a flaky connector)
  -Timeout      per codex-exec-call timeout in seconds (default 360)

Target: Windows PowerShell 5.1. No &&, no ternary, no ??.
#>
param(
  [string]$RunDir = (Join-Path $env:USERPROFILE 'work-inbox'),
  [switch]$RealBriefing,
  [switch]$Fast,        # ONE-TIME validation: refresh -> lane_b_call1 --domain calendar -> (if ok) fetch_inbox.py connector briefing. SKIPS the snapshot guard (step 3). ~6 min instead of ~12-15. Implies -RealBriefing.
  [string]$CodexHome = 'C:\WorkInboxAI\codex-laneb',  # DEDICATED codex login for Lane B (personal ChatGPT account). Kevin's interactive codex (Edu) is untouched.
  [string]$CodexBin = '',
  [int]$Retries = 3,
  [int]$Timeout = 360   # per codex-exec-call timeout (s). 0.151.0 cold-starts ~3+ min; the runner does one warm-up call first.
)
if ($Fast) { $RealBriefing = $true }

$ErrorActionPreference = 'Continue'
function Stamp { (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }
function Say($m) { Write-Host ("[{0}] {1}" -f (Stamp), $m) }

$base = 'https://raw.githubusercontent.com/begb0037admin/work-inbox/main'
$logDir = Join-Path $RunDir 'logs'
$null = New-Item -ItemType Directory -Force -Path $RunDir, $logDir
$log = Join-Path $logDir ("lane_b_test_{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$transcribing = $false
try { Start-Transcript -Path $log -Append | Out-Null; $transcribing = $true } catch { Write-Host "(transcript unavailable: $($_.Exception.Message))" }

Say "=== Run Lane B Calendar Test ==="
Say ("host {0}   user {1}\{2}   RunDir {3}" -f $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME, $RunDir)
Say ("Fast={0}   RealBriefing={1}   Retries={2}" -f [bool]$Fast, [bool]$RealBriefing, $Retries)

if (-not (Test-Path $RunDir)) { Say "FATAL: RunDir not found: $RunDir"; if ($transcribing) { Stop-Transcript | Out-Null }; exit 2 }
Set-Location $RunDir

# --- 0. codex identity: DEDICATED Lane B CODEX_HOME (personal ChatGPT account) ---
if ($CodexBin -ne '') { $env:WI_CODEX_BIN = $CodexBin; Say "WI_CODEX_BIN=$CodexBin" }
$env:WI_LANE_B_RETRIES      = "$Retries"
$env:WI_LANE_B_TIMEOUT      = "$Timeout"
$env:WI_LANE_B_SNAP_TIMEOUT = "$Timeout"
$env:PYTHONUTF8 = '1'
$env:CODEX_HOME            = $CodexHome
$env:WI_LANE_B_CODEX_HOME  = $CodexHome
Say "timeouts: per-call ${Timeout}s (+ a one-shot warm-up); retries $Retries"
Say "CODEX_HOME (Lane B, dedicated) = $CodexHome"

$authJson = Join-Path $CodexHome 'auth.json'
if (-not (Test-Path $authJson)) {
  Say "FATAL: $authJson not found -- the Lane B codex login does not exist yet."
  Say "FIX (in this RDP session, one time):"
  Say "    `$env:CODEX_HOME = '$CodexHome'"
  Say "    codex login          # browser flow -> sign in as your PERSONAL ChatGPT account (kevin@lelitte.co.uk)"
  Say "  then confirm the personal account has the Microsoft Outlook Calendar connector added (Teams already works there),"
  Say "  and re-run this script."
  if ($transcribing) { Stop-Transcript | Out-Null }
  exit 2
}
Say "--- codex identity (CODEX_HOME=$CodexHome) ---"
try { & codex --version } catch { Say "codex --version failed: $($_.Exception.Message)" }
try { & codex login status } catch { Say "codex login status failed: $($_.Exception.Message)" }
$acct = '(unparsed)'; $acctEmail = ''; $acctPlan = ''
try {
  $a = Get-Content $authJson -Raw | ConvertFrom-Json
  if ($a.tokens -and $a.tokens.account_id) { $acct = "$($a.tokens.account_id)" } elseif ($a.account_id) { $acct = "$($a.account_id)" }
  $idt = if ($a.tokens -and $a.tokens.id_token) { $a.tokens.id_token } elseif ($a.id_token) { $a.id_token } else { '' }
  if ($idt -and $idt.Split('.').Count -ge 2) {
    $p = $idt.Split('.')[1].Replace('-','+').Replace('_','/'); switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } }
    $raw = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p))
    if ($raw -match '"email"\s*:\s*"([^"]+)"') { $acctEmail = $Matches[1] }
    if ($raw -match '"chatgpt_plan_type"\s*:\s*"([^"]+)"') { $acctPlan = $Matches[1] }
    if ($acct -eq '(unparsed)' -and $raw -match '"chatgpt_account_id"\s*:\s*"([^"]+)"') { $acct = $Matches[1] }
  }
  Say "$CodexHome\auth.json  account_id=$acct  email=$acctEmail  plan=$acctPlan"
} catch { Say "could not parse $authJson : $($_.Exception.Message)" }
$cfg = Join-Path $CodexHome 'config.toml'
if (Test-Path $cfg) { Say ("$CodexHome\config.toml sha1 = {0}" -f (Get-FileHash $cfg -Algorithm SHA1).Hash) } else { Say "(no $CodexHome\config.toml yet -- codex writes it on first run)" }

# --- 1. refresh the 4 Lane B scripts from main ---
Say "--- refreshing scripts from main ---"
$t = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$ok = $true
foreach ($f in 'fetch_inbox.py','normalise_pull.py','lane_b_call1.py','lane_b_cal_guard.py') {
  try {
    Invoke-WebRequest -UseBasicParsing "$base/$f`?t=$t" -OutFile (Join-Path $RunDir $f)
    Say "  refreshed $f"
  } catch {
    Say "  FAILED to refresh $f : $($_.Exception.Message)"
    $ok = $false
  }
}
if (-not $ok) { Say "FATAL: could not pull one or more scripts from GitHub."; if ($transcribing) { Stop-Transcript | Out-Null }; exit 2 }

# --- 2. calendar Call-1 ---
Say "--- STEP 2: python lane_b_call1.py --domain calendar ---"
& python -u (Join-Path $RunDir 'lane_b_call1.py') --domain calendar
$call1rc = $LASTEXITCODE
Say "lane_b_call1.py exit $call1rc"

$normPath = Join-Path $RunDir 'data\lane_b\lane_b_normalised.json'
$calStatus = '(no file)'
if (Test-Path $normPath) {
  try {
    $doc = Get-Content $normPath -Raw | ConvertFrom-Json
    $cal = $doc.meta.lane_b.domains.calendar
    $calStatus = $cal.status
    Say "  meta.lane_b.ts          = $($doc.meta.lane_b.ts)"
    Say "  domains.calendar.status = $($cal.status)"
    Say "  domains.calendar.count  = $($cal.count)"
    Say "  domains.calendar.tool_calls = $($cal.tool_calls -join ', ')"
    Say "  domains.calendar.attempts:"
    Write-Host (($cal.attempts | ConvertTo-Json -Depth 6))
    Say "  calendar events (day / start / all_day / subject):"
    $doc.calendar | Select-Object day, start, all_day, subject | Format-Table -AutoSize | Out-String | Write-Host
  } catch {
    Say "  could not parse $normPath : $($_.Exception.Message)"
  }
} else {
  Say "  $normPath not written (connector unavailable this cycle, or an error above)."
}

# --- 3. the guard end-to-end (what the scheduled task will call) ---
if ($Fast) {
  $guardrc = 'skipped (-Fast)'
  $guardMeaning = 'snapshot HALT guard skipped for this one-time validation. lane_b_call1.py''s own verb-based re-contamination guard still ran in step 2 (HALTs on any write tool). The snapshot guard is a scheduled-cadence safety layer, wired into the bridge wrapper separately.'
  Say "--- STEP 3 skipped (-Fast) ---"
} else {
  Say "--- STEP 3: python lane_b_cal_guard.py --run ---"
  & python -u (Join-Path $RunDir 'lane_b_cal_guard.py') --run
  $guardrc = $LASTEXITCODE
  $guardMeaning = switch ($guardrc) {
    0 { 'CLEAN -- calendar verified unchanged across the read window; lane_b_normalised.json is trustworthy' }
    1 { 'PERSISTENT HALT -- a real calendar change during the window, OR a write tool was seen. lane_b_normalised.json quarantined. The bridge wrapper would Disable-ScheduledTask + toast here.' }
    3 { 'TRANSIENT -- connector unavailable / could not verify this cycle. No connector calendar this run; the task stays enabled. Re-run later.' }
    default { "unexpected exit $guardrc" }
  }
  Say "lane_b_cal_guard.py exit $guardrc -- $guardMeaning"
}

# --- 4. optional: a real connector briefing (produces + PUSHES briefing.json) ---
$briefrc = 'skipped'
if ($RealBriefing) {
  Say "--- STEP 4: real briefing  MAIL_BACKEND=imap CAL_BACKEND=connector AI_BACKEND=claude_code ---"
  $kevinCfg = 'C:\WorkInboxAI\kevin'
  if (-not (Test-Path (Join-Path $kevinCfg '.credentials.json'))) {
    Say "FATAL: $kevinCfg\.credentials.json not found -- the kevin@ isolated Claude Code config is not logged in. Skipping step 4."
    $briefrc = 'blocked (no kevin@ claude config)'
  } else {
    $env:MAIL_BACKEND = 'imap'
    $env:CAL_BACKEND = 'connector'
    $env:AI_BACKEND = 'claude_code'
    $env:WI_CLAUDE_CONFIG_DIR = $kevinCfg
    $env:WI_CLAUDE_CONFIG_DIR_FALLBACK = ''
    $env:ANTHROPIC_API_KEY = ''
    $env:WI_BRIDGE_ALLOW_EMPTY_CALENDAR = '1'   # if the connector flakes this run, do not veto the push
    & python -u (Join-Path $RunDir 'fetch_inbox.py')
    $briefrc = $LASTEXITCODE
    Say "fetch_inbox.py exit $briefrc"
  }
} else {
  Say "STEP 4 skipped (re-run with -RealBriefing to produce + PUSH a real briefing.json)."
}

# --- summary to paste back ---
Say "================ PASTE THIS BACK ================"
Say "host/user            : $env:COMPUTERNAME  $env:USERDOMAIN\$env:USERNAME"
Say "codex identity       : account_id=$acct  email=$acctEmail  plan=$acctPlan"
Say "step2 call1 exit     : $call1rc   (calendar status: $calStatus)"
Say "step3 guard exit     : $guardrc   ($guardMeaning)"
Say "step4 briefing exit  : $briefrc"
Say "transcript           : $log"
Say "================================================"
if ($transcribing) { Stop-Transcript | Out-Null }
exit 0
