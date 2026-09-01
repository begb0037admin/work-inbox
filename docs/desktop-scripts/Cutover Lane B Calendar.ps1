<#
Cutover Lane B Calendar.ps1
===========================
ONE-SHOT cutover of the live "Work Inbox Bridge Briefing" scheduled task from
CAL_BACKEND=com to CAL_BACKEND=connector (Lane B connector calendar), on the
personal ChatGPT account via the dedicated CODEX_HOME.

Kevin pastes this ONCE in his RDP session on 101L-DE013193, as AD-OAK\begb0037.
Expect it to take ~25-30 min end to end (a connector probe ~4 min, then a full
guarded briefing run ~15-20 min). It PRINTS a rollback one-liner and a
PASTE-THIS-BACK block at the end.

It changes NOTHING and exits non-zero if any precondition fails. It does NOT
touch the mail/IMAP path. It is safe to re-run (idempotent-ish: timestamped
backups, skips steps already done).

USAGE (in the RDP PowerShell session, as AD-OAK\begb0037):
    cd $env:USERPROFILE\work-inbox
    $u='https://raw.githubusercontent.com/begb0037admin/work-inbox/main/docs/desktop-scripts/Cutover%20Lane%20B%20Calendar.ps1'
    Invoke-WebRequest -UseBasicParsing "$u`?t=$([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())" -OutFile '.\Cutover Lane B Calendar.ps1'
    powershell -NoProfile -ExecutionPolicy Bypass -File '.\Cutover Lane B Calendar.ps1'

PARAMS
  -TaskName     default 'Work Inbox Bridge Briefing'
  -CodexHome    dedicated Lane B codex login dir (personal account). Default C:\WorkInboxAI\codex-laneb
  -Email        expected personal-account email. Default kevin@lelitte.co.uk
  -RunDir       default $env:USERPROFILE\work-inbox
  -SkipRun      do everything except the final force-run + verify
  -RunTimeoutMin  max minutes to wait for the force-run (default 22)

Target: Windows PowerShell 5.1. No &&, no ternary, no ??.
#>
param(
  [string]$TaskName = 'Work Inbox Bridge Briefing',
  [string]$CodexHome = 'C:\WorkInboxAI\codex-laneb',
  [string]$Email = 'kevin@lelitte.co.uk',
  [string]$RunDir = (Join-Path $env:USERPROFILE 'work-inbox'),
  [switch]$SkipRun,
  [int]$RunTimeoutMin = 22
)

$ErrorActionPreference = 'Continue'
function Stamp { (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }
function Say($m) { Write-Host ("[{0}] {1}" -f (Stamp), $m) }
function Fail($m) { Write-Host ("[{0}] FATAL: {1}" -f (Stamp), $m) -ForegroundColor Red; if ($transcribing) { Stop-Transcript | Out-Null }; exit 1 }

$base  = 'https://raw.githubusercontent.com/begb0037admin/work-inbox/main'
$ts    = Get-Date -Format 'yyyyMMdd-HHmmss'
$logdir = Join-Path $RunDir 'logs'
$null = New-Item -ItemType Directory -Force -Path $RunDir, $logdir
$transcript = Join-Path $logdir "cutover_lane_b_$ts.log"
$transcribing = $false
try { Start-Transcript -Path $transcript -Append | Out-Null; $transcribing = $true } catch { Write-Host "(transcript unavailable: $($_.Exception.Message))" }

Say "=== Cutover Lane B Calendar ==="
Say ("host {0}   user {1}\{2}" -f $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME)
Say ("TaskName='{0}'  CodexHome='{1}'  Email='{2}'  RunDir='{3}'  SkipRun={4}" -f $TaskName, $CodexHome, $Email, $RunDir, [bool]$SkipRun)
if (-not (Test-Path $RunDir)) { Fail "RunDir not found: $RunDir" }
Set-Location $RunDir

# result holders for the PASTE-THIS-BACK block
$R = [ordered]@{
  precond_auth = ''; precond_email = ''; precond_plan = ''; precond_account = ''; precond_probe = ''
  script_bak = ''; task_xml_bak = ''; live_ps1 = ''
  new_action = ''; guard_exit = ''; calendar_line = ''; briefing_before = ''; briefing_after = ''; fetch_exit = ''
  rollback_a = ''; rollback_b = ''; rollback_script = ''
}

# ============================================================================
# 1. PRECONDITIONS -- fail loud, change nothing
# ============================================================================
Say "--- 1. PRECONDITIONS ---"

$authJson = Join-Path $CodexHome 'auth.json'
if (-not (Test-Path $authJson)) {
  Fail "$authJson not found. In this session run:  `$env:CODEX_HOME='$CodexHome'; codex login   (sign in as your PERSONAL ChatGPT account, $Email), then re-run this script."
}
$R.precond_auth = "OK ($authJson exists)"

function Decode-JwtPayloadRaw($jwt) {
  $parts = $jwt.Split('.')
  if ($parts.Count -lt 2) { return $null }
  $p = $parts[1].Replace('-', '+').Replace('_', '/')
  switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } 1 { $p += '' } }
  try { return [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) } catch { return $null }
}

$auth = $null
try { $auth = Get-Content $authJson -Raw | ConvertFrom-Json } catch { Fail "could not parse $authJson : $($_.Exception.Message)" }
$idTok = $null
if ($auth.tokens -and $auth.tokens.id_token) { $idTok = $auth.tokens.id_token }
elseif ($auth.id_token) { $idTok = $auth.id_token }
if (-not $idTok) { Fail "no id_token found in $authJson (looked at .tokens.id_token and .id_token)" }

$claimsRaw = Decode-JwtPayloadRaw $idTok
if (-not $claimsRaw) { Fail "could not base64url-decode the id_token payload" }

$foundEmail = ''; $foundAccount = ''
try {
  $claims = $claimsRaw | ConvertFrom-Json
  if ($claims.email) { $foundEmail = "$($claims.email)" }
  $oa = $claims.'https://api.openai.com/auth'
  if ($oa -and $oa.chatgpt_account_id) { $foundAccount = "$($oa.chatgpt_account_id)" }
} catch { }
if (-not $foundEmail -and ($claimsRaw -match '"email"\s*:\s*"([^"]+)"')) { $foundEmail = $Matches[1] }
if (-not $foundAccount -and ($claimsRaw -match '"chatgpt_account_id"\s*:\s*"([^"]+)"')) { $foundAccount = $Matches[1] }
if (-not $foundAccount -and $auth.tokens -and $auth.tokens.account_id) { $foundAccount = "$($auth.tokens.account_id)" }
if (-not $foundAccount -and $auth.account_id) { $foundAccount = "$($auth.account_id)" }

$foundPlan = '(unknown)'
if ($claimsRaw -match '"chatgpt_plan_type"\s*:\s*"([^"]+)"') { $foundPlan = $Matches[1] }
elseif ($claimsRaw -match '"plan_type"\s*:\s*"([^"]+)"') { $foundPlan = $Matches[1] }

Say "id_token: email='$foundEmail'  plan='$foundPlan'  account_id='$foundAccount'"
$R.precond_email   = $foundEmail
$R.precond_plan    = $foundPlan
$R.precond_account = $foundAccount

# hard STOP on any Edu signal
if ($claimsRaw -match 'begb0037@ox\.ac\.uk' -or $foundPlan -eq 'education' -or $foundPlan -eq 'enterprise' -or $foundAccount -eq 'cc80356f-959e-449f-9721-add87a9ba0a5') {
  Fail "this CODEX_HOME is signed into the OXFORD EDU account (email='$foundEmail' plan='$foundPlan' account='$foundAccount'). Lane B must use the PERSONAL account. Run:  `$env:CODEX_HOME='$CodexHome'; codex login   as $Email, then re-run."
}
# positive assert: email must match, plan must be a personal plan (belt + braces)
if ($foundEmail.ToLower() -ne $Email.ToLower()) {
  Fail "id_token email is '$foundEmail', expected '$Email'. Not cutting over."
}
if ($foundPlan -eq '(unknown)') {
  Say "WARN: could not read chatgpt_plan_type from the id_token -- proceeding on the verified email ($foundEmail)"
} elseif (@('plus','pro','team') -notcontains $foundPlan) {
  Fail "id_token plan is '$foundPlan' -- expected a personal plan (plus / pro / team). Not cutting over."
}
Say "PRECONDITION: account identity OK -- $foundEmail / plan $foundPlan / account $foundAccount"

# --- probe: does the personal account have the Outlook Calendar connector? ---
Say "probing the connector via lane_b_call1.py (this makes ~1 real codex call, ~3-5 min on this box)..."
$t = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
foreach ($f in 'lane_b_call1.py', 'normalise_pull.py') {
  try { Invoke-WebRequest -UseBasicParsing "$base/$f`?t=$t" -OutFile (Join-Path $RunDir $f); Say "  refreshed $f" }
  catch { Fail "could not refresh $f from main ($($_.Exception.Message))" }
}
$env:CODEX_HOME            = $CodexHome
$env:WI_LANE_B_CODEX_HOME  = $CodexHome
$env:WI_LANE_B_RETRIES     = '1'
$env:WI_LANE_B_TIMEOUT     = '420'
$env:WI_LANE_B_SKIP_WARMUP = '1'
$env:PYTHONUTF8            = '1'
& python -u (Join-Path $RunDir 'lane_b_call1.py') --domain calendar 2>&1 | Tee-Object -FilePath $transcript -Append
Say "lane_b_call1.py exit $LASTEXITCODE"

$norm = Join-Path $RunDir 'data\lane_b\lane_b_normalised.json'
if (-not (Test-Path $norm)) { Fail "probe produced no $norm -- connector unavailable this run. If it keeps happening, add the Microsoft Outlook Calendar connector to your PERSONAL ChatGPT account (Settings -> Connectors/Apps), then re-run." }
$doc = $null
try { $doc = Get-Content $norm -Raw | ConvertFrom-Json } catch { Fail "could not parse $norm : $($_.Exception.Message)" }
$calStatus = "$($doc.meta.lane_b.domains.calendar.status)"
$calCount  = [int]$doc.meta.lane_b.domains.calendar.count
Say "probe: calendar status='$calStatus'  count=$calCount  tool_calls=$($doc.meta.lane_b.domains.calendar.tool_calls -join ', ')"
$R.precond_probe = "status=$calStatus count=$calCount"
if ($calStatus -eq 'halt') { Fail "the re-contamination guard HALTED on the probe (a write tool appeared). DO NOT cut over. See data\codex_runs\GUARD_TRIPPED_*." }
if ($calStatus -ne 'ok' -or $calCount -lt 1) {
  Fail "probe did not return real events (status='$calStatus', count=$calCount). The personal account most likely has NO Outlook Calendar connector attached -- add it in ChatGPT -> Settings -> Connectors/Apps, then re-run. (Teams already works on personal; Calendar is the one to add.)"
}
Say "PRECONDITION: connector calendar probe OK -- $calCount real events on the personal account"

# ============================================================================
# 2. BACKUP
# ============================================================================
Say "--- 2. BACKUP ---"
$task = $null
try { $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop } catch { Fail "scheduled task '$TaskName' not found on this box: $($_.Exception.Message)" }
$act0 = @($task.Actions)[0]
$exe0 = "$($act0.Execute)"
$args0 = "$($act0.Arguments)"
$wd0  = "$($act0.WorkingDirectory)"
Say "current action: $exe0 $args0"
Say "current workingdir: $wd0"

# locate the .ps1 the task actually runs
$livePs1 = $null
if     ($args0 -match '-File\s+"([^"]+\.ps1)"')       { $livePs1 = $Matches[1] }
elseif ($args0 -match "-File\s+'([^']+\.ps1)'")       { $livePs1 = $Matches[1] }
elseif ($args0 -match '-File\s+(\S+\.ps1)')           { $livePs1 = $Matches[1] }
elseif ($args0 -match '"([A-Za-z]:\\[^"]+\.ps1)"')    { $livePs1 = $Matches[1] }
elseif ($args0 -match "'([A-Za-z]:\\[^']+\.ps1)'")    { $livePs1 = $Matches[1] }
if (-not $livePs1 -and ($args0 -match '"?([A-Za-z]:\\[^"]+\.vbs)"?')) {
  $vbs = $Matches[1]
  if (Test-Path $vbs) {
    $vt = Get-Content $vbs -Raw
    if ($vt -match '([A-Za-z]:\\[^"'']+\.ps1)') { $livePs1 = $Matches[1] }
  }
}
if (-not $livePs1) {
  $livePs1 = Join-Path $env:USERPROFILE 'work-inbox\Run Laptop Bridge Briefing.ps1'
  Say "WARN: could not parse a .ps1 from the task action; defaulting to $livePs1"
}
if (-not (Test-Path $livePs1)) { Fail "live wrapper not found at $livePs1 -- re-run with the task action pointing at a real .ps1" }
$R.live_ps1 = $livePs1
Say "live wrapper script: $livePs1"

$scriptBak = "$livePs1.bak-$ts"
Copy-Item $livePs1 $scriptBak -Force
$R.script_bak = $scriptBak
Say "backed up wrapper -> $scriptBak"

$xmlBak = Join-Path $logdir ("WorkInboxBridgeBriefing_task_$ts.xml")
try {
  Export-ScheduledTask -TaskName $TaskName | Out-File -FilePath $xmlBak -Encoding Unicode
  $R.task_xml_bak = $xmlBak
  Say "exported task XML -> $xmlBak"
} catch { Fail "Export-ScheduledTask failed: $($_.Exception.Message)" }

# ============================================================================
# 3. FLIP  ($LaneBCodexHome in the wrapper, pulled fresh from main)
# ============================================================================
Say "--- 3. FLIP `$LaneBCodexHome ---"
$freshWrapper = Join-Path $logdir "Run Laptop Bridge Briefing.frommain.$ts.ps1"
try {
  Invoke-WebRequest -UseBasicParsing "$base/docs/desktop-scripts/Run%20Laptop%20Bridge%20Briefing.ps1`?t=$([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())" -OutFile $freshWrapper
  Say "pulled Run Laptop Bridge Briefing.ps1 fresh from main -> $freshWrapper"
} catch { Fail "could not pull Run Laptop Bridge Briefing.ps1 from main: $($_.Exception.Message)" }

$wtxt = Get-Content $freshWrapper -Raw
if ($wtxt -notmatch "(?m)^\s*\`$LaneBCodexHome\s*=") {
  Fail "the fresh wrapper from main has no `\`$LaneBCodexHome` assignment line -- its shape changed. Do NOT proceed; ping Drew."
}
$wtxt2 = [regex]::Replace($wtxt, "(?m)^(\s*\`$LaneBCodexHome\s*=\s*).*?$", "`$1'$CodexHome'")
if ($wtxt2 -notmatch [regex]::Escape("`$LaneBCodexHome = '$CodexHome'") -and $wtxt2 -notmatch [regex]::Escape("`$LaneBCodexHome='$CodexHome'")) {
  Fail "the `\`$LaneBCodexHome` flip did not apply cleanly -- aborting before touching the live wrapper."
}
Set-Content -Path $livePs1 -Value $wtxt2 -Encoding UTF8
Say "wrote patched wrapper (with `\`$LaneBCodexHome = '$CodexHome' and the guard glue) -> $livePs1"

# ============================================================================
# 4. RE-REGISTER  (Set-ScheduledTask -Action -- preserves triggers/principal/settings)
# ============================================================================
Say "--- 4. RE-REGISTER task action (+ -CalBackend connector) ---"
$newArgs = $args0
if ($newArgs -match '-CalBackend\s+connector') {
  Say "task action already has -CalBackend connector -- leaving args as-is (idempotent)"
}
elseif ($newArgs -match '-CalBackend\s+com') {
  $newArgs = $newArgs -replace '-CalBackend\s+com', '-CalBackend connector'
}
else {
  $newArgs = ($newArgs.TrimEnd() + ' -CalBackend connector')
}

$actSplat = @{ Execute = $exe0; Argument = $newArgs }
if ($wd0) { $actSplat.WorkingDirectory = $wd0 }
try {
  Set-ScheduledTask -TaskName $TaskName -Action (New-ScheduledTaskAction @actSplat) -ErrorAction Stop | Out-Null
} catch {
  Say "Set-ScheduledTask failed: $($_.Exception.Message)"
  Say "MANUAL FALLBACK: Register-ScheduledTask -Xml (Get-Content '$xmlBak' -Raw) -TaskName '$TaskName' -Force   (then edit the <Arguments> line to add ' -CalBackend connector')"
  Fail "task re-register failed -- nothing else changed except the wrapper script (restore with: Copy-Item '$scriptBak' '$livePs1' -Force)"
}

$after = @((Get-ScheduledTask -TaskName $TaskName).Actions)[0]
$R.new_action = "$($after.Execute) $($after.Arguments)"
Say "new action: $($R.new_action)"
if ($after.Arguments -notmatch '-CalBackend\s+connector') { Fail "re-register did not take -- action still: $($after.Arguments)" }

# triggers (read + print, do not change)
$trg = (Get-ScheduledTask -TaskName $TaskName).Triggers
Say "triggers (unchanged):"
$trg | ForEach-Object { Say ("  {0}  {1}" -f $_.CimClass.CimClassName, $_.StartBoundary) }

# ============================================================================
# 5. FORCE RUN + VERIFY
# ============================================================================
function Get-BriefingSha {
  try {
    $h = @{ 'User-Agent' = 'lane-b-cutover' }
    $r = Invoke-RestMethod -UseBasicParsing -Headers $h "https://api.github.com/repos/begb0037admin/work-inbox/commits?path=data/briefing.json&per_page=1"
    return "$($r[0].sha.Substring(0,10))  $($r[0].commit.committer.date)"
  } catch { return "(github api unavailable: $($_.Exception.Message))" }
}

if ($SkipRun) {
  Say "--- 5. FORCE RUN skipped (-SkipRun). The task is cut over; it will run at its next trigger. ---"
} else {
  Say "--- 5. FORCE RUN + VERIFY (up to $RunTimeoutMin min) ---"
  $R.briefing_before = Get-BriefingSha
  Say "briefing.json HEAD before: $($R.briefing_before)"
  Start-ScheduledTask -TaskName $TaskName
  Say "Start-ScheduledTask issued; polling State..."
  $deadline = (Get-Date).AddMinutes($RunTimeoutMin)
  do {
    Start-Sleep -Seconds 20
    $state = "$((Get-ScheduledTask -TaskName $TaskName).State)"
    Write-Host ("[{0}] state={1}" -f (Stamp), $state)
  } while ($state -eq 'Running' -and (Get-Date) -lt $deadline)
  $tinfo = Get-ScheduledTaskInfo -TaskName $TaskName
  Say "task finished: state=$state  LastTaskResult=$($tinfo.LastTaskResult)  LastRunTime=$($tinfo.LastRunTime)"
  if ($state -eq 'Running') { Say "WARN: still Running after $RunTimeoutMin min -- capturing what we have; check the bridge log + re-poll manually." }

  Start-Sleep -Seconds 5
  $latestLog = Get-ChildItem $logdir -Filter 'bridge_briefing_*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($latestLog) {
    Say "bridge log: $($latestLog.FullName)"
    $lt = Get-Content $latestLog.FullName
    $g = $lt | Select-String 'lane_b_cal_guard\.py exit' | Select-Object -Last 1
    $c = $lt | Select-String 'Lane B connector calendar:' | Select-Object -Last 1
    $f = $lt | Select-String 'fetch_inbox\.py exit' | Select-Object -Last 1
    $R.guard_exit    = if ($g) { "$($g.Line.Trim())" } else { '(not found in log)' }
    $R.calendar_line = if ($c) { "$($c.Line.Trim())" } else { '(not found in log)' }
    $R.fetch_exit    = if ($f) { "$($f.Line.Trim())" } else { '(not found in log)' }
    Say "  $($R.guard_exit)"
    Say "  $($R.calendar_line)"
    Say "  $($R.fetch_exit)"
  } else {
    Say "WARN: no bridge_briefing_*.log found -- did the task action run the wrapper?"
  }
  $R.briefing_after = Get-BriefingSha
  Say "briefing.json HEAD after:  $($R.briefing_after)"
  if ($R.briefing_after -ne $R.briefing_before -and $R.briefing_after -notmatch 'unavailable') {
    Say "=> a NEW briefing.json was pushed."
  } else {
    Say "=> briefing.json HEAD did not change -- check the fetch_inbox exit + the bridge log above (a Phase-4 safe-write veto, or the run is still finishing)."
  }
}

# ============================================================================
# 6. ROLLBACK (printed, NOT run)
# ============================================================================
$R.rollback_a      = "Register-ScheduledTask -Xml (Get-Content '$xmlBak' -Raw) -TaskName '$TaskName' -Force"
$R.rollback_b      = "`$a=(Get-ScheduledTask -TaskName '$TaskName').Actions[0]; Set-ScheduledTask -TaskName '$TaskName' -Action (New-ScheduledTaskAction -Execute `$a.Execute -Argument (`$a.Arguments -replace ' -CalBackend connector','') -WorkingDirectory `$a.WorkingDirectory)"
$R.rollback_script = "Copy-Item '$scriptBak' '$livePs1' -Force"

Say "--- 6. ROLLBACK (copy-paste to revert; you do NOT need all three) ---"
Say "  task, exact restore : $($R.rollback_a)"
Say "  task, just drop flag: $($R.rollback_b)"
Say "  wrapper script back : $($R.rollback_script)"

# ============================================================================
# 7. PASTE THIS BACK
# ============================================================================
Say "================ PASTE THIS BACK ================"
Say "host/user            : $env:COMPUTERNAME  $env:USERDOMAIN\$env:USERNAME"
Say "precond auth.json    : $($R.precond_auth)"
Say "precond account      : email=$($R.precond_email)  plan=$($R.precond_plan)  account_id=$($R.precond_account)"
Say "precond connector    : $($R.precond_probe)"
Say "wrapper backup       : $($R.script_bak)"
Say "task XML backup      : $($R.task_xml_bak)"
Say "live wrapper         : $($R.live_ps1)"
Say "new task action      : $($R.new_action)"
Say "guard exit line      : $($R.guard_exit)"
Say "calendar line        : $($R.calendar_line)"
Say "fetch_inbox exit     : $($R.fetch_exit)"
Say "briefing before      : $($R.briefing_before)"
Say "briefing after       : $($R.briefing_after)"
Say "ROLLBACK (task)      : $($R.rollback_a)"
Say "ROLLBACK (wrapper)   : $($R.rollback_script)"
Say "transcript           : $transcript"
Say "================================================"
if ($transcribing) { Stop-Transcript | Out-Null }
exit 0
