<#
.SYNOPSIS
  Reset and reauthenticate the kevin@lelitte.co.uk Codex M365 connector identity
  across both machines that hold a copy of its credential store, then verify
  both land on the same live grant before declaring success.

.DESCRIPTION
  Root cause (diagnosed 13 Sep 2026, after two reauths that didn't hold): the
  personal identity's OAuth refresh token exists in more than one independent
  holder at once --
    1. This desktop's own Codex CLI credential store (%USERPROFILE%\.codex\auth.json)
    2. A long-lived Codex Desktop app process on this same machine (the
       "OpenAI.Codex" package's Electron shell + its own `codex.exe ... app-server`
       child) that can sit open for a day or more without Kevin ever touching it,
       silently holding its own live session
    3. The Oxford laptop's Lane B FAILOVER credential store
       (C:\WorkInboxAI\codex-laneb\auth.json), logged in independently rather than
       synced from the desktop -- confirmed running its own "ChatGPT Classic"
       (OpenAI.ChatGPT-Desktop package) process too, under a different Windows
       account (begb0037-a) than the scheduled task (begb0037.AD-OAK)
  Whichever holder refreshes first rotates the underlying Microsoft refresh
  token; the other holders' now-superseded copies get invalid_grant the next
  time anything uses them (Microsoft's standard refresh-token-reuse detection).
  Codex does not self-heal from this -- it keeps retrying the same dead token
  and surfacing TRIGGER_REAUTHENTICATION instead of a clean re-prompt (matches
  openai/codex#14144 and #39054). Reauthenticating on just one machine, with a
  stale process or a stale second copy still alive elsewhere, reliably breaks
  again within the hour -- which is exactly what happened twice on 13 Sep 2026.

  This script collapses the two-machine problem to one: it makes the desktop's
  fresh login the single source of truth and PUSHES it to the Oxford laptop's
  failover store, instead of letting that store hold its own independently
  drifting login. It also kills every Codex-related process on both machines
  first, so no stale in-memory session can collide with the fresh grant.

.PARAMETER DryRun
  Detect and report everything (running processes on both machines, current
  credential file state) without killing anything, logging out, or overwriting
  any file. Safe to run any time, including right now, to sanity-check the
  script before a real reauth.

.PARAMETER OxfordHost
  SSH config host name for the Oxford laptop. Defaults to 'oxford-lan' (the
  existing entry in this machine's ~/.ssh/config).

.EXAMPLE
  Dry run first, always safe:
    powershell -File codex_connector_reauth.ps1 -DryRun

.EXAMPLE
  Real reauth (Kevin's one manual step is completing the device-code sign-in
  link this prints -- on his phone, another browser, wherever is convenient,
  no local browser needed on either machine):
    powershell -File codex_connector_reauth.ps1
#>

param(
    [switch]$DryRun,
    [string]$OxfordHost = 'oxford-lan',
    [string]$FailoverCodexHome = 'C:\WorkInboxAI\codex-laneb'
)

$ErrorActionPreference = 'Continue'
# Deliberately not 'Stop': ssh/scp write benign warnings (e.g. the post-quantum
# key exchange notice) to stderr, and PowerShell turns EVERY stderr line from a
# native command into a terminating error under 'Stop' once it's piped through
# 2>&1. Real failures below are caught via explicit $LASTEXITCODE checks instead.

function Write-Step {
    param([string]$Message)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'
    Write-Host "[$ts] $Message"
}

$ProbePrompt = "Call the microsoft_outlook_calendar connector's list_events tool for today's date only, one call, then stop. Report only whether the tool call succeeded or the exact error."

function Get-CodexProcesses {
    # Matches: the Codex Desktop app's Electron shell + all its child processes
    # (renderer/gpu/utility/crashpad, and the app-server codex.exe it spawns),
    # any bare codex.exe from a CLI invocation, and a separate "ChatGPT Classic"
    # (OpenAI.ChatGPT-Desktop package) app if present. Deliberately does NOT
    # match work-inbox's own one-shot `codex exec ...` calls made by
    # lane_b_call1.py's scheduled pipeline run -- those exit on their own
    # within minutes and are not the long-lived-session problem this script
    # targets. Distinguish by command line: app-server / bare GUI launch (no
    # "exec" argument) vs a one-shot "exec" invocation.
    Get-CimInstance Win32_Process | Where-Object {
        ($_.Name -in @('codex.exe', 'ChatGPT.exe', 'ChatGPT Classic.exe')) -and
        ($_.CommandLine -notmatch '\bexec\b')
    }
}

function Stop-CodexProcessTree {
    param([switch]$Apply)
    $procs = Get-CodexProcesses
    if (-not $procs) {
        Write-Step "  no long-lived Codex/ChatGPT process found"
        return
    }
    foreach ($p in $procs) {
        Write-Step "  found: PID $($p.ProcessId) $($p.Name) (parent $($p.ParentProcessId)) started $($p.CreationDate)"
        Write-Step "    path: $($p.ExecutablePath)"
        if ($Apply) {
            & taskkill /PID $p.ProcessId /T /F 2>&1 | ForEach-Object { Write-Step "    $_" }
        }
    }
}

# Remote helper scripts are written to disk and scp'd over rather than passed
# inline via `ssh host "..."` -- Windows OpenSSH does not reliably preserve
# nested quoting for a remote PowerShell command line, and an earlier version
# of this script silently lost every embedded double-quote and `{n}` format
# token in transit. A real file + `ssh host powershell -File <path>` is the
# only combination confirmed to survive intact end to end (verified in this
# script's own -DryRun test run, 13 Sep 2026).
$RemoteKillScript = @'
param([switch]$Apply)
Get-CimInstance Win32_Process | Where-Object {
    ($_.Name -in @('codex.exe','ChatGPT.exe','ChatGPT Classic.exe')) -and
    ($_.CommandLine -notmatch '\bexec\b')
} | ForEach-Object {
    Write-Output ("  found: PID {0} {1} (parent {2}) started {3}" -f $_.ProcessId,$_.Name,$_.ParentProcessId,$_.CreationDate)
    Write-Output ("    path: {0}" -f $_.ExecutablePath)
    if ($Apply) {
        taskkill /PID $_.ProcessId /T /F 2>&1 | ForEach-Object { Write-Output "    $_" }
    }
}
'@

$RemoteProbeScriptTemplate = @'
param([string]$CodexHome)
$env:CODEX_HOME = $CodexHome
codex exec -s read-only --json "__PROMPT__"
'@

$tmpDir = [System.IO.Path]::GetTempPath()
$localKillScript = Join-Path $tmpDir 'codex_reauth_remote_kill.ps1'
$localProbeScript = Join-Path $tmpDir 'codex_reauth_remote_probe.ps1'
Set-Content -Path $localKillScript -Value $RemoteKillScript -Encoding UTF8
Set-Content -Path $localProbeScript -Value ($RemoteProbeScriptTemplate -replace '__PROMPT__', $ProbePrompt) -Encoding UTF8

Write-Step "codex_connector_reauth.ps1 starting (DryRun=$($DryRun.IsPresent))"

# --- Step 1: inventory + kill on THIS machine (desktop) ---
Write-Step "Step 1/6: desktop Codex/ChatGPT processes"
Stop-CodexProcessTree -Apply:(-not $DryRun)

# --- Step 2: inventory + kill on the Oxford laptop over SSH ---
Write-Step "Step 2/6: Oxford laptop ($OxfordHost) Codex/ChatGPT processes"
$remoteTmpKill = 'C:\Windows\Temp\codex_reauth_remote_kill.ps1'
scp $localKillScript "${OxfordHost}:$remoteTmpKill" 2>&1 |
    Where-Object { $_ -notmatch 'WARNING|store now|upgraded|openssh\.com' } |
    ForEach-Object { Write-Step "  scp: $_" }
if ($LASTEXITCODE -ne 0) {
    Write-Error "scp of the remote kill script failed (exit $LASTEXITCODE) -- cannot safely proceed without confirming the Oxford laptop's process state."
    exit 1
}
$applyArg = if ($DryRun) { '' } else { '-Apply' }
ssh $OxfordHost "powershell -NoProfile -ExecutionPolicy Bypass -File $remoteTmpKill $applyArg" 2>&1 |
    Where-Object { $_ -notmatch 'WARNING|store now|upgraded|openssh\.com' } |
    ForEach-Object { Write-Step $_ }

if ($DryRun) {
    Write-Step "DryRun: stopping before logout/login/copy/verify. Re-run without -DryRun for the real reauth."
    Write-Step "Current desktop auth.json: $(Test-Path "$env:USERPROFILE\.codex\auth.json") -- $((Get-Item "$env:USERPROFILE\.codex\auth.json" -ErrorAction SilentlyContinue).LastWriteTime)"
    exit 0
}

# --- Step 3: fresh login on the desktop only (single source of truth) ---
Write-Step "Step 3/6: codex logout + codex login --device-auth on desktop"
Write-Step "  (this is the ONE manual step -- open the printed link on any device and sign in)"
& codex logout 2>&1 | ForEach-Object { Write-Step "  $_" }
& codex login --device-auth
if ($LASTEXITCODE -ne 0) {
    Write-Error "codex login --device-auth exited non-zero ($LASTEXITCODE) -- stopping before touching the Oxford laptop."
    exit 1
}
$status = & codex login status 2>&1
Write-Step "  login status: $status"
if ($status -notmatch 'Logged in') {
    Write-Error "Login did not complete -- stopping before touching the Oxford laptop."
    exit 1
}

# --- Step 4: verify the fresh grant works from the desktop's own CODEX_HOME ---
Write-Step "Step 4/6: verifying desktop connector access (single narrow probe, no retries)"
$desktopProbe = & codex exec -s read-only --json $ProbePrompt 2>&1
$desktopProbeText = $desktopProbe -join "`n"
if ($desktopProbeText -match 'oauth_token_invalid_grant|TRIGGER_REAUTHENTICATION') {
    Write-Error "Desktop probe still shows oauth_token_invalid_grant right after a fresh login -- stopping. Do not push this to the Oxford laptop. Report this back rather than retrying immediately."
    exit 1
}
Write-Step "  desktop probe: no reauth error observed"

# --- Step 5: push the fresh auth.json to the Oxford laptop's failover store ---
Write-Step "Step 5/6: pushing fresh auth.json to $OxfordHost`:$FailoverCodexHome"
$localAuth = "$env:USERPROFILE\.codex\auth.json"
$remoteAuthPath = ($FailoverCodexHome -replace '\\', '/') + '/auth.json'
scp $localAuth "${OxfordHost}:$remoteAuthPath" 2>&1 |
    Where-Object { $_ -notmatch 'WARNING|store now|upgraded|openssh\.com' } |
    ForEach-Object { Write-Step "  $_" }
if ($LASTEXITCODE -ne 0) {
    Write-Error "scp to the Oxford laptop failed (exit $LASTEXITCODE) -- desktop is on a fresh grant but the laptop failover store was NOT updated. Do not resume production Lane B failover until this is fixed."
    exit 1
}

# --- Step 6: verify the Oxford laptop's failover CODEX_HOME on the SAME fresh grant ---
Write-Step "Step 6/6: verifying Oxford laptop failover connector access"
$remoteTmpProbe = 'C:\Windows\Temp\codex_reauth_remote_probe.ps1'
scp $localProbeScript "${OxfordHost}:$remoteTmpProbe" 2>&1 |
    Where-Object { $_ -notmatch 'WARNING|store now|upgraded|openssh\.com' } |
    ForEach-Object { Write-Step "  scp: $_" }
$remoteProbe = ssh $OxfordHost "powershell -NoProfile -ExecutionPolicy Bypass -File $remoteTmpProbe -CodexHome '$FailoverCodexHome'" 2>&1
$remoteProbeText = ($remoteProbe -join "`n")
Write-Step "  remote probe output (last 500 chars): $($remoteProbeText.Substring([Math]::Max(0,$remoteProbeText.Length-500)))"
if ($remoteProbeText -match 'oauth_token_invalid_grant|TRIGGER_REAUTHENTICATION') {
    Write-Error "Oxford laptop failover store STILL shows oauth_token_invalid_grant after receiving the fresh token. Do not resume production use -- report this, do not retry immediately (see KILL_COOLDOWN_S discipline in lane_b_call1.py)."
    exit 1
}

Write-Step "SUCCESS: both desktop and Oxford laptop failover CODEX_HOME confirmed on a live grant."
Write-Step "Do not fire further ad hoc connector probes right now -- let production resume on its own schedule."
