<#
Watch-BridgeBriefing.ps1
=======================
BRIDGE-PERIOD desktop toast watcher. While the work-inbox pipeline runs on the
Oxford LAPTOP (desktop M365 device-registration broken), the desktop's own
"Work Inbox Briefing" task is disabled and fires no "briefing updated" toast.
This polls GitHub for a new data/briefing.json commit and fires one instead.

Read-only. No Outlook, no M365, no local pipeline dependency -- just an
authenticated GitHub commits API call + a BurntToast pop-up.

  - GET https://api.github.com/repos/begb0037admin/work-inbox/commits?path=data/briefing.json&per_page=1
  - if the newest commit's message starts with "chore: update briefing" AND its
    SHA differs from the last one we toasted -> toast + remember the SHA.
  - first run (no state file) seeds the SHA silently (no toast for a commit that
    pre-dates the watcher).

State:  %LOCALAPPDATA%\WorkInboxAI\bridge_toast_state.json
Log:    %LOCALAPPDATA%\WorkInboxAI\bridge_toast_watcher.log   (appended, timestamped)

TEMPORARY. Remove when the pipeline's permanent home + notification routing is
settled (that design is Markey's, not this). Unregister:
  Unregister-ScheduledTask -TaskName 'Work Inbox Briefing Toast Watcher' -Confirm:$false
#>
$ErrorActionPreference = 'Continue'

$stateDir = Join-Path $env:LOCALAPPDATA 'WorkInboxAI'
New-Item -ItemType Directory -Force -Path $stateDir | Out-Null
$stateFile = Join-Path $stateDir 'bridge_toast_state.json'
$logFile   = Join-Path $stateDir 'bridge_toast_watcher.log'

function Log($m) {
  try { Add-Content -LiteralPath $logFile -Value ("{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m) } catch {}
}

try {
  $pat = [Environment]::GetEnvironmentVariable('GITHUB_PAT','User')
  if (-not $pat) { $pat = [Environment]::GetEnvironmentVariable('GITHUB_PAT','Machine') }
  if (-not $pat) { $pat = $env:GITHUB_PAT }
  if (-not $pat) { Log 'no GITHUB_PAT available -- cannot poll. Exiting 0.'; exit 0 }

  $headers = @{
    Authorization = "Bearer $pat"
    'User-Agent'  = 'work-inbox-bridge-toast-watcher'
    Accept        = 'application/vnd.github+json'
  }
  $url = 'https://api.github.com/repos/begb0037admin/work-inbox/commits?path=data/briefing.json&per_page=1'
  $resp = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 30 -ErrorAction Stop
  if (-not $resp -or -not $resp[0].sha) { Log 'commits API returned nothing usable. Exiting 0.'; exit 0 }

  $sha     = [string]$resp[0].sha
  $msg     = [string]$resp[0].commit.message
  $isoDate = [string]$resp[0].commit.committer.date
  try { $when = ([DateTimeOffset]::Parse($isoDate)).ToLocalTime().ToString('ddd HH:mm') } catch { $when = $isoDate }

  $lastSha = $null
  if (Test-Path $stateFile) {
    try { $lastSha = (Get-Content -Raw -LiteralPath $stateFile | ConvertFrom-Json).lastSha } catch {}
  }

  if (-not $lastSha) {
    ($([ordered]@{ lastSha = $sha; seededAt = (Get-Date -Format 's') }) | ConvertTo-Json) | Set-Content -LiteralPath $stateFile
    Log "seeded (no toast) at $($sha.Substring(0,7))  msg=[$($msg.Split("`n")[0])]"
    exit 0
  }

  if ($sha -eq $lastSha) { Log "no change ($($sha.Substring(0,7)))"; exit 0 }

  if ($msg -notmatch '^chore: update briefing') {
    # a non-briefing commit touched briefing.json (rare) -- advance the pointer, don't toast
    ($([ordered]@{ lastSha = $sha; seededAt = (Get-Date -Format 's') }) | ConvertTo-Json) | Set-Content -LiteralPath $stateFile
    Log "new SHA $($sha.Substring(0,7)) but msg not a briefing update -- pointer advanced, no toast. msg=[$($msg.Split("`n")[0])]"
    exit 0
  }

  try {
    Import-Module BurntToast -ErrorAction Stop
    New-BurntToastNotification -Text @(
      'Work Inbox Briefing updated (bridge)',
      "$when  -  commit $($sha.Substring(0,7))",
      'Running on the laptop while the desktop is offline. Refresh the dashboard.'
    ) | Out-Null
    Log "TOASTED  $($sha.Substring(0,7))  ($when)"
  } catch {
    Log "toast failed: $($_.Exception.Message)"
  }

  ($([ordered]@{ lastSha = $sha; seededAt = (Get-Date -Format 's') }) | ConvertTo-Json) | Set-Content -LiteralPath $stateFile
  exit 0
}
catch {
  Log "ERROR (non-fatal): $($_.Exception.Message)"
  exit 0
}
