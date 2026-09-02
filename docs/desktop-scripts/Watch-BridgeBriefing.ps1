<#
Watch-BridgeBriefing.ps1
=======================
BRIDGE / MIGRATION-PERIOD desktop toast watcher. While work-inbox runs on the
Oxford LAPTOP (desktop M365 device-registration broken), the desktop's own
scheduled tasks are disabled and fire no toasts. This polls GitHub every few
minutes and toasts for:

  1. BRIEFING OK    -- a new "chore: update briefing" commit on data/briefing.json
  2. BRIEFING FAIL  -- data/laptop_status/briefing_status.json flips to result=failed
  3. DRAFT DIFF     -- data/laptop_status/draftdiff_status.json changes:
                       result=ok    -> "ran (pairs N / backlog M)"
                       result=failed-> "FAILED (exit N)"
  4. LANE B GUARD HALT -- data/laptop_status/briefing_status.json's lane_b_guard
                       field ='halted' (added 2 Sept 2026). Fires independently of
                       #2 -- a guard trip falls back to CAL_BACKEND=com and the
                       briefing still ships (result=ok), so #2 alone would miss it.

Read-only GitHub API only (commits API for #1, contents API for #2/#3). No
Outlook, no M365, no local pipeline dependency.

State:  %LOCALAPPDATA%\WorkInboxAI\bridge_toast_state.json
Log:    %LOCALAPPDATA%\WorkInboxAI\bridge_toast_watcher.log   (appended, timestamped)

TEMPORARY. This is deliberately a dumb poll+toast -- just "don't regress
notifications while the pipeline is migrating". The proper consolidated
notification design (briefing / draft-diff / failure routing across machines)
is queued for MARKEY, not this. Unregister:
  Unregister-ScheduledTask -TaskName 'Work Inbox Briefing Toast Watcher' -Confirm:$false
#>
$ErrorActionPreference = 'Continue'
$repo = 'begb0037admin/work-inbox'

$stateDir = Join-Path $env:LOCALAPPDATA 'WorkInboxAI'
New-Item -ItemType Directory -Force -Path $stateDir | Out-Null
$stateFile = Join-Path $stateDir 'bridge_toast_state.json'
$logFile   = Join-Path $stateDir 'bridge_toast_watcher.log'

function Log($m) {
  try { Add-Content -LiteralPath $logFile -Value ("{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m) } catch {}
}

function Toast($title, $line2, $line3) {
  try {
    Import-Module BurntToast -ErrorAction Stop
    New-BurntToastNotification -Text @($title, $line2, $line3) | Out-Null
    Log "TOASTED  [$title] $line2"
  } catch {
    Log "toast failed: $($_.Exception.Message)"
  }
}

# --- load state ---
$state = [ordered]@{ lastSha = $null; lastBriefingStatusSha = $null; lastDraftDiffSha = $null }
if (Test-Path $stateFile) {
  try {
    $j = Get-Content -Raw -LiteralPath $stateFile | ConvertFrom-Json
    foreach ($k in @('lastSha','lastBriefingStatusSha','lastDraftDiffSha')) {
      if ($j.PSObject.Properties.Name -contains $k) { $state[$k] = $j.$k }
    }
  } catch {}
}
function Save-State { ($state | ConvertTo-Json) | Set-Content -LiteralPath $stateFile }

try {
  $pat = [Environment]::GetEnvironmentVariable('GITHUB_PAT','User')
  if (-not $pat) { $pat = [Environment]::GetEnvironmentVariable('GITHUB_PAT','Machine') }
  if (-not $pat) { $pat = $env:GITHUB_PAT }
  if (-not $pat) { Log 'no GITHUB_PAT available -- cannot poll. Exiting 0.'; exit 0 }
  $pat = $pat.Trim()

  $headers = @{
    Authorization = "Bearer $pat"
    'User-Agent'  = 'work-inbox-bridge-toast-watcher'
    Accept        = 'application/vnd.github+json'
  }

  # ------------------------------------------------------------------ #
  #  1. BRIEFING OK -- newest commit on data/briefing.json
  # ------------------------------------------------------------------ #
  try {
    $url  = "https://api.github.com/repos/$repo/commits?path=data/briefing.json&per_page=1"
    $resp = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 30 -ErrorAction Stop
    if ($resp -and $resp[0].sha) {
      $sha     = [string]$resp[0].sha
      $msg     = [string]$resp[0].commit.message
      try { $when = ([DateTimeOffset]::Parse([string]$resp[0].commit.committer.date)).ToLocalTime().ToString('ddd HH:mm') } catch { $when = '' }
      if (-not $state.lastSha) {
        $state.lastSha = $sha; Save-State
        Log "seeded briefing pointer (no toast) at $($sha.Substring(0,7))"
      } elseif ($sha -ne $state.lastSha) {
        if ($msg -match '^chore: update briefing') {
          Toast 'Work Inbox Briefing updated (laptop)' "$when  -  commit $($sha.Substring(0,7))" 'Running on the laptop while the desktop is offline. Refresh the dashboard.'
        } else {
          Log "briefing.json new SHA $($sha.Substring(0,7)) but msg not a briefing update -- pointer advanced, no toast."
        }
        $state.lastSha = $sha; Save-State
      } else {
        Log "briefing: no change ($($sha.Substring(0,7)))"
      }
    }
  } catch { Log "briefing-commit poll error (non-fatal): $($_.Exception.Message)" }

  # ------------------------------------------------------------------ #
  #  helper: GET a contents-API file -> @{ sha=..; obj=<parsed json> } or $null (incl. 404)
  # ------------------------------------------------------------------ #
  function Get-StatusFile($path) {
    try {
      $u = "https://api.github.com/repos/$repo/contents/$path`?ref=main"
      $r = Invoke-RestMethod -Uri $u -Headers $headers -TimeoutSec 30 -ErrorAction Stop
      if (-not $r.sha) { return $null }
      $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(($r.content -replace '\s','')))
      return @{ sha = [string]$r.sha; obj = ($json | ConvertFrom-Json) }
    } catch {
      $code = $null
      if ($_.Exception.Response) { $code = $_.Exception.Response.StatusCode.value__ }
      if ($code -ne 404) { Log "contents poll $path error (non-fatal): $($_.Exception.Message)" }
      return $null
    }
  }

  # ------------------------------------------------------------------ #
  #  2. BRIEFING FAIL -- data/laptop_status/briefing_status.json
  # ------------------------------------------------------------------ #
  $bs = Get-StatusFile 'data/laptop_status/briefing_status.json'
  if ($bs) {
    if (-not $state.lastBriefingStatusSha) {
      $state.lastBriefingStatusSha = $bs.sha; Save-State
      Log "seeded briefing_status pointer (no toast) at $($bs.sha.Substring(0,7))"
    } elseif ($bs.sha -ne $state.lastBriefingStatusSha) {
      if ([string]$bs.obj.result -eq 'failed') {
        $ts = [string]$bs.obj.ts
        Toast 'Work Inbox Briefing - FAILED (laptop)' "exit $($bs.obj.exit_code)  -  $ts" 'The laptop briefing run did not complete. Check logs\bridge_briefing_last_run.log on the laptop.'
      } else {
        Log "briefing_status new ($($bs.sha.Substring(0,7))) result=$($bs.obj.result) -- no toast (success covered by the commit poll)."
      }
      # Lane B calendar-guard HALT -- fires independently of $bs.obj.result, because
      # a guard trip falls back to CAL_BACKEND=com and the briefing still ships
      # (result=ok) that cycle. This is the one that must be hard to miss even away
      # from the laptop: it means a scheduled task got disabled and Lane B needs
      # investigating. Added 2 Sept 2026 (Kevin: laptop toast alone is not enough
      # since the whole point is being away from the laptop).
      if ([string]$bs.obj.lane_b_guard -eq 'halted') {
        $ts = [string]$bs.obj.ts
        Toast 'Work Inbox - Lane B calendar guard HALTED (laptop)' "$ts" "$([string]$bs.obj.lane_b_guard_detail)"
      }
      $state.lastBriefingStatusSha = $bs.sha; Save-State
    } else {
      Log "briefing_status: no change ($($bs.sha.Substring(0,7)))"
    }
  }

  # ------------------------------------------------------------------ #
  #  3. DRAFT DIFF -- data/laptop_status/draftdiff_status.json
  # ------------------------------------------------------------------ #
  $ds = Get-StatusFile 'data/laptop_status/draftdiff_status.json'
  if ($ds) {
    if (-not $state.lastDraftDiffSha) {
      $state.lastDraftDiffSha = $ds.sha; Save-State
      Log "seeded draftdiff_status pointer (no toast) at $($ds.sha.Substring(0,7))"
    } elseif ($ds.sha -ne $state.lastDraftDiffSha) {
      $o = $ds.obj
      if ([string]$o.result -eq 'failed') {
        Toast 'Work Inbox Draft Diff - FAILED (laptop)' "exit $($o.exit_code)  -  $([string]$o.ts)" 'The laptop draft/final diff capture did not complete. Check logs\draft_diff_last_run.log on the laptop.'
      } else {
        function _st($n) { if ($o.stats -and ($o.stats.PSObject.Properties.Name -contains $n)) { $o.stats.$n } else { '?' } }
        $van = _st 'drafts_vanished_since_last_run'
        $pairs = _st 'draft_final_pairs_found'
        $cls = _st 'pairs_classified_this_run'
        $bk = _st 'backlog_size_after_this_run'
        Toast 'Work Inbox Draft Diff ran (laptop)' "vanished $van  /  pairs $pairs  /  classified $cls  /  backlog $bk" "$([string]$o.ts) - correlated+redacted pairs staged locally; enrichment via claude -p."
      }
      $state.lastDraftDiffSha = $ds.sha; Save-State
    } else {
      Log "draftdiff_status: no change ($($ds.sha.Substring(0,7)))"
    }
  }

  exit 0
}
catch {
  Log "ERROR (non-fatal): $($_.Exception.Message)"
  exit 0
}
