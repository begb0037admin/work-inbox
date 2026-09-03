<#
Push-LaptopRunStatus.ps1
========================
Publishes a TINY run-status file to GitHub so the desktop toast watcher can
surface success AND failure of a work-inbox run that now happens on the laptop.

  data/laptop_status/briefing_status.json   (from "Run Laptop Bridge Briefing.ps1")
  data/laptop_status/draftdiff_status.json  (from "Run Laptop Draft Diff.ps1")

CONTENT PUSHED  -- counts / exit code / timestamp ONLY. Never any email body,
subject, sender, or recipient. For -Kind draftdiff, -StatsFile is the JSON the
Python script writes via --stats-out (its own stats dict: pure counts + paths).
This script copies through only a whitelisted set of those keys.

Read-only except this one file. Uses GITHUB_PAT (User env var, then Machine,
then process). Best-effort: any failure just logs to stderr and exits 0 -- a
status push must never fail a real run.

-LaneBGuard / -LaneBGuardDetail (added 2 Sept 2026, -Kind briefing only): the
Lane B calendar-guard outcome for this run ('not-run'|'clean'|'halted'|
'transient'|'unexpected-<n>') + a short detail string, so the desktop toast
watcher can surface a guard HALT even when the overall run exit code is 0 (the
guard falls back to CAL_BACKEND=com and the briefing still ships that cycle --
exit code alone would hide the trip). Still just a short fixed-vocabulary
string -- no email/calendar content.

-LaneBDomains (added 3 Sept 2026, regression-fix verification): a small
hashtable {domain -> {status,count,served_by,primary_failover_identical}},
pulled by the wrapper from the freshest data\lane_b\*_lane_b.json run log.
PURE COUNTS/STATUS/IDENTITY-LABEL ONLY -- never raw_items, never tool call
arguments/results, never any calendar/Teams content. This exists so the Lane B
both-domain regression fix (2 Sept incident -> 3 Sept fix) can be verified
against the LIVE scheduled task's own real unattended runs without needing
host/RDP access -- this file already round-trips through GitHub for exactly
that kind of cross-machine visibility.

USAGE
  Push-LaptopRunStatus.ps1 -Kind briefing  -ExitCode 0
  Push-LaptopRunStatus.ps1 -Kind briefing  -ExitCode 0 -LaneBGuard halted -LaneBGuardDetail "task disabled"
  Push-LaptopRunStatus.ps1 -Kind briefing  -ExitCode 0 -LaneBGuard clean -LaneBDomains @{calendar=@{status='ok';count=51;served_by='primary';primary_failover_identical=$false}}
  Push-LaptopRunStatus.ps1 -Kind draftdiff -ExitCode 1 -StatsFile C:\...\stats.json
  Push-LaptopRunStatus.ps1 -Kind draftdiff -ExitCode 3 -Note 'WithAI requested, no key'
#>
param(
  [Parameter(Mandatory)] [ValidateSet('briefing','draftdiff')] [string]$Kind,
  [Parameter(Mandatory)] [int]$ExitCode,
  [string]$StatsFile,
  [string]$Note,
  [string]$LaneBGuard,
  [string]$LaneBGuardDetail,
  [hashtable]$LaneBDomains
)

$ErrorActionPreference = 'Stop'
$repo = 'begb0037admin/work-inbox'
$path = "data/laptop_status/${Kind}_status.json"

function Warn($m) { Write-Host "Push-LaptopRunStatus WARN: $m" }

try {
  $pat = [Environment]::GetEnvironmentVariable('GITHUB_PAT','User')
  if (-not $pat) { $pat = [Environment]::GetEnvironmentVariable('GITHUB_PAT','Machine') }
  if (-not $pat) { $pat = $env:GITHUB_PAT }
  if (-not $pat) { Warn 'no GITHUB_PAT -- cannot push status. Exiting 0.'; exit 0 }
  $pat = $pat.Trim()
  if (-not $pat) { Warn 'GITHUB_PAT empty after trim. Exiting 0.'; exit 0 }

  $result = if ($ExitCode -eq 0) { 'ok' } else { 'failed' }
  $body = [ordered]@{
    kind      = $Kind
    result    = $result
    exit_code = $ExitCode
    ts        = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ssK')
    host      = $env:COMPUTERNAME
    user      = "$env:USERDOMAIN\$env:USERNAME"
  }
  if ($Note) { $body.note = $Note }
  if ($LaneBGuard) { $body.lane_b_guard = $LaneBGuard }
  if ($LaneBGuardDetail) { $body.lane_b_guard_detail = $LaneBGuardDetail }
  if ($LaneBDomains -and $LaneBDomains.Count -gt 0) {
    # Strict whitelist -- status/count/served_by/primary_failover_identical only,
    # per-key, even if the caller's hashtable happened to carry more. No content
    # ever passes through this path.
    $domainsOut = [ordered]@{}
    foreach ($d in $LaneBDomains.Keys) {
      $src = $LaneBDomains[$d]
      $domainsOut[$d] = [ordered]@{
        status                       = $src.status
        count                        = $src.count
        served_by                   = $src.served_by
        primary_failover_identical  = $src.primary_failover_identical
      }
    }
    $body.lane_b_domains = $domainsOut
  }

  if ($Kind -eq 'draftdiff' -and $StatsFile -and (Test-Path $StatsFile)) {
    try {
      $s = Get-Content -Raw -LiteralPath $StatsFile | ConvertFrom-Json
      # whitelist -- pure counts / flags / paths, no email content is possible here
      $keep = 'mail_backend','ai_backend','drafts_tracked_now','drafts_vanished_since_last_run',
              'abandoned_or_discarded','draft_final_pairs_found','pairs_excluded_by_redaction',
              'pairs_classified_this_run','pairs_permanently_failed','backlog_size_after_this_run',
              'ai_unavailable_this_run','total_diffs_accumulated','window_hours','run_time','error'
      $stats = [ordered]@{}
      foreach ($k in $keep) {
        if ($s.PSObject.Properties.Name -contains $k) { $stats[$k] = $s.$k }
      }
      $body.stats = $stats
    } catch {
      Warn "could not parse StatsFile ($($_.Exception.Message)) -- pushing status without stats"
    }
  }

  $json    = ($body | ConvertTo-Json -Depth 6)
  $headers = @{
    Authorization = "token $pat"
    'User-Agent'  = 'work-inbox-laptop-status'
    Accept        = 'application/vnd.github+json'
  }
  $apiUrl = "https://api.github.com/repos/$repo/contents/$path"

  $sha = $null
  try {
    $cur = Invoke-RestMethod -Uri "$apiUrl`?ref=main" -Headers $headers -TimeoutSec 30 -ErrorAction Stop
    if ($cur.sha) { $sha = [string]$cur.sha }
  } catch {
    if ($_.Exception.Response -and $_.Exception.Response.StatusCode.value__ -ne 404) {
      Warn "GET current status file failed: $($_.Exception.Message)"
    }
  }

  $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
  $payload = @{
    message = "chore: laptop $Kind run-status ($result, exit $ExitCode) $((Get-Date).ToString('yyyy-MM-dd HH:mm'))"
    content = $b64
    branch  = 'main'
  }
  if ($sha) { $payload.sha = $sha }

  $resp = Invoke-RestMethod -Uri $apiUrl -Method Put -Headers $headers `
            -Body ($payload | ConvertTo-Json -Depth 6) -ContentType 'application/json' `
            -TimeoutSec 30 -ErrorAction Stop
  Write-Host "Push-LaptopRunStatus: pushed $path ($result, exit $ExitCode) commit $($resp.commit.sha.Substring(0,7))"
  exit 0
}
catch {
  Warn "push failed (non-fatal): $($_.Exception.Message)"
  exit 0
}
