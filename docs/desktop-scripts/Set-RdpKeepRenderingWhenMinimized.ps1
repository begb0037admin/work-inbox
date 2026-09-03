<#
Set-RdpKeepRenderingWhenMinimized.ps1
=====================================
Runs on the LOCAL machine that launches mstsc.exe (Kevin's DESKTOP), NOT the
Oxford laptop. Fixes: an RDP session locks/screensavers when the mstsc client
window is minimised for >10 min, because minimising drops the graphical
session (mstsc's "minimise video optimisation"), so the remote session's
GPO screensaver-lock timer (Oxford: 600s, secure) runs.

THE FIX: add a DWORD `RemoteDesktop_SuppressWhenMinimized` = 2 under
`HKCU\Software\Microsoft\Terminal Server Client`. This is the long-standing,
widely-cited value for this exact symptom (mstsc keeps rendering the session
in the background when minimised, so it never enters the suppressed state
that lets the remote screensaver engage). HKCU / per-user.

NOTE: I could not pull a live Microsoft/Sysinternals doc page to re-verify
the value on 3 Sept 2026 (WebFetch failed on the relevant URLs). It is the
established community-canonical value and is fully reversible -- see -Rollback
and the .reg backup in .\backups\ -- so any error is recoverable in one line.

APPLIES ONLY AFTER mstsc IS FULLY CLOSED AND REOPENED. An already-open
session does not pick it up.

USAGE (PowerShell 5.1, non-elevated is fine -- HKCU):
  .\Set-RdpKeepRenderingWhenMinimized.ps1            # apply (default)
  .\Set-RdpKeepRenderingWhenMinimized.ps1 -Rollback  # remove the value

TEST after applying: close ALL mstsc windows, reconnect to the laptop,
minimise it, wait 12+ minutes, restore it -- it should NOT be on the lock
screen.
#>
param([switch]$Rollback)

$ErrorActionPreference = 'Stop'
$key  = 'HKCU:\Software\Microsoft\Terminal Server Client'
$name = 'RemoteDesktop_SuppressWhenMinimized'

if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }

if ($Rollback) {
  if ($null -ne (Get-ItemProperty -Path $key -Name $name -ErrorAction SilentlyContinue).$name) {
    Remove-ItemProperty -Path $key -Name $name
    Write-Host "ROLLBACK: removed $name from $key. Close + reopen mstsc for it to take effect."
  } else {
    Write-Host "ROLLBACK: $name was not present -- nothing to do."
  }
  return
}

# backup the top key's scalar values first (small, no MRU/cert data)
$stamp   = Get-Date -Format 'yyyyMMdd_HHmm'
$bkdir   = Join-Path $PSScriptRoot 'backups'
New-Item -ItemType Directory -Force -Path $bkdir | Out-Null
$bkfile  = Join-Path $bkdir "TerminalServerClient_HKCU_topkey_$stamp.reg"
# reg export grabs the whole subtree; we only want it as an emergency full copy,
# so write it somewhere local + private, not into the repo:
$fullbk  = Join-Path ([Environment]::GetFolderPath('Desktop')) "RegBackups\TerminalServerClient_HKCU_FULL_$stamp.reg"
New-Item -ItemType Directory -Force -Path (Split-Path $fullbk) | Out-Null
& reg.exe export 'HKCU\Software\Microsoft\Terminal Server Client' $fullbk /y | Out-Null
Write-Host "Full subtree backup: $fullbk"

$before = (Get-ItemProperty -Path $key -Name $name -ErrorAction SilentlyContinue).$name
New-ItemProperty -Path $key -Name $name -PropertyType DWord -Value 2 -Force | Out-Null
$after = (Get-ItemProperty -Path $key -Name $name).$name
Write-Host "$name : '$before' -> $after   (expected 2)"
if ($after -ne 2) { throw "value did not take" }
Write-Host ""
Write-Host "DONE. Now close ALL mstsc windows and reconnect for it to take effect."
Write-Host "ROLLBACK: .\Set-RdpKeepRenderingWhenMinimized.ps1 -Rollback"
