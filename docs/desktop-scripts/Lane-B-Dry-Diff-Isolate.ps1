# Lane B calendar guard — dry-diff isolation diagnostic
# Run in RDP as AD-OAK\begb0037 on 101L-DE013193, inside work-inbox repo.
$ErrorActionPreference = 'Stop'
Set-Location "$env:USERPROFILE\work-inbox"
Write-Host "[$(Get-Date -Format o)] Starting Lane B dry-diff isolation..."

python .\lane_b_cal_guard.py --snapshot --out data\codex_runs\test1.json
python .\lane_b_cal_guard.py --snapshot --out data\codex_runs\test2.json

$c1 = (Get-Content data\codex_runs\test1.json | ConvertFrom-Json).fp.PSObject.Properties.Name.Count
$c2 = (Get-Content data\codex_runs\test2.json | ConvertFrom-Json).fp.PSObject.Properties.Name.Count
Write-Host "test1 event count: $c1"
Write-Host "test2 event count: $c2"

Write-Host "--- test1 leave-event matches (first pull) ---"
Select-String -Path data\codex_runs\test1.json -Pattern "Annual Leave","Non-working day"

Write-Host "--- test2 leave-event matches (second pull) ---"
Select-String -Path data\codex_runs\test2.json -Pattern "Annual Leave","Non-working day"

Write-Host "--- formal diff ---"
python .\lane_b_cal_guard.py --diff --pre data\codex_runs\test1.json --post data\codex_runs\test2.json

Write-Host "[$(Get-Date -Format o)] Done. Copy everything above and send it back."
