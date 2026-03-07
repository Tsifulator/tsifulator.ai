$ErrorActionPreference = "Continue"

$statePath = Join-Path (Get-Location) "docs\phase-reports\autopilot-runs\watchdog-state.json"
if (-not (Test-Path $statePath)) {
  Write-Host "[warn] No watchdog state file found." -ForegroundColor DarkYellow
  exit 0
}

try {
  $state = Get-Content $statePath -Raw | ConvertFrom-Json
  if ($state.pid) {
    Stop-Process -Id $state.pid -Force -ErrorAction SilentlyContinue
    Write-Host "[ok] Stopped watchdog PID $($state.pid)" -ForegroundColor Green
  }
} catch {
  Write-Host "[warn] Could not parse watchdog state." -ForegroundColor DarkYellow
}

Remove-Item $statePath -Force -ErrorAction SilentlyContinue
Write-Host "[done] Watchdog stopped." -ForegroundColor Cyan
