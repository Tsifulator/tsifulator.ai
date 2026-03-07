$ErrorActionPreference = "Continue"

$statePath = Join-Path (Get-Location) "docs\phase-reports\autopilot-runs\overnight-state.json"
if (-not (Test-Path $statePath)) {
  Write-Host "[warn] No overnight state file found." -ForegroundColor DarkYellow
  exit 0
}

$state = Get-Content $statePath -Raw | ConvertFrom-Json

foreach ($procId in @($state.supervisorPid, $state.gatewayPid)) {
  if ($procId) {
    try {
      Stop-Process -Id $procId -Force -ErrorAction Stop
      Write-Host "[ok] Stopped process PID $procId" -ForegroundColor Green
    } catch {
      Write-Host "[warn] Could not stop PID $procId (already stopped)." -ForegroundColor DarkYellow
    }
  }
}

Remove-Item $statePath -Force -ErrorAction SilentlyContinue
Write-Host "[done] Overnight mode stopped." -ForegroundColor Cyan
