$ErrorActionPreference = "Continue"

$runDir = Join-Path (Get-Location) "docs\phase-reports\autopilot-runs"
$statePath = Join-Path $runDir "watchdog-state.json"
$heartbeatPath = Join-Path $runDir "watchdog-heartbeat.json"

if (Test-Path $statePath) {
  $state = Get-Content $statePath -Raw | ConvertFrom-Json
  Write-Host "[info] Watchdog state found" -ForegroundColor Cyan
  Write-Host "PID: $($state.pid)"
  Write-Host "Started: $($state.startedAt)"
  Write-Host "Session: $($state.sessionId)"
  Write-Host "Agent: $($state.agentId)"
  Write-Host "IntervalSeconds: $($state.intervalSeconds)"

  if ($state.pid) {
    $proc = Get-Process -Id $state.pid -ErrorAction SilentlyContinue
    if ($proc) {
      Write-Host "[ok] Watchdog process running" -ForegroundColor Green
    } else {
      Write-Host "[warn] Watchdog PID not running" -ForegroundColor DarkYellow
    }
  }
} else {
  Write-Host "[warn] No watchdog state file found." -ForegroundColor DarkYellow
}

if (Test-Path $heartbeatPath) {
  Write-Host "\n[info] Watchdog heartbeat:" -ForegroundColor Cyan
  Get-Content $heartbeatPath
} else {
  Write-Host "[warn] No watchdog heartbeat yet." -ForegroundColor DarkYellow
}
