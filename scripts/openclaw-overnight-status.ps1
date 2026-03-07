$ErrorActionPreference = "Continue"

$runDir = Join-Path (Get-Location) "docs\phase-reports\autopilot-runs"
$statePath = Join-Path $runDir "overnight-state.json"
$heartbeatPath = Join-Path $runDir "supervisor-heartbeat.json"

if (Test-Path $statePath) {
  $state = Get-Content $statePath -Raw | ConvertFrom-Json
  Write-Host "[info] Overnight state found" -ForegroundColor Cyan
  Write-Host "Session: $($state.sessionId)"
  Write-Host "Started: $($state.startedAt)"
  Write-Host "Gateway PID: $($state.gatewayPid)"
  Write-Host "Supervisor PID: $($state.supervisorPid)"

  foreach ($procId in @($state.gatewayPid, $state.supervisorPid)) {
    if ($procId) {
      $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
      if ($proc) {
        Write-Host "[ok] PID $procId running ($($proc.ProcessName))" -ForegroundColor Green
      } else {
        Write-Host "[warn] PID $procId not running" -ForegroundColor DarkYellow
      }
    }
  }
} else {
  Write-Host "[warn] Overnight state file missing." -ForegroundColor DarkYellow
}

if (Test-Path $heartbeatPath) {
  Write-Host "\n[info] Supervisor heartbeat:" -ForegroundColor Cyan
  Get-Content $heartbeatPath
} else {
  Write-Host "[warn] No supervisor heartbeat file yet." -ForegroundColor DarkYellow
}

$latest = $null
if ($state -and $state.sessionId) {
  $logs = Get-ChildItem $runDir -Filter "run-*.log" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
  foreach ($log in $logs) {
    $head = Get-Content $log.FullName -TotalCount 20 -ErrorAction SilentlyContinue | Out-String
    if ($head -match [regex]::Escape("SessionId: $($state.sessionId)")) {
      $latest = $log
      break
    }
  }
}

if (-not $latest) {
  $latest = Get-ChildItem $runDir -Filter "run-*.log" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}
if ($latest) {
  Write-Host "\n[info] Latest run log: $($latest.FullName)" -ForegroundColor Cyan
  Get-Content $latest.FullName -Tail 40
}
