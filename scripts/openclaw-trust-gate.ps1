$ErrorActionPreference = "Stop"

Write-Host "== OpenClaw Tier-2 Trust Gate ==" -ForegroundColor Cyan

$allPassed = $true

function Invoke-Step {
  param(
    [string]$Name,
    [scriptblock]$Action
  )

  Write-Host "[run] $Name" -ForegroundColor Yellow
  try {
    $global:LASTEXITCODE = 0
    & $Action
    if ($LASTEXITCODE -ne 0) {
      throw "Exit code $LASTEXITCODE"
    }
    Write-Host "[ok] $Name" -ForegroundColor Green
  }
  catch {
    $script:allPassed = $false
    Write-Host "[fail] $Name :: $($_.Exception.Message)" -ForegroundColor Red
  }
}

Invoke-Step -Name "OpenClaw CLI available" -Action {
  openclaw --version | Out-Null
}

Invoke-Step -Name "Gateway health" -Action {
  openclaw gateway health | Out-Null
}

Invoke-Step -Name "Typecheck" -Action {
  npx tsc --noEmit
}

Invoke-Step -Name "Tests" -Action {
  npm test
}

Invoke-Step -Name "Autopilot dry-run" -Action {
  npm run openclaw:auto:dry
}

$runDir = Join-Path (Get-Location) "docs\phase-reports\autopilot-runs"
$slotLockPath = Join-Path $runDir "supervisor-slot.lock"
$skipQuickRun = $false

if (Test-Path $slotLockPath) {
  try {
    $slot = Get-Content $slotLockPath -Raw | ConvertFrom-Json
    if ($slot.pid) {
      $proc = Get-Process -Id $slot.pid -ErrorAction SilentlyContinue
      if ($proc) {
        $skipQuickRun = $true
        Write-Host "[info] Supervisor already running (PID $($slot.pid)); skipping quick run and validating heartbeat only." -ForegroundColor Cyan
      }
      else {
        Remove-Item $slotLockPath -Force -ErrorAction SilentlyContinue
      }
    }
  }
  catch {
    Remove-Item $slotLockPath -Force -ErrorAction SilentlyContinue
  }
}

if (-not $skipQuickRun) {
  Invoke-Step -Name "Supervisor quick run" -Action {
    npm run openclaw:supervisor:quick
  }
}

$heartbeatPath = Join-Path (Get-Location) "docs\phase-reports\autopilot-runs\supervisor-heartbeat.json"
if (Test-Path $heartbeatPath) {
  Write-Host "[run] Supervisor heartbeat check" -ForegroundColor Yellow
  try {
    $heartbeat = Get-Content $heartbeatPath -Raw | ConvertFrom-Json
    $status = [string]$heartbeat.status
    if ($status -eq "cycle-complete" -or $status -eq "gate-passed") {
      Write-Host "[ok] Supervisor heartbeat status is '$status'" -ForegroundColor Green
    }
    else {
      $allPassed = $false
      Write-Host "[fail] Supervisor heartbeat status is '$status' (expected cycle-complete or gate-passed)" -ForegroundColor Red
    }
  }
  catch {
    $allPassed = $false
    Write-Host "[fail] Could not parse supervisor heartbeat" -ForegroundColor Red
  }
}
else {
  $allPassed = $false
  Write-Host "[fail] Missing supervisor heartbeat file" -ForegroundColor Red
}

if ($allPassed) {
  Write-Host "READY_FOR_TIER2=true" -ForegroundColor Green
  exit 0
}

Write-Host "READY_FOR_TIER2=false" -ForegroundColor Red
exit 1
