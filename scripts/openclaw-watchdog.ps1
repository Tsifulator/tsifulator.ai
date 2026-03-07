param(
  [string]$PromptFile = ".\docs\openclaw-cross-app-mvp-prompt.md",
  [string]$SessionId = "tsif-phase1",
  [string]$AgentId = "main",
  [int]$IntervalSeconds = 180,
  [int]$LockStaleMinutes = 2
)

$ErrorActionPreference = "Continue"

$runDir = Join-Path (Get-Location) "docs\phase-reports\autopilot-runs"
if (-not (Test-Path $runDir)) {
  New-Item -ItemType Directory -Path $runDir | Out-Null
}

$heartbeatPath = Join-Path $runDir "watchdog-heartbeat.json"
$statePath = Join-Path $runDir "overnight-state.json"

function Write-Heartbeat {
  param(
    [string]$Status,
    [string]$Detail = ""
  )

  @{
    timestamp = (Get-Date).ToString("o")
    status = $Status
    detail = $Detail
    sessionId = $SessionId
    agentId = $AgentId
    intervalSeconds = $IntervalSeconds
  } | ConvertTo-Json | Out-File -FilePath $heartbeatPath -Encoding utf8
}

function Clear-StaleLocks {
  $sessionsPath = Join-Path $env:USERPROFILE ".openclaw\agents\$AgentId\sessions"
  if (-not (Test-Path $sessionsPath)) {
    return
  }

  $locks = Get-ChildItem $sessionsPath -Filter "*.jsonl.lock" -ErrorAction SilentlyContinue
  foreach ($lock in $locks) {
    $ageMinutes = ((Get-Date) - $lock.LastWriteTime).TotalMinutes
    if ($ageMinutes -ge $LockStaleMinutes) {
      Remove-Item $lock.FullName -Force -ErrorAction SilentlyContinue
      Write-Host "[watchdog] removed stale lock: $($lock.Name)" -ForegroundColor DarkYellow
    }
  }
}

function Test-SupervisorRunning {
  if (-not (Test-Path $statePath)) {
    return $false
  }

  try {
    $state = Get-Content $statePath -Raw | ConvertFrom-Json
    if (-not $state.supervisorPid) {
      return $false
    }

    $proc = Get-Process -Id $state.supervisorPid -ErrorAction SilentlyContinue
    return [bool]$proc
  }
  catch {
    return $false
  }
}

Write-Host "[watchdog] started. interval=${IntervalSeconds}s lockStale=${LockStaleMinutes}m" -ForegroundColor Cyan

while ($true) {
  Clear-StaleLocks

  if (-not (Test-SupervisorRunning)) {
    Write-Host "[watchdog] supervisor missing; restarting overnight run..." -ForegroundColor Yellow
    Write-Heartbeat -Status "restarting-overnight" -Detail "Supervisor not running"

    & powershell -ExecutionPolicy Bypass -File ".\scripts\openclaw-overnight-start.ps1" -SessionId $SessionId -AgentId $AgentId -PromptFile $PromptFile

    Start-Sleep -Seconds 8
    if (Test-SupervisorRunning) {
      Write-Heartbeat -Status "running" -Detail "Overnight restarted"
    }
    else {
      Write-Heartbeat -Status "error" -Detail "Restart attempted but supervisor still missing"
    }
  }
  else {
    Write-Heartbeat -Status "running" -Detail "Supervisor healthy"
  }

  Start-Sleep -Seconds $IntervalSeconds
}
