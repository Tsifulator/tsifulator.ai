param(
  [string]$PromptFile = ".\docs\openclaw-autopilot-prompt.md",
  [string]$SessionId = "tsif-phase1",
  [string]$AgentId = "main",
  [int]$Cycles = 999,
  [int]$TurnsPerCycle = 60,
  [int]$AgentTimeoutSeconds = 90,
  [int]$TurnDelaySeconds = 25,
  [int]$RateLimitCooldownSeconds = 240,
  [int]$CyclePauseSeconds = 30,
  [int]$QueueWaitTimeoutSeconds = 600,
  [switch]$CompactResumePrompt,
  [switch]$RunGateEachCycle
)

$ErrorActionPreference = "Continue"

Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition @"
  [DllImport("kernel32.dll", SetLastError = true)]
  public static extern uint SetThreadExecutionState(uint esFlags);
"@

$ES_CONTINUOUS = [uint32]2147483648
$ES_SYSTEM_REQUIRED = [uint32]1
try {
  $null = [Win32.NativeMethods]::SetThreadExecutionState($ES_CONTINUOUS -bor $ES_SYSTEM_REQUIRED)
}
catch {
  Write-Host "[warn] Could not set execution state; continuing without sleep-prevention hint." -ForegroundColor DarkYellow
}

$statusDir = Join-Path (Get-Location) "docs\phase-reports\autopilot-runs"
if (-not (Test-Path $statusDir)) {
  New-Item -ItemType Directory -Path $statusDir | Out-Null
}
$heartbeatPath = Join-Path $statusDir "supervisor-heartbeat.json"
$supervisorLockPath = Join-Path $statusDir "supervisor-slot.lock"
$supervisorLockAcquired = $false

function Acquire-SupervisorSlot {
  param(
    [int]$TimeoutSeconds
  )

  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  while ((Get-Date) -lt $deadline) {
    if (-not (Test-Path $supervisorLockPath)) {
      @{ pid = $PID; startedAt = (Get-Date).ToString("o"); sessionId = $SessionId } | ConvertTo-Json | Out-File -FilePath $supervisorLockPath -Encoding utf8
      return $true
    }

    try {
      $existing = Get-Content $supervisorLockPath -Raw | ConvertFrom-Json
      if ($existing.pid) {
        $proc = Get-Process -Id $existing.pid -ErrorAction SilentlyContinue
        if (-not $proc) {
          Remove-Item $supervisorLockPath -Force -ErrorAction SilentlyContinue
          continue
        }
      }
    } catch {
      Remove-Item $supervisorLockPath -Force -ErrorAction SilentlyContinue
      continue
    }

    Start-Sleep -Seconds (Get-Random -Minimum 2 -Maximum 6)
  }

  return $false
}

function Release-SupervisorSlot {
  if (-not (Test-Path $supervisorLockPath)) {
    return
  }

  try {
    $existing = Get-Content $supervisorLockPath -Raw | ConvertFrom-Json
    if ($existing.pid -eq $PID) {
      Remove-Item $supervisorLockPath -Force -ErrorAction SilentlyContinue
    }
  } catch {
    Remove-Item $supervisorLockPath -Force -ErrorAction SilentlyContinue
  }
}

if (-not (Acquire-SupervisorSlot -TimeoutSeconds $QueueWaitTimeoutSeconds)) {
  Write-Host "[fail] Timed out waiting for supervisor slot after $QueueWaitTimeoutSeconds seconds." -ForegroundColor Red
  exit 1
}
$supervisorLockAcquired = $true

function Test-Gateway {
  $result = & openclaw gateway health 2>&1 | Out-String
  return ($LASTEXITCODE -eq 0 -and $result -notmatch "error|failed|unreachable")
}

function Invoke-Gate {
  param(
    [int]$CycleNumber
  )

  $phaseName = "autopilot-cycle-$CycleNumber"
  Write-Host "[gate] Running quality gate for $phaseName..." -ForegroundColor Yellow
  & npm run gate -- --Phase $phaseName
  return $LASTEXITCODE
}

Write-Host "[info] OpenClaw supervisor starting..." -ForegroundColor Cyan
Write-Host "[info] Cycles=$Cycles TurnsPerCycle=$TurnsPerCycle Session=$SessionId" -ForegroundColor Cyan

try {
for ($cycle = 1; $cycle -le $Cycles; $cycle++) {
  Write-Host "[cycle $cycle/$Cycles] checking gateway..." -ForegroundColor Yellow

  @{
    timestamp = (Get-Date).ToString("o")
    cycle = $cycle
    cycles = $Cycles
    sessionId = $SessionId
    status = "checking-gateway"
  } | ConvertTo-Json | Out-File -FilePath $heartbeatPath -Encoding utf8

  if (-not (Test-Gateway)) {
    Write-Host "[warn] Gateway health check failed. Attempting to start gateway in background..." -ForegroundColor DarkYellow
    Start-Process -FilePath "openclaw.cmd" -ArgumentList "gateway","run","--force" -WindowStyle Minimized | Out-Null
    Start-Sleep -Seconds 8
  }

  Write-Host "[cycle $cycle/$Cycles] running autopilot cycle..." -ForegroundColor Yellow

  @{
    timestamp = (Get-Date).ToString("o")
    cycle = $cycle
    cycles = $Cycles
    sessionId = $SessionId
    status = "running-autopilot"
  } | ConvertTo-Json | Out-File -FilePath $heartbeatPath -Encoding utf8
  $autopilotCommandArgs = @(
    "-ExecutionPolicy", "Bypass",
    "-File", ".\\scripts\\openclaw-autopilot.ps1",
    "-PromptFile", $PromptFile,
    "-AgentId", $AgentId,
    "-SessionId", $SessionId,
    "-MaxTurns", "$TurnsPerCycle",
    "-AgentTimeoutSeconds", "$AgentTimeoutSeconds",
    "-TurnDelaySeconds", "$TurnDelaySeconds",
    "-RateLimitCooldownSeconds", "$RateLimitCooldownSeconds",
    "-QueueWaitTimeoutSeconds", "$QueueWaitTimeoutSeconds"
  )

  if ($CompactResumePrompt) {
    $autopilotCommandArgs += "-CompactResumePrompt"
  }

  & powershell @autopilotCommandArgs

  $exitCode = $LASTEXITCODE
  if ($exitCode -ne 0) {
    Write-Host "[warn] Autopilot cycle exited with code $exitCode. Cooling down..." -ForegroundColor DarkYellow

    @{
      timestamp = (Get-Date).ToString("o")
      cycle = $cycle
      cycles = $Cycles
      sessionId = $SessionId
      status = "cooldown-after-error"
      exitCode = $exitCode
    } | ConvertTo-Json | Out-File -FilePath $heartbeatPath -Encoding utf8

    Start-Sleep -Seconds ([Math]::Max(60, $RateLimitCooldownSeconds))
  } else {
    Write-Host "[ok] Autopilot cycle completed." -ForegroundColor Green

    @{
      timestamp = (Get-Date).ToString("o")
      cycle = $cycle
      cycles = $Cycles
      sessionId = $SessionId
      status = "cycle-complete"
      exitCode = $exitCode
    } | ConvertTo-Json | Out-File -FilePath $heartbeatPath -Encoding utf8

    Start-Sleep -Seconds $CyclePauseSeconds
  }

  if ($RunGateEachCycle) {
    @{
      timestamp = (Get-Date).ToString("o")
      cycle = $cycle
      cycles = $Cycles
      sessionId = $SessionId
      status = "running-gate"
    } | ConvertTo-Json | Out-File -FilePath $heartbeatPath -Encoding utf8

    $gateExitCode = Invoke-Gate -CycleNumber $cycle
    if ($gateExitCode -ne 0) {
      Write-Host "[warn] Gate failed for cycle $cycle (exit $gateExitCode). Continuing loop..." -ForegroundColor DarkYellow
      @{
        timestamp = (Get-Date).ToString("o")
        cycle = $cycle
        cycles = $Cycles
        sessionId = $SessionId
        status = "gate-failed"
        gateExitCode = $gateExitCode
      } | ConvertTo-Json | Out-File -FilePath $heartbeatPath -Encoding utf8
    } else {
      Write-Host "[ok] Gate passed for cycle $cycle." -ForegroundColor Green
      @{
        timestamp = (Get-Date).ToString("o")
        cycle = $cycle
        cycles = $Cycles
        sessionId = $SessionId
        status = "gate-passed"
        gateExitCode = $gateExitCode
      } | ConvertTo-Json | Out-File -FilePath $heartbeatPath -Encoding utf8
    }
  }
}
}
finally {
  if ($supervisorLockAcquired) {
    Release-SupervisorSlot
  }
}

$null = [Win32.NativeMethods]::SetThreadExecutionState($ES_CONTINUOUS)

Write-Host "[done] Supervisor finished all cycles." -ForegroundColor Green
