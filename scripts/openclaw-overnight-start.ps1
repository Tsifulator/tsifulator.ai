param(
  [string]$SessionId = "tsif-phase1",
  [string]$PromptFile = ".\docs\openclaw-autopilot-prompt.md",
  [string]$AgentId = "main",
  [bool]$SingleControllerMode = $true
)

$ErrorActionPreference = "Continue"

$stateDir = Join-Path (Get-Location) "docs\phase-reports\autopilot-runs"
if (-not (Test-Path $stateDir)) {
  New-Item -ItemType Directory -Path $stateDir | Out-Null
}
$statePath = Join-Path $stateDir "overnight-state.json"

if (-not (Test-Path $PromptFile)) {
  Write-Host "[fail] Prompt file not found: $PromptFile" -ForegroundColor Red
  exit 1
}

function Test-Gateway {
  $result = & openclaw gateway health 2>&1 | Out-String
  return ($LASTEXITCODE -eq 0 -and $result -notmatch "error|failed|unreachable")
}

Write-Host "[info] Starting OpenClaw overnight mode..." -ForegroundColor Cyan

function Stop-StaleRunners {
  $runnerProcs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object {
      $_.Name -eq "powershell.exe" -and
      $_.CommandLine -match "openclaw-supervisor.ps1|openclaw-autopilot.ps1"
    }

  foreach ($proc in $runnerProcs) {
    try {
      Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop
      Write-Host "[warn] Stopped stale runner PID $($proc.ProcessId)" -ForegroundColor DarkYellow
    } catch {
      Write-Host "[warn] Could not stop stale runner PID $($proc.ProcessId)" -ForegroundColor DarkYellow
    }
  }
}

Stop-StaleRunners

if ($SingleControllerMode) {
  Write-Host "[info] Single-controller mode enabled. Stopping watchdog to avoid controller overlap..." -ForegroundColor Cyan
  & powershell -ExecutionPolicy Bypass -File ".\scripts\openclaw-watchdog-stop.ps1"
}

if (Test-Path $statePath) {
  Write-Host "[warn] Existing overnight run detected. Stopping previous run first..." -ForegroundColor DarkYellow
  & powershell -ExecutionPolicy Bypass -File ".\scripts\openclaw-overnight-stop.ps1"
}

if ($SessionId -eq "tsif-phase1") {
  $SessionId = "tsif-overnight-$(Get-Date -Format 'yyyyMMdd-HHmm')"
}

$gatewayPid = $null
if (-not (Test-Gateway)) {
  $gateway = Start-Process -FilePath "openclaw.cmd" -ArgumentList "gateway","run","--force" -WindowStyle Minimized -PassThru
  Start-Sleep -Seconds 6
  $gatewayPid = $gateway.Id
}

$supervisorArgs = @(
  "-NoProfile",
  "-ExecutionPolicy", "Bypass",
  "-File", ".\\scripts\\openclaw-supervisor.ps1",
  "-PromptFile", $PromptFile,
  "-AgentId", $AgentId,
  "-SessionId", $SessionId,
  "-Cycles", "999",
  "-TurnsPerCycle", "4",
  "-AgentTimeoutSeconds", "120",
  "-TurnDelaySeconds", "110",
  "-RateLimitCooldownSeconds", "420",
  "-CyclePauseSeconds", "150",
  "-CompactResumePrompt"
)
$supervisor = Start-Process -FilePath "powershell.exe" -ArgumentList $supervisorArgs -WindowStyle Minimized -PassThru

$state = @{
  startedAt = (Get-Date).ToString("o")
  sessionId = $SessionId
  agentId = $AgentId
  gatewayPid = $gatewayPid
  supervisorPid = $supervisor.Id
  statePath = $statePath
}
$state | ConvertTo-Json | Out-File -FilePath $statePath -Encoding utf8

Write-Host "[ok] Overnight mode started." -ForegroundColor Green
Write-Host "[info] Session ID: $SessionId" -ForegroundColor Green
Write-Host "[info] Agent ID: $AgentId" -ForegroundColor Green
if ($gatewayPid) {
  Write-Host "[info] Gateway PID: $gatewayPid" -ForegroundColor Green
} else {
  Write-Host "[info] Gateway already healthy; reusing existing instance." -ForegroundColor Green
}
Write-Host "[info] Supervisor PID: $($supervisor.Id)" -ForegroundColor Green
Write-Host "[info] State file: $statePath" -ForegroundColor Cyan
