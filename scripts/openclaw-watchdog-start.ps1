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
$statePath = Join-Path $runDir "watchdog-state.json"
$overnightStatePath = Join-Path $runDir "overnight-state.json"

if (Test-Path $overnightStatePath) {
  try {
    $overnight = Get-Content $overnightStatePath -Raw | ConvertFrom-Json
    if ($overnight.supervisorPid) {
      $supervisorProc = Get-Process -Id $overnight.supervisorPid -ErrorAction SilentlyContinue
      if ($supervisorProc) {
        Write-Host "[warn] Overnight supervisor is running (PID $($overnight.supervisorPid)); skipping watchdog start in single-controller mode." -ForegroundColor DarkYellow
        exit 0
      }
    }
  } catch {}
}

if (Test-Path $statePath) {
  try {
    $existing = Get-Content $statePath -Raw | ConvertFrom-Json
    if ($existing.pid) {
      $proc = Get-Process -Id $existing.pid -ErrorAction SilentlyContinue
      if ($proc) {
        Write-Host "[info] Watchdog already running (PID $($existing.pid))." -ForegroundColor Green
        exit 0
      }
    }
  } catch {}
}

$args = @(
  "-NoProfile",
  "-ExecutionPolicy", "Bypass",
  "-File", ".\\scripts\\openclaw-watchdog.ps1",
  "-PromptFile", $PromptFile,
  "-SessionId", $SessionId,
  "-AgentId", $AgentId,
  "-IntervalSeconds", "$IntervalSeconds",
  "-LockStaleMinutes", "$LockStaleMinutes"
)

$proc = Start-Process -FilePath "powershell.exe" -ArgumentList $args -WindowStyle Minimized -PassThru

@{
  startedAt = (Get-Date).ToString("o")
  pid = $proc.Id
  promptFile = $PromptFile
  sessionId = $SessionId
  agentId = $AgentId
  intervalSeconds = $IntervalSeconds
  lockStaleMinutes = $LockStaleMinutes
} | ConvertTo-Json | Out-File -FilePath $statePath -Encoding utf8

Write-Host "[ok] Watchdog started (PID $($proc.Id))." -ForegroundColor Green
Write-Host "[info] State file: $statePath" -ForegroundColor Cyan
