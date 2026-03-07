param(
  [string]$TaskName = "TsifulatorOpenClawWatchdog",
  [string]$PromptFile = ".\docs\openclaw-cross-app-mvp-prompt.md",
  [string]$SessionId = "tsif-phase1",
  [string]$AgentId = "main",
  [int]$IntervalSeconds = 180,
  [int]$LockStaleMinutes = 2
)

$ErrorActionPreference = "Stop"

$workspace = (Get-Location).Path
$startScript = Join-Path $workspace "scripts\openclaw-watchdog-start.ps1"
if (-not (Test-Path $startScript)) {
  Write-Host "[fail] Missing script: $startScript" -ForegroundColor Red
  exit 1
}

$startupDir = [Environment]::GetFolderPath("Startup")
if (-not (Test-Path $startupDir)) {
  New-Item -ItemType Directory -Path $startupDir | Out-Null
}

$launcherPath = Join-Path $startupDir "TsifulatorOpenClawWatchdog.cmd"
$cmd = @"
@echo off
cd /d "$workspace"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\openclaw-watchdog-start.ps1" -PromptFile "$PromptFile" -SessionId "$SessionId" -AgentId "$AgentId" -IntervalSeconds $IntervalSeconds -LockStaleMinutes $LockStaleMinutes
"@

$cmd | Out-File -FilePath $launcherPath -Encoding ascii

Write-Host "[ok] Auto-start launcher created: $launcherPath" -ForegroundColor Green
Write-Host "[info] Starting watchdog now..." -ForegroundColor Cyan
& powershell -ExecutionPolicy Bypass -File ".\scripts\openclaw-watchdog-start.ps1" -PromptFile $PromptFile -SessionId $SessionId -AgentId $AgentId -IntervalSeconds $IntervalSeconds -LockStaleMinutes $LockStaleMinutes
Write-Host "[done] Auto-start enabled via Startup folder." -ForegroundColor Cyan
