param(
  [string]$TaskName = "TsifulatorOpenClawWatchdog"
)

$ErrorActionPreference = "Continue"

$startupDir = [Environment]::GetFolderPath("Startup")
$launcherPath = Join-Path $startupDir "TsifulatorOpenClawWatchdog.cmd"

if (-not (Test-Path $launcherPath)) {
  Write-Host "[warn] Auto-start launcher not found." -ForegroundColor DarkYellow
  exit 0
}

Remove-Item $launcherPath -Force -ErrorAction SilentlyContinue
Write-Host "[done] Auto-start disabled via Startup folder launcher removal." -ForegroundColor Cyan
