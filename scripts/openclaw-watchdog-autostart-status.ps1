param(
  [string]$TaskName = "TsifulatorOpenClawWatchdog"
)

$ErrorActionPreference = "Continue"

$startupDir = [Environment]::GetFolderPath("Startup")
$launcherPath = Join-Path $startupDir "TsifulatorOpenClawWatchdog.cmd"

if (Test-Path $launcherPath) {
  Write-Host "[ok] Auto-start launcher present:" -ForegroundColor Green
  Write-Host $launcherPath
} else {
  Write-Host "[warn] Auto-start launcher not found in Startup folder." -ForegroundColor DarkYellow
}
