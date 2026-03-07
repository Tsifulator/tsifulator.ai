param(
  [string]$Date,
  [string]$Owner = "",
  [switch]$Force
)

$ErrorActionPreference = "Stop"

$workspace = Get-Location
$templatePath = Join-Path $workspace "docs\daily-beta-triage-template.md"
if (-not (Test-Path $templatePath)) {
  throw "Missing template file: $templatePath"
}

if (-not $Date) {
  $Date = Get-Date -Format "yyyy-MM-dd"
}

if ($Date -notmatch "^\d{4}-\d{2}-\d{2}$") {
  throw "Date must be yyyy-MM-dd"
}

$triageDir = Join-Path $workspace "docs\daily-triage"
if (-not (Test-Path $triageDir)) {
  New-Item -ItemType Directory -Path $triageDir | Out-Null
}

$targetPath = Join-Path $triageDir "$Date.md"
if ((Test-Path $targetPath) -and (-not $Force)) {
  Write-Host "[skip] Triage file already exists: $targetPath" -ForegroundColor DarkYellow
  Write-Host "Use -Force to overwrite." -ForegroundColor DarkYellow
  exit 0
}

$content = Get-Content $templatePath -Raw
$content = $content -replace "- YYYY-MM-DD:", "- YYYY-MM-DD: $Date"
if ($Owner) {
  $content = $content -replace "- Owner:", "- Owner: $Owner"
}

$header = "# Daily Beta Triage - $Date`n`n"
if ($content -match "^#\s+Daily Beta Triage Template") {
  $content = $content -replace "^#\s+Daily Beta Triage Template", $header.TrimEnd()
}

$content | Out-File -FilePath $targetPath -Encoding utf8

Write-Host "[ok] Created triage file: $targetPath" -ForegroundColor Green
Write-Host "[next] Open and fill KPI + feedback sections after today's sessions." -ForegroundColor Cyan
