param(
  [string]$Email = "partner.a@company.com",
  [string]$PartnerEmailsCsv = "partner.a@company.com,partner.b@company.com",
  [string]$Owner = "",
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string]$Date
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }

$dailyOpsScript = Join-Path $PSScriptRoot "beta-daily-ops.ps1"
$dailyValidateScript = Join-Path $PSScriptRoot "beta-daily-validate.ps1"

foreach ($script in @($dailyOpsScript, $dailyValidateScript)) {
  if (-not (Test-Path $script)) {
    throw "Missing required script: $script"
  }
}

Write-Host "== tsifulator.ai daily full ==" -ForegroundColor Cyan
Write-Host "Date: $today"

Write-Host "[run] daily ops" -ForegroundColor Yellow
$opsArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $dailyOpsScript,
  "-Email", $Email,
  "-PartnerEmailsCsv", $PartnerEmailsCsv,
  "-BaseUrl", $BaseUrl,
  "-Date", $today
)
if ($Owner) {
  $opsArgs += @("-Owner", $Owner)
}
& powershell @opsArgs
if ($LASTEXITCODE -ne 0) {
  throw "Daily ops failed"
}

Write-Host "[run] strict validation" -ForegroundColor Yellow
$validateArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $dailyValidateScript,
  "-BaseUrl", $BaseUrl,
  "-Date", $today,
  "-Strict"
)
& powershell @validateArgs
if ($LASTEXITCODE -ne 0) {
  throw "Daily strict validation failed"
}

Write-Host ""
Write-Host "[done] Daily full workflow passed (ops + strict validation)." -ForegroundColor Green
