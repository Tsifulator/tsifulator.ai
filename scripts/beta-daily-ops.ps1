param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string]$Email = "partner.a@company.com",
  [string[]]$PartnerEmails = @(),
  [string]$PartnerEmailsCsv = "",
  [string]$Owner = "",
  [string]$Date
)

$ErrorActionPreference = "Stop"

if ($BaseUrl -notmatch '^https?://') {
  throw "Invalid -BaseUrl '$BaseUrl'. Use an http(s) URL. For multiple partners, prefer -PartnerEmailsCsv 'a@x.com,b@y.com'."
}

$effectivePartnerEmails = @()

if ($PartnerEmailsCsv) {
  $effectivePartnerEmails += ($PartnerEmailsCsv -split '[,;\s]+' | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
}

if ($PartnerEmails -and $PartnerEmails.Count -gt 0) {
  foreach ($entry in $PartnerEmails) {
    $effectivePartnerEmails += ($entry -split '[,;\s]+' | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
  }
}

$effectivePartnerEmails = $effectivePartnerEmails | Select-Object -Unique

if (-not $effectivePartnerEmails -or $effectivePartnerEmails.Count -eq 0) {
  $effectivePartnerEmails = @("partner.a@company.com", "partner.b@company.com")
}

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }

Write-Host "== tsifulator.ai daily ops ==" -ForegroundColor Cyan
Write-Host "Date: $today"

$checkpointScript = Join-Path $PSScriptRoot "beta-checkpoint-today.ps1"
$compareScript = Join-Path $PSScriptRoot "beta-compare-partners.ps1"
$trendScript = Join-Path $PSScriptRoot "beta-partner-trend.ps1"

foreach ($script in @($checkpointScript, $compareScript, $trendScript)) {
  if (-not (Test-Path $script)) {
    throw "Missing required script: $script"
  }
}

Write-Host "[run] checkpoint append" -ForegroundColor Yellow
$checkpointArgs = @("-ExecutionPolicy", "Bypass", "-File", $checkpointScript, "-Email", $Email, "-BaseUrl", $BaseUrl, "-Date", $today, "-AppendToTriage")
if ($Owner) {
  $checkpointArgs += @("-Owner", $Owner)
}
& powershell @checkpointArgs
if ($LASTEXITCODE -ne 0) {
  throw "Checkpoint append failed"
}

Write-Host "[run] partner comparison append + rolling CSV" -ForegroundColor Yellow
$compareArgs = @("-ExecutionPolicy", "Bypass", "-File", $compareScript, "-BaseUrl", $BaseUrl, "-Date", $today, "-AppendToTriage", "-AppendCsvDaily")
if ($effectivePartnerEmails -and $effectivePartnerEmails.Count -ge 2) {
  $compareArgs += "-Emails"
  $compareArgs += ($effectivePartnerEmails -join ",")
}
else {
  Write-Host "[warn] Less than two partner emails parsed; using compare script defaults." -ForegroundColor DarkYellow
}
& powershell @compareArgs
if ($LASTEXITCODE -ne 0) {
  throw "Partner comparison append failed"
}

Write-Host "[run] partner trend append" -ForegroundColor Yellow
$trendArgs = @("-ExecutionPolicy", "Bypass", "-File", $trendScript, "-Date", $today, "-AppendToTriage")
& powershell @trendArgs
if ($LASTEXITCODE -ne 0) {
  throw "Partner trend append failed"
}

Write-Host ""
Write-Host "[done] Daily ops complete (checkpoint + comparison + trend)." -ForegroundColor Green
