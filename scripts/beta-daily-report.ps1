param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string]$Email = "partner.a@company.com",
  [string]$Date,
  [string]$OutputPath = "",
  [switch]$AppendToTriage
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"
$historyPath = "docs/reports/partner-compare-history.csv"
$targetPath = if ($OutputPath) { $OutputPath } else { "docs/reports/daily-status-$today.json" }

$targetDir = Split-Path -Path $targetPath -Parent
if ($targetDir -and -not (Test-Path $targetDir)) {
  New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

Write-Host "== tsifulator.ai daily report ==" -ForegroundColor Cyan
Write-Host "Date: $today"

$healthStatus = "unreachable"
$kpi = $null

try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health" -TimeoutSec 5
  $healthStatus = if ($health.status) { [string]$health.status } else { "ok" }

  $loginBody = @{ email = $Email } | ConvertTo-Json
  $login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType "application/json" -Body $loginBody
  $headers = @{ authorization = "Bearer $($login.token)" }
  $kpiResponse = Invoke-RestMethod -Method Get -Uri "$BaseUrl/telemetry/counters" -Headers $headers
  $kpi = $kpiResponse.counters
}
catch {
  Write-Host "[warn] API not reachable for KPI snapshot at $BaseUrl" -ForegroundColor DarkYellow
}

$kpiCount = 0
$comparisonCount = 0
$trendCount = 0

if (Test-Path $triagePath) {
  $kpiCount = (Select-String -Path $triagePath -Pattern '^- \[[^\]]+\] KPI refresh:' -AllMatches).Matches.Count
  $comparisonCount = (Select-String -Path $triagePath -Pattern 'Partner comparison snapshot:' -AllMatches).Matches.Count
  $trendCount = (Select-String -Path $triagePath -Pattern 'Partner trend snapshot:' -AllMatches).Matches.Count
}

$historyRows = 0
if (Test-Path $historyPath) {
  $rows = Import-Csv -Path $historyPath
  $historyRows = if ($rows) { $rows.Count } else { 0 }
}

$report = [ordered]@{
  generatedAt = (Get-Date).ToString("o")
  date = $today
  health = $healthStatus
  source = [ordered]@{
    triagePath = $triagePath
    historyPath = $historyPath
    apiBaseUrl = $BaseUrl
    email = $Email
  }
  triage = [ordered]@{
    exists = (Test-Path $triagePath)
    kpiLineCount = $kpiCount
    comparisonLineCount = $comparisonCount
    trendLineCount = $trendCount
    dedupeHealthy = ($kpiCount -le 1 -and $comparisonCount -le 1 -and $trendCount -le 1)
  }
  history = [ordered]@{
    exists = (Test-Path $historyPath)
    rowCount = $historyRows
  }
  kpi = $kpi
}

$report | ConvertTo-Json -Depth 8 | Set-Content -Path $targetPath -Encoding UTF8
Write-Host "[ok] Daily report written: $targetPath" -ForegroundColor Green

if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host "[warn] Triage file not found: $triagePath" -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = $report.generatedAt
  $dedupeHealthyText = if ($report.triage.dedupeHealthy) { "true" } else { "false" }
  $historyRowsText = $report.history.rowCount
  $kpiPromptText = if ($null -ne $report.kpi) { [string]$report.kpi.promptsSent24h } else { "n/a" }
  $line = "[$stamp] Daily status report: health=$($report.health), dedupeHealthy=$dedupeHealthyText, historyRows=$historyRowsText, prompts24h=$kpiPromptText, report=$targetPath"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch "Daily status report:" }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host "[ok] Upserted daily status report in triage: $triagePath" -ForegroundColor Green
}

Write-Host "[done] Daily report complete." -ForegroundColor Green
