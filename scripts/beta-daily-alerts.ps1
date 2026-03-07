param(
  [string]$ReportPath = "",
  [string]$Date,
  [switch]$AppendToTriage,
  [switch]$Json,

  # --- configurable thresholds ---
  [int]$MinPrompts24h      = 1,
  [int]$MinDailyActive     = 1,
  [double]$MinApprovalRate = 0.5,
  [int]$MaxBlockedAttempts = 10,
  [double]$MinStreamRate   = 0.8,
  [int]$MinHistoryRows     = 1
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"

$effectiveReportPath = if ($ReportPath) { $ReportPath } else { "docs/reports/daily-status-$today.json" }

if (-not (Test-Path $effectiveReportPath)) {
  Write-Host ('  [error] Daily report not found: ' + $effectiveReportPath) -ForegroundColor Red
  Write-Host 'Run  npm run beta:daily:report  first.' -ForegroundColor DarkYellow
  exit 1
}

Write-Host "== tsifulator.ai daily alerts ==" -ForegroundColor Cyan
Write-Host "Date: $today"
Write-Host "Report: $effectiveReportPath"

$report = Get-Content -Path $effectiveReportPath -Raw | ConvertFrom-Json

$alerts = @()

# ---------- health check ----------
if ($report.health -ne "ok") {
  $alerts += [ordered]@{
    level   = "critical"
    check   = "api_health"
    message = "API health is '$($report.health)' (expected 'ok')"
    value   = $report.health
  }
}

# ---------- triage dedupe check ----------
if (-not $report.triage.dedupeHealthy) {
  $alerts += [ordered]@{
    level   = "warning"
    check   = "triage_dedupe"
    message = "Triage dedupe unhealthy: kpi=$($report.triage.kpiLineCount), comparison=$($report.triage.comparisonLineCount), trend=$($report.triage.trendLineCount)"
    value   = $false
  }
}

# ---------- triage existence ----------
if (-not $report.triage.exists) {
  $alerts += [ordered]@{
    level   = "warning"
    check   = "triage_missing"
    message = "Triage file does not exist: $($report.source.triagePath)"
    value   = $false
  }
}

# ---------- KPI checks (skip when KPI data unavailable) ----------
if ($null -ne $report.kpi) {

  # prompts sent in last 24h
  if ($report.kpi.promptsSent24h -lt $MinPrompts24h) {
    $alerts += [ordered]@{
      level   = "warning"
      check   = "low_prompts_24h"
      message = "promptsSent24h=$($report.kpi.promptsSent24h) < threshold $MinPrompts24h"
      value   = $report.kpi.promptsSent24h
      threshold = $MinPrompts24h
    }
  }

  # daily active users
  if ($report.kpi.dailyActiveUsers -lt $MinDailyActive) {
    $alerts += [ordered]@{
      level   = "warning"
      check   = "low_daily_active"
      message = "dailyActiveUsers=$($report.kpi.dailyActiveUsers) < threshold $MinDailyActive"
      value   = $report.kpi.dailyActiveUsers
      threshold = $MinDailyActive
    }
  }

  # approval rate
  $proposed = $report.kpi.applyActionsProposed
  $confirmed = $report.kpi.applyActionsConfirmed
  if ($proposed -gt 0) {
    $approvalRate = [math]::Round($confirmed / $proposed, 2)
    if ($approvalRate -lt $MinApprovalRate) {
      $alerts += [ordered]@{
        level     = "warning"
        check     = "low_approval_rate"
        message   = "approvalRate=$approvalRate ($confirmed/$proposed) < threshold $MinApprovalRate"
        value     = $approvalRate
        threshold = $MinApprovalRate
      }
    }
  }

  # blocked command attempts
  if ($report.kpi.blockedCommandAttempts -gt $MaxBlockedAttempts) {
    $alerts += [ordered]@{
      level     = "warning"
      check     = "high_blocked_attempts"
      message   = "blockedCommandAttempts=$($report.kpi.blockedCommandAttempts) > threshold $MaxBlockedAttempts"
      value     = $report.kpi.blockedCommandAttempts
      threshold = $MaxBlockedAttempts
    }
  }

  # stream success rate
  if ($report.kpi.streamSuccessRate -lt $MinStreamRate) {
    $alerts += [ordered]@{
      level     = "warning"
      check     = "low_stream_success"
      message   = "streamSuccessRate=$($report.kpi.streamSuccessRate) < threshold $MinStreamRate"
      value     = $report.kpi.streamSuccessRate
      threshold = $MinStreamRate
    }
  }
}
else {
  $alerts += [ordered]@{
    level   = "warning"
    check   = "kpi_unavailable"
    message = "KPI data is null in the daily report (API may have been down)"
    value   = $null
  }
}

# ---------- rolling history check ----------
if (-not $report.history.exists) {
  $alerts += [ordered]@{
    level   = "warning"
    check   = "history_missing"
    message = "Rolling history CSV does not exist: $($report.source.historyPath)"
    value   = $false
  }
}
elseif ($report.history.rowCount -lt $MinHistoryRows) {
  $alerts += [ordered]@{
    level   = "warning"
    check   = "low_history_rows"
    message = "historyRows=$($report.history.rowCount) < threshold $MinHistoryRows"
    value   = $report.history.rowCount
    threshold = $MinHistoryRows
  }
}

# ---------- output ----------
$alertCount = $alerts.Count
$criticalCount = ($alerts | Where-Object { $_.level -eq "critical" }).Count
$warningCount = ($alerts | Where-Object { $_.level -eq "warning" }).Count

$result = [ordered]@{
  date          = $today
  checkedAt     = (Get-Date).ToString("o")
  reportPath    = $effectiveReportPath
  alertCount    = $alertCount
  criticalCount = $criticalCount
  warningCount  = $warningCount
  status        = if ($criticalCount -gt 0) { "CRITICAL" } elseif ($warningCount -gt 0) { "WARN" } else { "OK" }
  alerts        = $alerts
  thresholds    = [ordered]@{
    minPrompts24h     = $MinPrompts24h
    minDailyActive    = $MinDailyActive
    minApprovalRate   = $MinApprovalRate
    maxBlockedAttempts = $MaxBlockedAttempts
    minStreamRate     = $MinStreamRate
    minHistoryRows    = $MinHistoryRows
  }
}

if ($Json) {
  $result | ConvertTo-Json -Depth 8
}
else {
  Write-Host ""
  if ($alertCount -eq 0) {
    Write-Host '[OK] No alerts - all thresholds passed.' -ForegroundColor Green
  }
  else {
    $statusColor = if ($criticalCount -gt 0) { "Red" } else { "Yellow" }
    Write-Host "$($result.status) | $alertCount alert(s): $criticalCount critical, $warningCount warning" -ForegroundColor $statusColor
    Write-Host ""
    foreach ($a in $alerts) {
      $icon = if ($a.level -eq 'critical') { '!!' } else { '! ' }
      $color = if ($a.level -eq 'critical') { 'Red' } else { 'DarkYellow' }
      Write-Host "  $icon $($a.check): $($a.message)" -ForegroundColor $color
    }
  }
  Write-Host ""
}

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host ('  [warn] Triage file not found: ' + $triagePath) -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $statusTag = $result.status
  $alertSummary = if ($alertCount -eq 0) { "all_clear" } else { ($alerts | ForEach-Object { $_.check }) -join "," }
  $line = "$stamp Daily alerts: status=$statusTag, alerts=$alertCount, critical=$criticalCount, warn=$warningCount, checks=$alertSummary"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch "Daily alerts:" }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted daily alerts in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] Daily alerts complete.' -ForegroundColor Green
