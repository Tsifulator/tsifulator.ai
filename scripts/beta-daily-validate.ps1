param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string]$Date,
  [switch]$Strict
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"
$historyPath = "docs/reports/partner-compare-history.csv"

$failed = $false

function Assert-Check {
  param(
    [string]$Name,
    [bool]$Condition,
    [string]$OnFail
  )

  if ($Condition) {
    Write-Host "[ok] $Name" -ForegroundColor Green
    return
  }

  Write-Host "[warn] $Name :: $OnFail" -ForegroundColor DarkYellow
  if ($Strict) {
    $script:failed = $true
  }
}

Write-Host "== tsifulator.ai daily validate ==" -ForegroundColor Cyan
Write-Host "Date: $today"

try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health" -TimeoutSec 5
  Assert-Check -Name "API health" -Condition ($health.status -eq "ok") -OnFail "unexpected /health status"
}
catch {
  Write-Host "[warn] API health :: API not reachable at $BaseUrl" -ForegroundColor DarkYellow
  if ($Strict) {
    $failed = $true
  }
}

Assert-Check -Name "Triage file exists" -Condition (Test-Path $triagePath) -OnFail "missing $triagePath"

if (Test-Path $triagePath) {
  $kpiCount = (Select-String -Path $triagePath -Pattern '^- \[[^\]]+\] KPI refresh:' -AllMatches).Matches.Count
  $comparisonCount = (Select-String -Path $triagePath -Pattern 'Partner comparison snapshot:' -AllMatches).Matches.Count
  $trendCount = (Select-String -Path $triagePath -Pattern 'Partner trend snapshot:' -AllMatches).Matches.Count
  $statusCount = (Select-String -Path $triagePath -Pattern 'Daily status report:' -AllMatches).Matches.Count
  $alertCount = (Select-String -Path $triagePath -Pattern 'Daily alerts:' -AllMatches).Matches.Count
  $feedbackCount = (Select-String -Path $triagePath -Pattern 'Feedback summary:' -AllMatches).Matches.Count
  $weeklyCount = (Select-String -Path $triagePath -Pattern 'Weekly digest' -AllMatches).Matches.Count
  $scorecardCount = (Select-String -Path $triagePath -Pattern 'Daily scorecard:' -AllMatches).Matches.Count
  $rotationCount = (Select-String -Path $triagePath -Pattern 'History rotation:' -AllMatches).Matches.Count
  $scorecardTrendCount = (Select-String -Path $triagePath -Pattern 'Scorecard trend:' -AllMatches).Matches.Count
  $sentimentCount = (Select-String -Path $triagePath -Pattern 'Feedback sentiment:' -AllMatches).Matches.Count
  $onboardingCount = (Select-String -Path $triagePath -Pattern 'Partner onboarding:' -AllMatches).Matches.Count
  $heatmapCount = (Select-String -Path $triagePath -Pattern 'Engagement heatmap:' -AllMatches).Matches.Count
  $churnCount = (Select-String -Path $triagePath -Pattern 'Churn risk:' -AllMatches).Matches.Count
  $sessionDepthCount = (Select-String -Path $triagePath -Pattern 'Session depth:' -AllMatches).Matches.Count
  $retentionCount = (Select-String -Path $triagePath -Pattern 'Retention curve:' -AllMatches).Matches.Count
  $slaCount = (Select-String -Path $triagePath -Pattern 'SLA monitor:' -AllMatches).Matches.Count
  $funnelCount = (Select-String -Path $triagePath -Pattern 'Adoption funnel:' -AllMatches).Matches.Count
  $anomalyCount = (Select-String -Path $triagePath -Pattern 'Anomaly detector:' -AllMatches).Matches.Count
  $qualityCount = (Select-String -Path $triagePath -Pattern 'Data quality audit:' -AllMatches).Matches.Count
  $engagementCount = (Select-String -Path $triagePath -Pattern 'Engagement score:' -AllMatches).Matches.Count
  $riskAuditCount = (Select-String -Path $triagePath -Pattern 'Command risk audit:' -AllMatches).Matches.Count
  $driftCount = (Select-String -Path $triagePath -Pattern 'Daily drift detector:' -AllMatches).Matches.Count
  $timelineCount = (Select-String -Path $triagePath -Pattern 'Partner activity timeline:' -AllMatches).Matches.Count

  $calendarCount = (Select-String -Path $triagePath -Pattern 'Partner session calendar:' -AllMatches).Matches.Count
  Write-Host "[info] CALENDAR_LINE_COUNT=$calendarCount"

  Write-Host "[info] KPI_LINE_COUNT=$kpiCount"
  Write-Host "[info] COMPARISON_LINE_COUNT=$comparisonCount"
  Write-Host "[info] TREND_LINE_COUNT=$trendCount"
  Write-Host "[info] STATUS_LINE_COUNT=$statusCount"
  Write-Host "[info] ALERTS_LINE_COUNT=$alertCount"
  Write-Host "[info] FEEDBACK_LINE_COUNT=$feedbackCount"
  Write-Host "[info] WEEKLY_LINE_COUNT=$weeklyCount"
  Write-Host "[info] SCORECARD_LINE_COUNT=$scorecardCount"
  Write-Host "[info] ROTATION_LINE_COUNT=$rotationCount"
  Write-Host "[info] SCORECARD_TREND_LINE_COUNT=$scorecardTrendCount"
  Write-Host "[info] SENTIMENT_LINE_COUNT=$sentimentCount"
  Write-Host "[info] ONBOARDING_LINE_COUNT=$onboardingCount"
  Write-Host "[info] HEATMAP_LINE_COUNT=$heatmapCount"
  Write-Host "[info] CHURN_LINE_COUNT=$churnCount"
  Write-Host "[info] SESSION_DEPTH_LINE_COUNT=$sessionDepthCount"
  Write-Host "[info] RETENTION_LINE_COUNT=$retentionCount"
  Write-Host "[info] SLA_LINE_COUNT=$slaCount"
  Write-Host "[info] FUNNEL_LINE_COUNT=$funnelCount"
  Write-Host "[info] ANOMALY_LINE_COUNT=$anomalyCount"
  Write-Host "[info] QUALITY_LINE_COUNT=$qualityCount"
  Write-Host "[info] ENGAGEMENT_LINE_COUNT=$engagementCount"
  Write-Host "[info] RISK_AUDIT_LINE_COUNT=$riskAuditCount"
  Write-Host "[info] DRIFT_LINE_COUNT=$driftCount"
  Write-Host "[info] TIMELINE_LINE_COUNT=$timelineCount"

  Assert-Check -Name "Partner session calendar line dedupe" -Condition ($calendarCount -le 1) -OnFail "expected <= 1, got $calendarCount"

  Assert-Check -Name "KPI line dedupe" -Condition ($kpiCount -le 1) -OnFail "expected <= 1, got $kpiCount"
  Assert-Check -Name "Comparison line dedupe" -Condition ($comparisonCount -le 1) -OnFail "expected <= 1, got $comparisonCount"
  Assert-Check -Name "Trend line dedupe" -Condition ($trendCount -le 1) -OnFail "expected <= 1, got $trendCount"
  Assert-Check -Name "Status line dedupe" -Condition ($statusCount -le 1) -OnFail "expected <= 1, got $statusCount"
  Assert-Check -Name "Alerts line dedupe" -Condition ($alertCount -le 1) -OnFail "expected <= 1, got $alertCount"
  Assert-Check -Name "Feedback line dedupe" -Condition ($feedbackCount -le 1) -OnFail "expected <= 1, got $feedbackCount"
  Assert-Check -Name "Weekly line dedupe" -Condition ($weeklyCount -le 1) -OnFail "expected <= 1, got $weeklyCount"
  Assert-Check -Name "Scorecard line dedupe" -Condition ($scorecardCount -le 1) -OnFail "expected <= 1, got $scorecardCount"
  Assert-Check -Name "Rotation line dedupe" -Condition ($rotationCount -le 1) -OnFail "expected <= 1, got $rotationCount"
  Assert-Check -Name "Scorecard trend line dedupe" -Condition ($scorecardTrendCount -le 1) -OnFail "expected <= 1, got $scorecardTrendCount"
  Assert-Check -Name "Sentiment line dedupe" -Condition ($sentimentCount -le 1) -OnFail "expected <= 1, got $sentimentCount"
  Assert-Check -Name "Onboarding line dedupe" -Condition ($onboardingCount -le 1) -OnFail "expected <= 1, got $onboardingCount"
  Assert-Check -Name "Heatmap line dedupe" -Condition ($heatmapCount -le 1) -OnFail "expected <= 1, got $heatmapCount"
  Assert-Check -Name "Churn line dedupe" -Condition ($churnCount -le 1) -OnFail "expected <= 1, got $churnCount"
  Assert-Check -Name "Session depth line dedupe" -Condition ($sessionDepthCount -le 1) -OnFail "expected <= 1, got $sessionDepthCount"
  Assert-Check -Name "Retention curve line dedupe" -Condition ($retentionCount -le 1) -OnFail "expected <= 1, got $retentionCount"
  Assert-Check -Name "SLA monitor line dedupe" -Condition ($slaCount -le 1) -OnFail "expected <= 1, got $slaCount"
  Assert-Check -Name "Adoption funnel line dedupe" -Condition ($funnelCount -le 1) -OnFail "expected <= 1, got $funnelCount"
  Assert-Check -Name "Anomaly detector line dedupe" -Condition ($anomalyCount -le 1) -OnFail "expected <= 1, got $anomalyCount"
  Assert-Check -Name "Data quality line dedupe" -Condition ($qualityCount -le 1) -OnFail "expected <= 1, got $qualityCount"
  Assert-Check -Name "Engagement score line dedupe" -Condition ($engagementCount -le 1) -OnFail "expected <= 1, got $engagementCount"
  Assert-Check -Name "Command risk audit line dedupe" -Condition ($riskAuditCount -le 1) -OnFail "expected <= 1, got $riskAuditCount"
  Assert-Check -Name "Daily drift line dedupe" -Condition ($driftCount -le 1) -OnFail "expected <= 1, got $driftCount"
  Assert-Check -Name "Partner activity timeline line dedupe" -Condition ($timelineCount -le 1) -OnFail "expected <= 1, got $timelineCount"
}

Assert-Check -Name "Rolling history exists" -Condition (Test-Path $historyPath) -OnFail "missing $historyPath"

if (Test-Path $historyPath) {
  $rows = Import-Csv -Path $historyPath
  Assert-Check -Name "Rolling history has rows" -Condition ($rows.Count -gt 0) -OnFail "no rows in $historyPath"

  $requiredCols = @(
    "snapshotDate", "snapshotAt", "email", "prompts24h", "proposed", "confirmed", "blocked",
    "streamRequests", "streamCompletions", "streamRatio", "medianChatLatencyMs", "medianStreamFirstTokenLatencyMs"
  )

  $headers = if ($rows.Count -gt 0) { $rows[0].PSObject.Properties.Name } else { @() }
  $missing = @($requiredCols | Where-Object { $headers -notcontains $_ })
  Assert-Check -Name "Rolling history schema" -Condition ($missing.Count -eq 0) -OnFail ("missing columns: " + ($missing -join ", "))
}

Write-Host ""
if ($failed) {
  Write-Host "[done] Daily validation failed in strict mode." -ForegroundColor Red
  exit 1
}

Write-Host "[done] Daily validation complete." -ForegroundColor Green
