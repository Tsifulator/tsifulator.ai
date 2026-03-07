$calendarScript = Join-Path $PSScriptRoot "beta-session-calendar.ps1"

foreach ($script in @($dailyFullScript, $dailyReportScript, $dailyAlertsScript, $feedbackScript, $sentimentScript, $weeklyDigestScript, $scorecardScript, $scorecardTrendScript, $rotateScript, $onboardingScript, $heatmapScript, $churnScript, $sessionDepthScript, $retentionScript, $slaScript, $funnelScript, $anomalyScript, $qualityScript, $engagementScript, $riskScript, $driftScript, $timelineScript, $calendarScript)) {
  if (-not (Test-Path $script)) {
    throw "Missing required script: $script"
  }
}
Write-Host '[run] partner session calendar' -ForegroundColor Yellow
$calendarArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $calendarScript,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $calList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($calList.Count -gt 0) {
    $calendarArgs += "-Emails"
    $calendarArgs += ($calList -join ",")
  }
}
& powershell @calendarArgs
if ($LASTEXITCODE -ne 0) {
  throw "Partner session calendar failed"
}
param(
  [string]$Email = "partner.a@company.com",
  [string]$PartnerEmailsCsv = "partner.a@company.com,partner.b@company.com",
  [string]$Owner = "",
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string]$Date,
  [string]$ReportOutputPath = ""
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }

$dailyFullScript = Join-Path $PSScriptRoot "beta-daily-full.ps1"
$dailyReportScript = Join-Path $PSScriptRoot "beta-daily-report.ps1"
$dailyAlertsScript = Join-Path $PSScriptRoot "beta-daily-alerts.ps1"
$feedbackScript = Join-Path $PSScriptRoot "beta-feedback-summary.ps1"
$sentimentScript = Join-Path $PSScriptRoot "beta-feedback-sentiment.ps1"
$weeklyDigestScript = Join-Path $PSScriptRoot "beta-weekly-digest.ps1"
$scorecardScript = Join-Path $PSScriptRoot "beta-daily-scorecard.ps1"
$scorecardTrendScript = Join-Path $PSScriptRoot "beta-scorecard-trend.ps1"
$rotateScript = Join-Path $PSScriptRoot "beta-history-rotate.ps1"
$onboardingScript = Join-Path $PSScriptRoot "beta-partner-onboarding.ps1"
$heatmapScript = Join-Path $PSScriptRoot "beta-engagement-heatmap.ps1"
$churnScript = Join-Path $PSScriptRoot "beta-churn-risk.ps1"
$sessionDepthScript = Join-Path $PSScriptRoot "beta-session-depth.ps1"
$retentionScript = Join-Path $PSScriptRoot "beta-retention-curve.ps1"
$slaScript = Join-Path $PSScriptRoot "beta-sla-monitor.ps1"
$funnelScript = Join-Path $PSScriptRoot "beta-adoption-funnel.ps1"
$anomalyScript = Join-Path $PSScriptRoot "beta-anomaly-detector.ps1"
$qualityScript = Join-Path $PSScriptRoot "beta-data-quality.ps1"
$engagementScript = Join-Path $PSScriptRoot "beta-engagement-score.ps1"
$riskScript = Join-Path $PSScriptRoot "beta-command-risk-audit.ps1"
 $driftScript = Join-Path $PSScriptRoot "beta-drift-detector.ps1"
 $timelineScript = Join-Path $PSScriptRoot "beta-activity-timeline.ps1"

foreach ($script in @($dailyFullScript, $dailyReportScript, $dailyAlertsScript, $feedbackScript, $sentimentScript, $weeklyDigestScript, $scorecardScript, $scorecardTrendScript, $rotateScript, $onboardingScript, $heatmapScript, $churnScript, $sessionDepthScript, $retentionScript, $slaScript, $funnelScript, $anomalyScript, $qualityScript, $engagementScript, $riskScript, $driftScript, $timelineScript)) {
  if (-not (Test-Path $script)) {
    throw "Missing required script: $script"
  }
}

Write-Host "== tsifulator.ai daily bundle ==" -ForegroundColor Cyan
Write-Host "Date: $today"

Write-Host "[run] daily full" -ForegroundColor Yellow
$fullArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $dailyFullScript,
  "-Email", $Email,
  "-PartnerEmailsCsv", $PartnerEmailsCsv,
  "-BaseUrl", $BaseUrl,
  "-Date", $today
)
if ($Owner) {
  $fullArgs += @("-Owner", $Owner)
}
& powershell @fullArgs
if ($LASTEXITCODE -ne 0) {
  throw "Daily full failed"
}

Write-Host "[run] daily report" -ForegroundColor Yellow
$reportArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $dailyReportScript,
  "-Email", $Email,
  "-BaseUrl", $BaseUrl,
  "-Date", $today,
  "-AppendToTriage"
)
if ($ReportOutputPath) {
  $reportArgs += @("-OutputPath", $ReportOutputPath)
}
& powershell @reportArgs
if ($LASTEXITCODE -ne 0) {
  throw "Daily report failed"
}

# --- resolve the report path for alerts ---
$effectiveReportPath = if ($ReportOutputPath) { $ReportOutputPath } else { "docs/reports/daily-status-$today.json" }

Write-Host "[run] daily alerts" -ForegroundColor Yellow
$alertArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $dailyAlertsScript,
  "-ReportPath", $effectiveReportPath,
  "-Date", $today,
  "-AppendToTriage"
)
& powershell @alertArgs
if ($LASTEXITCODE -ne 0) {
  throw "Daily alerts failed"
}

Write-Host '[run] feedback summary' -ForegroundColor Yellow
$feedbackArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $feedbackScript,
  "-BaseUrl", $BaseUrl,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $emailList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($emailList.Count -gt 0) {
    $feedbackArgs += "-Emails"
    $feedbackArgs += ($emailList -join ",")
  }
}
& powershell @feedbackArgs
if ($LASTEXITCODE -ne 0) {
  throw "Feedback summary failed"
}

Write-Host '[run] feedback sentiment' -ForegroundColor Yellow
$sentimentArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $sentimentScript,
  "-BaseUrl", $BaseUrl,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $sentList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($sentList.Count -gt 0) {
    $sentimentArgs += "-Emails"
    $sentimentArgs += ($sentList -join ",")
  }
}
& powershell @sentimentArgs
if ($LASTEXITCODE -ne 0) {
  throw "Feedback sentiment failed"
}

Write-Host '[run] weekly digest' -ForegroundColor Yellow
$digestArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $weeklyDigestScript,
  "-Date", $today,
  "-AppendToTriage"
)
& powershell @digestArgs
if ($LASTEXITCODE -ne 0) {
  throw "Weekly digest failed"
}

Write-Host '[run] daily scorecard' -ForegroundColor Yellow
$scorecardArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $scorecardScript,
  "-BaseUrl", $BaseUrl,
  "-Email", $Email,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $scorecardArgs += "-Emails"
  $scorecardArgs += $PartnerEmailsCsv
}
& powershell @scorecardArgs
if ($LASTEXITCODE -ne 0) {
  throw "Daily scorecard failed"
}

Write-Host '[run] scorecard trend' -ForegroundColor Yellow
$trendArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $scorecardTrendScript,
  "-BaseUrl", $BaseUrl,
  "-Email", $Email,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $trendArgs += "-Emails"
  $trendArgs += $PartnerEmailsCsv
}
& powershell @trendArgs
if ($LASTEXITCODE -ne 0) {
  throw "Scorecard trend failed"
}

Write-Host '[run] partner onboarding' -ForegroundColor Yellow
$onboardingArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $onboardingScript,
  "-BaseUrl", $BaseUrl,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $obList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($obList.Count -gt 0) {
    $onboardingArgs += "-Emails"
    $onboardingArgs += ($obList -join ",")
  }
}
& powershell @onboardingArgs
if ($LASTEXITCODE -ne 0) {
  throw "Partner onboarding failed"
}

Write-Host '[run] engagement heatmap' -ForegroundColor Yellow
$heatmapArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $heatmapScript,
  "-BaseUrl", $BaseUrl,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $hmList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($hmList.Count -gt 0) {
    $heatmapArgs += "-Emails"
    $heatmapArgs += ($hmList -join ",")
  }
}
& powershell @heatmapArgs
if ($LASTEXITCODE -ne 0) {
  throw "Engagement heatmap failed"
}

Write-Host '[run] churn risk' -ForegroundColor Yellow
$churnArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $churnScript,
  "-BaseUrl", $BaseUrl,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $crList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($crList.Count -gt 0) {
    $churnArgs += "-Emails"
    $churnArgs += ($crList -join ",")
  }
}
& powershell @churnArgs
if ($LASTEXITCODE -ne 0) {
  throw "Churn risk failed"
}

Write-Host '[run] session depth' -ForegroundColor Yellow
$depthArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $sessionDepthScript,
  "-BaseUrl", $BaseUrl,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $sdList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($sdList.Count -gt 0) {
    $depthArgs += "-Emails"
    $depthArgs += ($sdList -join ",")
  }
}
& powershell @depthArgs
if ($LASTEXITCODE -ne 0) {
  throw "Session depth failed"
}

Write-Host '[run] retention curve' -ForegroundColor Yellow
$retArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $retentionScript,
  "-BaseUrl", $BaseUrl,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $retList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($retList.Count -gt 0) {
    $retArgs += "-Emails"
    $retArgs += ($retList -join ",")
  }
}
& powershell @retArgs
if ($LASTEXITCODE -ne 0) {
  throw "Retention curve failed"
}

Write-Host '[run] SLA monitor' -ForegroundColor Yellow
$slaArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $slaScript,
  "-BaseUrl", $BaseUrl,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $slaList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($slaList.Count -gt 0) {
    $slaArgs += "-Emails"
    $slaArgs += ($slaList -join ",")
  }
}
& powershell @slaArgs
if ($LASTEXITCODE -ne 0) {
  throw "SLA monitor failed"
}

Write-Host '[run] adoption funnel' -ForegroundColor Yellow
$funnelArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $funnelScript,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $funnelList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($funnelList.Count -gt 0) {
    $funnelArgs += "-Emails"
    $funnelArgs += ($funnelList -join ",")
  }
}
& powershell @funnelArgs
if ($LASTEXITCODE -ne 0) {
  throw "Adoption funnel failed"
}

Write-Host '[run] anomaly detector' -ForegroundColor Yellow
$anomalyArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $anomalyScript,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $anomalyList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($anomalyList.Count -gt 0) {
    $anomalyArgs += "-Emails"
    $anomalyArgs += ($anomalyList -join ",")
  }
}
& powershell @anomalyArgs
if ($LASTEXITCODE -ne 0) {
  throw "Anomaly detector failed"
}

Write-Host '[run] data quality audit' -ForegroundColor Yellow
$qualityArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $qualityScript,
  "-Date", $today,
  "-AppendToTriage"
)
& powershell @qualityArgs
if ($LASTEXITCODE -ne 0) {
  throw "Data quality audit failed"
}

Write-Host '[run] engagement score' -ForegroundColor Yellow
$engagementArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $engagementScript,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $engList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($engList.Count -gt 0) {
    $engagementArgs += "-Emails"
    $engagementArgs += ($engList -join ",")
  }
}
& powershell @engagementArgs
if ($LASTEXITCODE -ne 0) {
  throw "Engagement score failed"
}

Write-Host '[run] command risk audit' -ForegroundColor Yellow
$riskArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $riskScript,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $rskList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($rskList.Count -gt 0) {
    $riskArgs += "-Emails"
    $riskArgs += ($rskList -join ",")
  }
}
& powershell @riskArgs
if ($LASTEXITCODE -ne 0) {
  throw "Command risk audit failed"
}


Write-Host '[run] daily drift detector' -ForegroundColor Yellow
$driftArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $driftScript,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $dftList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($dftList.Count -gt 0) {
    $driftArgs += "-Emails"
    $driftArgs += ($dftList -join ",")
  }
}
& powershell @driftArgs
if ($LASTEXITCODE -ne 0) {
  throw "Daily drift detector failed"
}

Write-Host '[run] partner activity timeline' -ForegroundColor Yellow
$timelineArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $timelineScript,
  "-Date", $today,
  "-AppendToTriage"
)
if ($PartnerEmailsCsv) {
  $tlList = $PartnerEmailsCsv -split '[,;]+' | Where-Object { $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() }
  if ($tlList.Count -gt 0) {
    $timelineArgs += "-Emails"
    $timelineArgs += ($tlList -join ",")
  }
}
& powershell @timelineArgs
if ($LASTEXITCODE -ne 0) {
  throw "Partner activity timeline failed"
}

Write-Host '[run] history rotation' -ForegroundColor Yellow
$rotateArgs = @(
  "-ExecutionPolicy", "Bypass",
  "-File", $rotateScript,
  "-Date", $today,
  "-Archive",
  "-AppendToTriage"
)
& powershell @rotateArgs
if ($LASTEXITCODE -ne 0) {
  throw "History rotation failed"
}

Write-Host ''
Write-Host '[done] Daily bundle complete (full + report + alerts + feedback + sentiment + digest + scorecard + trend + onboarding + heatmap + churn + depth + retention + sla + funnel + anomaly + quality + engagement + risk + drift + timeline + calendar + rotation).' -ForegroundColor Green
