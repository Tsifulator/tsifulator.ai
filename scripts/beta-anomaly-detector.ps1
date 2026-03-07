param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [double]$Threshold = 2.0,
  [switch]$Json,
  [switch]$AppendToTriage,
  [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"
$historyPath = "docs/reports/partner-compare-history.csv"

# Normalize emails
$normalizedEmails = @()
foreach ($entry in $Emails) {
  $normalizedEmails += ($entry -split '[,;\s]+' | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
}
$Emails = $normalizedEmails | Select-Object -Unique

Write-Host '== tsifulator.ai anomaly detector ==' -ForegroundColor Cyan
Write-Host "Date: $today"
Write-Host "Threshold: $($Threshold) sigma"

# ---------- load rolling history ----------
if (-not (Test-Path $historyPath)) {
  Write-Host '  [warn] No rolling history found — cannot detect anomalies' -ForegroundColor DarkYellow
  exit 0
}

$allRows = Import-Csv -Path $historyPath
if (-not $allRows -or $allRows.Count -eq 0) {
  Write-Host '  [warn] Rolling history is empty' -ForegroundColor DarkYellow
  exit 0
}

# ---------- collect today's live metrics from API ----------
$liveMetrics = @{}
foreach ($em in $Emails) {
  try {
    $lb = @{ email = $em } | ConvertTo-Json
    $lg = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $lb
    $hd = @{ authorization = "Bearer $($lg.token)" }
    $kpiResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/telemetry/counters" -Headers $hd
    $kpi = $kpiResp.counters

    $sess = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=100" -Headers $hd
    $sessionCount = if ($sess.sessions) { $sess.sessions.Count } else { 0 }
    $eventCount = 0
    $feedbackCount = 0
    foreach ($s in $sess.sessions) {
      try {
        $ev = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($s.id)/events" -Headers $hd
        $eventCount += $ev.events.Count
        $feedbackCount += @($ev.events | Where-Object { $_.type -eq 'user_feedback' }).Count
      } catch {}
    }

    $liveMetrics[$em] = [ordered]@{
      prompts24h   = [int]$kpi.promptsSent24h
      proposed     = [int]$kpi.applyActionsProposed
      confirmed    = [int]$kpi.applyActionsConfirmed
      blocked      = [int]$kpi.blockedCommandAttempts
      sessions     = $sessionCount
      events       = $eventCount
      feedback     = $feedbackCount
    }
  }
  catch {
    Write-Host "  [warn] Could not fetch live metrics for $em" -ForegroundColor DarkYellow
  }
}

# ---------- compute historical baselines per partner ----------
# Metrics to analyze from rolling CSV
$csvMetrics = @('prompts24h', 'proposed', 'confirmed', 'blocked')

$anomalies = @()
$partnerResults = @()

foreach ($em in $Emails) {
  $partnerRows = @($allRows | Where-Object { $_.email -eq $em })
  $partnerAnomalies = @()

  if ($partnerRows.Count -lt 2) {
    $partnerResults += [ordered]@{
      email     = $em
      dataPoints = $partnerRows.Count
      status    = 'insufficient_data'
      anomalies = @()
    }
    continue
  }

  # Analyze CSV-based metrics
  foreach ($metric in $csvMetrics) {
    $values = @($partnerRows | ForEach-Object { [double]$_.$metric })
    if ($values.Count -lt 2) { continue }

    $avg = ($values | Measure-Object -Average).Average
    $sum = 0
    foreach ($v in $values) { $sum += ($v - $avg) * ($v - $avg) }
    $stddev = [math]::Sqrt($sum / $values.Count)

    # Get latest value
    $latest = $values[-1]

    # Also check live value if available
    $liveVal = $null
    if ($liveMetrics.Contains($em) -and $liveMetrics[$em].Contains($metric)) {
      $liveVal = [double]$liveMetrics[$em][$metric]
    }
    $checkVal = if ($null -ne $liveVal) { $liveVal } else { $latest }

    # Compute z-score
    $zscore = if ($stddev -gt 0) { [math]::Round(($checkVal - $avg) / $stddev, 2) } else { 0 }
    $isAnomaly = [math]::Abs($zscore) -ge $Threshold

    if ($isAnomaly) {
      $direction = if ($zscore -gt 0) { 'spike' } else { 'drop' }
      $anom = [ordered]@{
        metric    = $metric
        value     = $checkVal
        mean      = [math]::Round($avg, 2)
        stddev    = [math]::Round($stddev, 2)
        zscore    = $zscore
        direction = $direction
        severity  = if ([math]::Abs($zscore) -ge 3) { 'critical' } else { 'warning' }
      }
      $partnerAnomalies += $anom
      $anomalies += [ordered]@{ email = $em; anomaly = $anom }
    }
  }

  # Analyze approval rate (derived metric)
  $approvalRates = @()
  foreach ($row in $partnerRows) {
    $prop = [double]$row.proposed
    $conf = [double]$row.confirmed
    $rate = if ($prop -gt 0) { [math]::Round($conf / $prop, 4) } else { 0 }
    $approvalRates += $rate
  }
  if ($approvalRates.Count -ge 2) {
    $avg = ($approvalRates | Measure-Object -Average).Average
    $sum = 0
    foreach ($v in $approvalRates) { $sum += ($v - $avg) * ($v - $avg) }
    $stddev = [math]::Sqrt($sum / $approvalRates.Count)
    $latest = $approvalRates[-1]
    $zscore = if ($stddev -gt 0) { [math]::Round(($latest - $avg) / $stddev, 2) } else { 0 }
    if ([math]::Abs($zscore) -ge $Threshold) {
      $direction = if ($zscore -gt 0) { 'spike' } else { 'drop' }
      $anom = [ordered]@{
        metric    = 'approvalRate'
        value     = [math]::Round($latest, 4)
        mean      = [math]::Round($avg, 4)
        stddev    = [math]::Round($stddev, 4)
        zscore    = $zscore
        direction = $direction
        severity  = if ([math]::Abs($zscore) -ge 3) { 'critical' } else { 'warning' }
      }
      $partnerAnomalies += $anom
      $anomalies += [ordered]@{ email = $em; anomaly = $anom }
    }
  }

  $status = if ($partnerAnomalies.Count -eq 0) { 'normal' }
            elseif ($partnerAnomalies | Where-Object { $_.severity -eq 'critical' }) { 'critical' }
            else { 'warning' }

  $partnerResults += [ordered]@{
    email      = $em
    dataPoints = $partnerRows.Count
    status     = $status
    anomalies  = $partnerAnomalies
  }
}

# ---------- overall health ----------
$criticalCount = @($anomalies | Where-Object { $_.anomaly.severity -eq 'critical' }).Count
$warningCount = @($anomalies | Where-Object { $_.anomaly.severity -eq 'warning' }).Count
$totalAnomalies = $anomalies.Count
$overallStatus = if ($criticalCount -gt 0) { 'critical' }
                 elseif ($warningCount -gt 0) { 'warning' }
                 else { 'normal' }

$result = [ordered]@{
  generatedAt     = (Get-Date).ToString("o")
  date            = $today
  threshold       = $Threshold
  overallStatus   = $overallStatus
  totalAnomalies  = $totalAnomalies
  criticalCount   = $criticalCount
  warningCount    = $warningCount
  historyRows     = $allRows.Count
  partners        = $partnerResults
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $result | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host "  [ok] Anomaly report written: $OutputPath" -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $result | ConvertTo-Json -Depth 8
}
else {
  $statusColor = switch ($overallStatus) {
    'normal'   { 'Green' }
    'warning'  { 'Yellow' }
    'critical' { 'Red' }
    default    { 'White' }
  }
  Write-Host "  Status: $overallStatus" -ForegroundColor $statusColor
  Write-Host "  Anomalies: $totalAnomalies (critical=$criticalCount, warning=$warningCount)"
  Write-Host "  History rows analyzed: $($allRows.Count)"
  Write-Host ''

  foreach ($pr in $partnerResults) {
    $prColor = switch ($pr.status) {
      'normal'            { 'Green' }
      'warning'           { 'Yellow' }
      'critical'          { 'Red' }
      'insufficient_data' { 'DarkGray' }
      default             { 'White' }
    }
    Write-Host "  $($pr.email):" -ForegroundColor $prColor
    Write-Host "    status=$($pr.status)  dataPoints=$($pr.dataPoints)  anomalies=$($pr.anomalies.Count)"

    if ($pr.anomalies.Count -gt 0) {
      foreach ($a in $pr.anomalies) {
        $icon = if ($a.severity -eq 'critical') { '!!' } else { '!' }
        $aColor = if ($a.severity -eq 'critical') { 'Red' } else { 'Yellow' }
        $line = "      $icon $($a.metric): value=$($a.value) mean=$($a.mean) z=$($a.zscore) ($($a.direction))"
        Write-Host $line -ForegroundColor $aColor
      }
    }
    else {
      Write-Host '      All metrics within normal range' -ForegroundColor DarkGray
    }
  }
  Write-Host ''
}

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host "  [warn] Triage file not found: $triagePath" -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $anomStr = if ($totalAnomalies -gt 0) {
    $details = ($anomalies | ForEach-Object {
      "$($_.email):$($_.anomaly.metric)(z=$($_.anomaly.zscore),$($_.anomaly.direction))"
    }) -join '; '
    $details
  } else { 'none' }
  $line = "$stamp Anomaly detector: status=$overallStatus, anomalies=$totalAnomalies (critical=$criticalCount, warning=$warningCount), details=[$anomStr]"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Anomaly detector:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host "  [ok] Upserted anomaly detector in triage: $triagePath" -ForegroundColor Green
}

Write-Host '[done] Anomaly detector complete.' -ForegroundColor Green
