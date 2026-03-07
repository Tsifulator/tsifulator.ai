param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [string]$HistoryPath = "docs/reports/partner-compare-history.csv",
  [int]$SessionLimit = 100,
  [int]$InactiveDaysThreshold = 3,
  [double]$DeclineThreshold = 0.5,
  [switch]$Json,
  [switch]$AppendToTriage,
  [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"

# Normalize emails
$normalizedEmails = @()
foreach ($entry in $Emails) {
  $normalizedEmails += ($entry -split '[,;\s]+' | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
}
$Emails = $normalizedEmails | Select-Object -Unique

Write-Host '== tsifulator.ai churn risk detector ==' -ForegroundColor Cyan
Write-Host "Date:     $today"
Write-Host "Partners: $($Emails -join ', ')"

# ---------- health check ----------
try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health" -TimeoutSec 5
  Write-Host ('[ok] API healthy: ' + $health.status) -ForegroundColor Green
}
catch {
  Write-Host '[error] API unreachable' -ForegroundColor Red
  exit 1
}

# ---------- load rolling history ----------
$historyExists = Test-Path $HistoryPath
$historyRows = @()
if ($historyExists) {
  $historyRows = @(Import-Csv -Path $HistoryPath)
  Write-Host "  History: $($historyRows.Count) rows from $HistoryPath" -ForegroundColor Gray
}
else {
  Write-Host '  [warn] No rolling history found, using live API data only' -ForegroundColor DarkYellow
}

# ---------- risk signals per partner ----------
$partnerResults = @()
$riskCounts = @{ low = 0; medium = 0; high = 0; critical = 0 }

foreach ($email in $Emails) {
  $emailClean = $email.Trim()
  if (-not $emailClean) { continue }

  $signals = @()
  $riskScore = 0  # 0-100, higher = more at risk

  # --- Signal 1: Days since last activity (from API) ---
  $daysSinceLastEvent = $null
  $sessionCount = 0
  $totalEvents = 0
  $latestEventAt = $null
  $feedbackCount = 0
  $reachable = $true

  try {
    $loginBody = @{ email = $emailClean } | ConvertTo-Json
    $login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $loginBody
    $headers = @{ authorization = "Bearer $($login.token)" }
    $sessResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=$SessionLimit" -Headers $headers
    $sessions = $sessResp.sessions
    $sessionCount = if ($sessions) { $sessions.Count } else { 0 }

    $allTimestamps = @()
    foreach ($s in $sessions) {
      try {
        $evResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($s.id)/events" -Headers $headers
        if ($evResp.events) {
          $totalEvents += $evResp.events.Count
          $feedbackCount += @($evResp.events | Where-Object { $_.type -eq 'user_feedback' }).Count
          foreach ($e in $evResp.events) {
            if ($e.createdAt) { $allTimestamps += $e.createdAt }
          }
        }
      } catch {}
    }

    if ($allTimestamps.Count -gt 0) {
      $sorted = $allTimestamps | Sort-Object
      $latestEventAt = $sorted[-1]
      try {
        $latestDt = [DateTime]::Parse($latestEventAt)
        $todayDt = [DateTime]::Parse($today)
        $daysSinceLastEvent = [math]::Max(0, ($todayDt - $latestDt).Days)
      } catch {}
    }
  }
  catch {
    $reachable = $false
    $signals += [ordered]@{ signal = 'unreachable'; detail = 'Cannot login or query sessions'; weight = 30 }
    $riskScore += 30
  }

  if ($reachable) {
    # No sessions at all
    if ($sessionCount -eq 0) {
      $signals += [ordered]@{ signal = 'no_sessions'; detail = 'Partner has zero sessions'; weight = 40 }
      $riskScore += 40
    }

    # Days since last activity
    if ($null -ne $daysSinceLastEvent) {
      if ($daysSinceLastEvent -ge ($InactiveDaysThreshold * 2)) {
        $signals += [ordered]@{ signal = 'long_inactive'; detail = "No activity for $daysSinceLastEvent days"; weight = 35 }
        $riskScore += 35
      }
      elseif ($daysSinceLastEvent -ge $InactiveDaysThreshold) {
        $signals += [ordered]@{ signal = 'inactive'; detail = "No activity for $daysSinceLastEvent days"; weight = 20 }
        $riskScore += 20
      }
    }
    elseif ($sessionCount -gt 0) {
      # Has sessions but no parseable timestamps
      $signals += [ordered]@{ signal = 'no_timestamps'; detail = 'Events lack timestamps'; weight = 10 }
      $riskScore += 10
    }

    # Low session count
    if ($sessionCount -gt 0 -and $sessionCount -lt 3) {
      $signals += [ordered]@{ signal = 'few_sessions'; detail = "Only $sessionCount sessions (< 3)"; weight = 10 }
      $riskScore += 10
    }

    # No feedback given
    if ($feedbackCount -eq 0 -and $sessionCount -gt 0) {
      $signals += [ordered]@{ signal = 'no_feedback'; detail = 'Zero feedback items'; weight = 10 }
      $riskScore += 10
    }

    # --- Signal 2: Declining prompts in rolling history ---
    if ($historyRows.Count -gt 0) {
      $partnerHistory = @($historyRows | Where-Object { $_.email -eq $emailClean } | Sort-Object { $_.snapshotDate })

      if ($partnerHistory.Count -ge 2) {
        $recentHalf = [math]::Ceiling($partnerHistory.Count / 2)
        $olderRows = $partnerHistory[0..($partnerHistory.Count - $recentHalf - 1)]
        $newerRows = $partnerHistory[($partnerHistory.Count - $recentHalf)..($partnerHistory.Count - 1)]

        $olderAvgPrompts = ($olderRows | ForEach-Object { [int]$_.prompts24h } | Measure-Object -Average).Average
        $newerAvgPrompts = ($newerRows | ForEach-Object { [int]$_.prompts24h } | Measure-Object -Average).Average

        if ($olderAvgPrompts -gt 0) {
          $promptChange = ($newerAvgPrompts - $olderAvgPrompts) / $olderAvgPrompts
          if ($promptChange -le (-$DeclineThreshold)) {
            $pctDrop = [math]::Round([math]::Abs($promptChange) * 100)
            $signals += [ordered]@{ signal = 'prompt_decline'; detail = "Prompts dropped $pctDrop% (older avg=$([math]::Round($olderAvgPrompts,1)), newer avg=$([math]::Round($newerAvgPrompts,1)))"; weight = 20 }
            $riskScore += 20
          }
        }

        # Declining confirmations
        $olderAvgConf = ($olderRows | ForEach-Object { [int]$_.confirmed } | Measure-Object -Average).Average
        $newerAvgConf = ($newerRows | ForEach-Object { [int]$_.confirmed } | Measure-Object -Average).Average

        if ($olderAvgConf -gt 0) {
          $confChange = ($newerAvgConf - $olderAvgConf) / $olderAvgConf
          if ($confChange -le (-$DeclineThreshold)) {
            $pctDrop = [math]::Round([math]::Abs($confChange) * 100)
            $signals += [ordered]@{ signal = 'approval_decline'; detail = "Approvals dropped $pctDrop%"; weight = 15 }
            $riskScore += 15
          }
        }
      }
      elseif ($partnerHistory.Count -eq 0) {
        $signals += [ordered]@{ signal = 'no_history'; detail = 'No rows in rolling history'; weight = 5 }
        $riskScore += 5
      }
    }
  }

  # Cap at 100
  $riskScore = [math]::Min(100, $riskScore)

  # Determine risk level
  $riskLevel = if ($riskScore -ge 60) { 'critical' }
    elseif ($riskScore -ge 40) { 'high' }
    elseif ($riskScore -ge 20) { 'medium' }
    else { 'low' }

  $riskCounts[$riskLevel]++

  $partnerResults += [ordered]@{
    email            = $emailClean
    reachable        = $reachable
    riskScore        = $riskScore
    riskLevel        = $riskLevel
    sessionCount     = $sessionCount
    totalEvents      = $totalEvents
    feedbackCount    = $feedbackCount
    daysSinceLast    = $daysSinceLastEvent
    latestEventAt    = $latestEventAt
    signalCount      = $signals.Count
    signals          = $signals
  }
}

# ---------- aggregate ----------
$overallRisk = if ($riskCounts.critical -gt 0) { 'critical' }
  elseif ($riskCounts.high -gt 0) { 'high' }
  elseif ($riskCounts.medium -gt 0) { 'medium' }
  else { 'low' }

$atRiskPartners = @($partnerResults | Where-Object { $_.riskLevel -in @('high', 'critical') }).Count

$report = [ordered]@{
  generatedAt       = (Get-Date).ToString("o")
  date              = $today
  totalPartners     = $Emails.Count
  overallRisk       = $overallRisk
  atRiskPartners    = $atRiskPartners
  riskDistribution  = [ordered]@{
    critical = $riskCounts.critical
    high     = $riskCounts.high
    medium   = $riskCounts.medium
    low      = $riskCounts.low
  }
  thresholds        = [ordered]@{
    inactiveDays    = $InactiveDaysThreshold
    declinePercent  = [math]::Round($DeclineThreshold * 100)
  }
  partners          = $partnerResults
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $report | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host ('  [ok] Report written: ' + $OutputPath) -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $report | ConvertTo-Json -Depth 8
}
else {
  $overallColor = switch ($overallRisk) {
    'critical' { 'Red' }
    'high'     { 'DarkYellow' }
    'medium'   { 'Yellow' }
    default    { 'Green' }
  }
  Write-Host "  Overall risk: $overallRisk" -ForegroundColor $overallColor
  Write-Host "  At-risk partners: $atRiskPartners/$($Emails.Count)" -ForegroundColor White
  Write-Host "  Distribution: critical=$($riskCounts.critical) high=$($riskCounts.high) medium=$($riskCounts.medium) low=$($riskCounts.low)" -ForegroundColor Gray
  Write-Host ''

  foreach ($pr in $partnerResults) {
    $pColor = switch ($pr.riskLevel) {
      'critical' { 'Red' }
      'high'     { 'DarkYellow' }
      'medium'   { 'Yellow' }
      default    { 'Green' }
    }

    $riskBar = switch ($pr.riskLevel) {
      'critical' { '[!!!!]' }
      'high'     { '[!!! ]' }
      'medium'   { '[!!  ]' }
      default    { '[    ]' }
    }

    Write-Host "  $riskBar $($pr.email): $($pr.riskLevel) (score=$($pr.riskScore))" -ForegroundColor $pColor

    if (-not $pr.reachable) {
      Write-Host '         [unreachable]' -ForegroundColor Red
      continue
    }

    Write-Host "         sessions=$($pr.sessionCount)  events=$($pr.totalEvents)  feedback=$($pr.feedbackCount)" -ForegroundColor Gray

    if ($null -ne $pr.daysSinceLast) {
      $daysColor = if ($pr.daysSinceLast -ge $InactiveDaysThreshold) { 'DarkYellow' } else { 'Green' }
      Write-Host "         days since last: $($pr.daysSinceLast)" -ForegroundColor $daysColor
    }

    if ($pr.signals.Count -gt 0) {
      Write-Host '         signals:' -ForegroundColor White
      foreach ($sig in $pr.signals) {
        $sigIcon = if ($sig.weight -ge 20) { '!' } else { '~' }
        Write-Host "           $sigIcon $($sig.signal): $($sig.detail) (w=$($sig.weight))" -ForegroundColor $pColor
      }
    }
    else {
      Write-Host '         No risk signals detected' -ForegroundColor Green
    }
    Write-Host ''
  }

  # Action items
  $highRisk = @($partnerResults | Where-Object { $_.riskLevel -in @('high', 'critical') })
  if ($highRisk.Count -gt 0) {
    Write-Host '  Action required:' -ForegroundColor Red
    foreach ($hr in $highRisk) {
      $topSignal = if ($hr.signals.Count -gt 0) { $hr.signals[0].signal } else { 'unknown' }
      Write-Host "    -> Reach out to $($hr.email) (top signal: $topSignal)" -ForegroundColor Red
    }
    Write-Host ''
  }
}

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host ('  [warn] Triage file not found: ' + $triagePath) -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $perPartnerStr = ($partnerResults | ForEach-Object {
    "$($_.email)=$($_.riskLevel)($($_.riskScore))"
  }) -join ', '

  $line = "$stamp Churn risk: overall=$overallRisk, atRisk=$atRiskPartners/$($Emails.Count), partners=($perPartnerStr)"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Churn risk:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted churn risk in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] Churn risk detection complete.' -ForegroundColor Green
