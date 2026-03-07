param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [switch]$Json,
  [switch]$AppendToTriage
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

Write-Host '== tsifulator.ai retention curve analyzer ==' -ForegroundColor Cyan
Write-Host "Date:     $today"
Write-Host "Partners: $($Emails -join ', ')"

# API health check
try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health" -TimeoutSec 5
  if ($health.status -ne 'ok') { throw 'bad status' }
  Write-Host '[ok] API healthy: ok' -ForegroundColor Green
} catch {
  Write-Host '[FAIL] API unreachable' -ForegroundColor Red
  exit 1
}

# ---------- collect session timestamps per partner ----------
$partnerResults = @()
$allSessions = @()

foreach ($em in $Emails) {
  $pr = [ordered]@{
    email            = $em
    reachable        = $false
    totalSessions    = 0
    firstSessionDate = $null
    lastSessionDate  = $null
    tenureDays       = 0
    avgDaysBetween   = 0
    medianDaysBetween = 0
    maxGapDays       = 0
    sessionsPerWeek  = 0
    returnRate       = 0
    cadenceLabel     = 'unknown'
    activeDays       = @()
    gaps             = @()
    trend            = 'unknown'
  }

  try {
    $lb = @{ email = $em } | ConvertTo-Json
    $lg = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $lb
    $hd = @{ authorization = "Bearer $($lg.token)" }
    $sess = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=100" -Headers $hd
    $pr.reachable = $true

    $sessionDates = @()
    if ($sess.sessions) {
      foreach ($s in $sess.sessions) {
        if ($s.createdAt) {
          try {
            $dt = [DateTime]::Parse($s.createdAt)
            $sessionDates += $dt
          } catch {}
        }
      }
    }

    $pr.totalSessions = $sessionDates.Count

    if ($sessionDates.Count -gt 0) {
      # Sort chronologically
      $sessionDates = $sessionDates | Sort-Object
      $pr.firstSessionDate = $sessionDates[0].ToString('yyyy-MM-dd')
      $pr.lastSessionDate  = $sessionDates[-1].ToString('yyyy-MM-dd')

      # Unique active days
      $uniqueDays = @($sessionDates | ForEach-Object { $_.ToString('yyyy-MM-dd') } | Select-Object -Unique)
      $pr.activeDays = $uniqueDays

      # Tenure: days from first to today
      $todayDt = [DateTime]::Parse($today)
      $pr.tenureDays = [math]::Max(1, [math]::Floor(($todayDt - $sessionDates[0]).TotalDays))

      # Sessions per week
      $weeks = [math]::Max(1, $pr.tenureDays / 7.0)
      $pr.sessionsPerWeek = [math]::Round($sessionDates.Count / $weeks, 1)

      # Gaps between unique active days
      if ($uniqueDays.Count -gt 1) {
        $gaps = @()
        for ($i = 1; $i -lt $uniqueDays.Count; $i++) {
          $d1 = [DateTime]::Parse($uniqueDays[$i - 1])
          $d2 = [DateTime]::Parse($uniqueDays[$i])
          $gapDays = [math]::Floor(($d2 - $d1).TotalDays)
          $gaps += $gapDays
        }
        $pr.gaps = $gaps
        $pr.avgDaysBetween = [math]::Round(($gaps | Measure-Object -Average).Average, 1)
        $pr.maxGapDays = ($gaps | Measure-Object -Maximum).Maximum

        # Median
        $sorted = $gaps | Sort-Object
        $mid = [math]::Floor($sorted.Count / 2)
        if ($sorted.Count % 2 -eq 0) {
          $pr.medianDaysBetween = [math]::Round(($sorted[$mid - 1] + $sorted[$mid]) / 2.0, 1)
        } else {
          $pr.medianDaysBetween = $sorted[$mid]
        }
      }

      # Return rate: proportion of tenure days that had activity
      $pr.returnRate = [math]::Round($uniqueDays.Count / $pr.tenureDays, 2)

      # Cadence label
      $pr.cadenceLabel = if ($pr.sessionsPerWeek -ge 5) { 'daily' }
                         elseif ($pr.sessionsPerWeek -ge 2) { 'regular' }
                         elseif ($pr.sessionsPerWeek -ge 0.5) { 'weekly' }
                         elseif ($pr.sessionsPerWeek -gt 0) { 'sporadic' }
                         else { 'inactive' }

      # Trend: compare first-half vs second-half session density
      if ($uniqueDays.Count -ge 3) {
        $halfIdx = [math]::Floor($uniqueDays.Count / 2)
        $firstHalfDays = @($uniqueDays[0..($halfIdx - 1)])
        $secondHalfDays = @($uniqueDays[$halfIdx..($uniqueDays.Count - 1)])

        $fhSpan = [math]::Max(1, ([DateTime]::Parse($firstHalfDays[-1]) - [DateTime]::Parse($firstHalfDays[0])).TotalDays + 1)
        $shSpan = [math]::Max(1, ([DateTime]::Parse($secondHalfDays[-1]) - [DateTime]::Parse($secondHalfDays[0])).TotalDays + 1)

        $fhDensity = $firstHalfDays.Count / $fhSpan
        $shDensity = $secondHalfDays.Count / $shSpan

        if ($shDensity -gt $fhDensity * 1.2) { $pr.trend = 'increasing' }
        elseif ($shDensity -lt $fhDensity * 0.8) { $pr.trend = 'decreasing' }
        else { $pr.trend = 'stable' }
      } elseif ($uniqueDays.Count -ge 1) {
        $pr.trend = 'new'
      }

      # Days since last session
      $daysSinceLast = [math]::Floor(($todayDt - $sessionDates[-1]).TotalDays)
      $pr['daysSinceLastSession'] = $daysSinceLast
    }

    $allSessions += $sessionDates
  } catch {
    Write-Host "  [warn] Could not reach $em" -ForegroundColor DarkYellow
  }

  $partnerResults += $pr
}

# ---------- global metrics ----------
$reachable = @($partnerResults | Where-Object { $_.reachable })
$globalTotalSessions = ($reachable | ForEach-Object { $_.totalSessions } | Measure-Object -Sum).Sum
$globalAvgGap = 0
$globalAvgCadence = 0
$globalRetention = 'unknown'

if ($reachable.Count -gt 0) {
  $globalAvgCadence = [math]::Round(($reachable | ForEach-Object { $_.sessionsPerWeek } | Measure-Object -Average).Average, 1)
  $gapsAll = @()
  foreach ($pr in $reachable) {
    if ($pr.gaps.Count -gt 0) { $gapsAll += $pr.gaps }
  }
  if ($gapsAll.Count -gt 0) {
    $globalAvgGap = [math]::Round(($gapsAll | Measure-Object -Average).Average, 1)
  }

  # Global retention label
  $atRisk = @($reachable | Where-Object { $_.trend -eq 'decreasing' -or $_.cadenceLabel -eq 'sporadic' -or $_.cadenceLabel -eq 'inactive' }).Count
  $globalRetention = if ($atRisk -eq 0) { 'healthy' }
                     elseif ($atRisk -le [math]::Floor($reachable.Count / 2)) { 'mixed' }
                     else { 'at_risk' }
}

# ---------- build output ----------
$output = [ordered]@{
  generatedAt   = (Get-Date).ToString('o')
  date          = $today
  retention     = $globalRetention
  totalSessions = $globalTotalSessions
  avgSessionsPerWeek = $globalAvgCadence
  avgDaysBetweenVisits = $globalAvgGap
  partners      = @()
}

foreach ($pr in $partnerResults) {
  $p = [ordered]@{
    email              = $pr.email
    reachable          = $pr.reachable
    totalSessions      = $pr.totalSessions
    firstSession       = $pr.firstSessionDate
    lastSession        = $pr.lastSessionDate
    tenureDays         = $pr.tenureDays
    daysSinceLastSession = $pr['daysSinceLastSession']
    sessionsPerWeek    = $pr.sessionsPerWeek
    avgDaysBetween     = $pr.avgDaysBetween
    medianDaysBetween  = $pr.medianDaysBetween
    maxGapDays         = $pr.maxGapDays
    returnRate         = $pr.returnRate
    cadence            = $pr.cadenceLabel
    trend              = $pr.trend
    activeDays         = $pr.activeDays
  }
  $output.partners += $p
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $output | ConvertTo-Json -Depth 8
} else {
  # Global summary
  $retColor = switch ($globalRetention) {
    'healthy' { 'Green' }
    'mixed'   { 'Yellow' }
    'at_risk' { 'Red' }
    default   { 'DarkGray' }
  }
  Write-Host "  Retention: $globalRetention ($globalTotalSessions sessions across $($reachable.Count) partners)" -ForegroundColor $retColor
  Write-Host "  Avg sessions/week: $globalAvgCadence   Avg gap: $($globalAvgGap)d" -ForegroundColor White
  Write-Host ''

  # Per-partner table
  Write-Host '  Per partner:' -ForegroundColor White
  $hdr = '  {0,-25} {1,4} {2,6} {3,6} {4,5} {5,7} {6,9} {7,8}' -f 'EMAIL','SESS','S/WK','GAP-D','MAXG','RETURN','CADENCE','TREND'
  Write-Host $hdr -ForegroundColor DarkGray

  foreach ($pr in $partnerResults) {
    if (-not $pr.reachable) {
      Write-Host "  $($pr.email): [unreachable]" -ForegroundColor Red
      continue
    }

    $tColor = switch ($pr.trend) {
      'increasing' { 'Green' }
      'stable'     { 'Green' }
      'new'        { 'Cyan' }
      'decreasing' { 'Red' }
      default      { 'DarkGray' }
    }

    $shortEmail = $pr.email
    if ($shortEmail.Length -gt 25) { $shortEmail = $shortEmail.Substring(0, 22) + '...' }
    $row = '  {0,-25} {1,4} {2,6} {3,6} {4,5} {5,7} {6,9} {7,8}' -f $shortEmail, $pr.totalSessions, $pr.sessionsPerWeek, $pr.avgDaysBetween, $pr.maxGapDays, $pr.returnRate, $pr.cadenceLabel, $pr.trend
    Write-Host $row -ForegroundColor $tColor

    # Timeline sparkline: show active days as dots on a calendar line
    if ($pr.activeDays.Count -gt 0 -and $pr.tenureDays -le 90) {
      $firstDt = [DateTime]::Parse($pr.activeDays[0])
      $span = [math]::Min($pr.tenureDays, 60)
      $timeline = ''
      for ($d = 0; $d -lt $span; $d++) {
        $checkDate = $firstDt.AddDays($d).ToString('yyyy-MM-dd')
        if ($pr.activeDays -contains $checkDate) { $timeline += '#' }
        else { $timeline += '.' }
      }
      Write-Host "    timeline: [$timeline]" -ForegroundColor DarkGray
    }
  }

  # Insights
  Write-Host ''
  Write-Host '  Insights:' -ForegroundColor White
  $decreasing = @($partnerResults | Where-Object { $_.trend -eq 'decreasing' })
  $sporadic = @($partnerResults | Where-Object { $_.cadenceLabel -eq 'sporadic' -or $_.cadenceLabel -eq 'inactive' })
  $longGap = @($partnerResults | Where-Object { $_['daysSinceLastSession'] -ge 7 })

  if ($decreasing.Count -gt 0) {
    Write-Host "    ! $($decreasing.Count) partner(s) with decreasing activity trend" -ForegroundColor Red
  }
  if ($sporadic.Count -gt 0) {
    Write-Host "    ! $($sporadic.Count) partner(s) with sporadic/inactive cadence" -ForegroundColor DarkYellow
  }
  if ($longGap.Count -gt 0) {
    $names = ($longGap | ForEach-Object { $_.email }) -join ', '
    Write-Host "    ~ $($longGap.Count) partner(s) not seen in 7+ days: $names" -ForegroundColor DarkYellow
  }
  if ($globalRetention -eq 'healthy') {
    Write-Host '    + All partners showing healthy return patterns' -ForegroundColor Green
  }
  if ($globalAvgCadence -ge 3) {
    Write-Host '    + Strong weekly cadence across the cohort' -ForegroundColor Green
  }
}

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host ('  [warn] Triage file not found: ' + $triagePath) -ForegroundColor DarkYellow
    exit 1
  }

  $cadences = ($partnerResults | Where-Object { $_.reachable } | ForEach-Object { $_.cadenceLabel }) -join '/'
  $trends = ($partnerResults | Where-Object { $_.reachable } | ForEach-Object { $_.trend }) -join '/'

  $stamp = (Get-Date).ToString('o')
  $line = "$stamp Retention curve: status=$globalRetention, sessions=$globalTotalSessions, avgPerWeek=$globalAvgCadence, avgGap=$($globalAvgGap)d, cadence=$cadences, trend=$trends"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Retention curve:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ''
  Write-Host ('  [ok] Upserted retention curve in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host ''
Write-Host '[done] Retention curve analysis complete.' -ForegroundColor Green
