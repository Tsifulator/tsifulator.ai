param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [int]$SessionLimit = 100,
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

Write-Host '== tsifulator.ai engagement heatmap ==' -ForegroundColor Cyan
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

# ---------- day-of-week labels ----------
$dowLabels = @('Mon','Tue','Wed','Thu','Fri','Sat','Sun')

# ---------- collect events per partner ----------
$partnerResults = @()
$globalHourBuckets = @{}     # hour -> count
$globalDowBuckets = @{}      # dow  -> count
$globalHourDow = @{}         # "dow:hour" -> count
$totalEvents = 0

foreach ($email in $Emails) {
  $emailClean = $email.Trim()
  if (-not $emailClean) { continue }

  $hourBuckets = @{}
  $dowBuckets = @{}
  $eventCount = 0
  $sessionCount = 0
  $firstAt = $null
  $latestAt = $null

  try {
    $loginBody = @{ email = $emailClean } | ConvertTo-Json
    $login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $loginBody
    $headers = @{ authorization = "Bearer $($login.token)" }
  }
  catch {
    Write-Host ('  [warn] Could not login as ' + $emailClean) -ForegroundColor DarkYellow
    $partnerResults += [ordered]@{
      email        = $emailClean
      reachable    = $false
      totalEvents  = 0
      sessions     = 0
      peakHour     = $null
      peakDay      = $null
      hourBreakdown = @{}
      dowBreakdown  = @{}
    }
    continue
  }

  try {
    $sessionsResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=$SessionLimit" -Headers $headers
    $sessions = $sessionsResp.sessions
    $sessionCount = if ($sessions) { $sessions.Count } else { 0 }
  }
  catch { $sessions = @() }

  foreach ($session in $sessions) {
    try {
      $eventsResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($session.id)/events" -Headers $headers
      $events = $eventsResp.events
    }
    catch { continue }

    if (-not $events) { continue }

    foreach ($evt in $events) {
      if (-not $evt.createdAt) { continue }

      try {
        $ts = [DateTime]::Parse($evt.createdAt)
      }
      catch { continue }

      $eventCount++
      $totalEvents++

      $hour = $ts.Hour
      $dow = [int]$ts.DayOfWeek
      # Convert Sunday=0 to Monday-first: Mon=0..Sun=6
      $dowIdx = if ($dow -eq 0) { 6 } else { $dow - 1 }

      # Per-partner buckets
      if (-not $hourBuckets.ContainsKey($hour)) { $hourBuckets[$hour] = 0 }
      $hourBuckets[$hour]++

      if (-not $dowBuckets.ContainsKey($dowIdx)) { $dowBuckets[$dowIdx] = 0 }
      $dowBuckets[$dowIdx]++

      # Global buckets
      if (-not $globalHourBuckets.ContainsKey($hour)) { $globalHourBuckets[$hour] = 0 }
      $globalHourBuckets[$hour]++

      if (-not $globalDowBuckets.ContainsKey($dowIdx)) { $globalDowBuckets[$dowIdx] = 0 }
      $globalDowBuckets[$dowIdx]++

      $hdKey = "$dowIdx`:$hour"
      if (-not $globalHourDow.ContainsKey($hdKey)) { $globalHourDow[$hdKey] = 0 }
      $globalHourDow[$hdKey]++

      # Track first/latest
      if ($null -eq $firstAt -or $ts -lt $firstAt) { $firstAt = $ts }
      if ($null -eq $latestAt -or $ts -gt $latestAt) { $latestAt = $ts }
    }
  }

  # Find peak hour and day
  $peakHour = $null
  $peakHourCount = 0
  foreach ($h in $hourBuckets.Keys) {
    if ($hourBuckets[$h] -gt $peakHourCount) {
      $peakHourCount = $hourBuckets[$h]
      $peakHour = $h
    }
  }

  $peakDow = $null
  $peakDowCount = 0
  foreach ($d in $dowBuckets.Keys) {
    if ($dowBuckets[$d] -gt $peakDowCount) {
      $peakDowCount = $dowBuckets[$d]
      $peakDow = $d
    }
  }

  # Build sorted hour breakdown for JSON
  $hourBreakdownSorted = [ordered]@{}
  for ($h = 0; $h -lt 24; $h++) {
    $hourBreakdownSorted["$h"] = if ($hourBuckets.ContainsKey($h)) { $hourBuckets[$h] } else { 0 }
  }

  $dowBreakdownSorted = [ordered]@{}
  for ($d = 0; $d -lt 7; $d++) {
    $dowBreakdownSorted[$dowLabels[$d]] = if ($dowBuckets.ContainsKey($d)) { $dowBuckets[$d] } else { 0 }
  }

  $partnerResults += [ordered]@{
    email         = $emailClean
    reachable     = $true
    totalEvents   = $eventCount
    sessions      = $sessionCount
    peakHour      = if ($null -ne $peakHour) { "${peakHour}:00" } else { 'n/a' }
    peakDay       = if ($null -ne $peakDow) { $dowLabels[$peakDow] } else { 'n/a' }
    firstEventAt  = if ($firstAt) { $firstAt.ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
    latestEventAt = if ($latestAt) { $latestAt.ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
    hourBreakdown = $hourBreakdownSorted
    dowBreakdown  = $dowBreakdownSorted
  }
}

# ---------- global peaks ----------
$globalPeakHour = $null; $globalPeakHourCount = 0
foreach ($h in $globalHourBuckets.Keys) {
  if ($globalHourBuckets[$h] -gt $globalPeakHourCount) {
    $globalPeakHourCount = $globalHourBuckets[$h]
    $globalPeakHour = $h
  }
}

$globalPeakDow = $null; $globalPeakDowCount = 0
foreach ($d in $globalDowBuckets.Keys) {
  if ($globalDowBuckets[$d] -gt $globalPeakDowCount) {
    $globalPeakDowCount = $globalDowBuckets[$d]
    $globalPeakDow = $d
  }
}

# Quiet hours: hours 0-23 with 0 global events
$quietHours = @()
for ($h = 0; $h -lt 24; $h++) {
  if (-not $globalHourBuckets.ContainsKey($h) -or $globalHourBuckets[$h] -eq 0) {
    $quietHours += "${h}:00"
  }
}

# Global hour breakdown
$globalHourSorted = [ordered]@{}
for ($h = 0; $h -lt 24; $h++) {
  $globalHourSorted["$h"] = if ($globalHourBuckets.ContainsKey($h)) { $globalHourBuckets[$h] } else { 0 }
}

$globalDowSorted = [ordered]@{}
for ($d = 0; $d -lt 7; $d++) {
  $globalDowSorted[$dowLabels[$d]] = if ($globalDowBuckets.ContainsKey($d)) { $globalDowBuckets[$d] } else { 0 }
}

$report = [ordered]@{
  generatedAt     = (Get-Date).ToString("o")
  date            = $today
  totalEvents     = $totalEvents
  totalPartners   = $Emails.Count
  globalPeakHour  = if ($null -ne $globalPeakHour) { "${globalPeakHour}:00" } else { 'n/a' }
  globalPeakDay   = if ($null -ne $globalPeakDow) { $dowLabels[$globalPeakDow] } else { 'n/a' }
  quietHoursCount = $quietHours.Count
  quietHours      = $quietHours
  globalHourBreakdown = $globalHourSorted
  globalDowBreakdown  = $globalDowSorted
  partners        = $partnerResults
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $report | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host ('  [ok] Heatmap written: ' + $OutputPath) -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $report | ConvertTo-Json -Depth 8
}
else {
  Write-Host "  Total events: $totalEvents   Partners: $($Emails.Count)" -ForegroundColor White
  Write-Host "  Global peak hour: $($report.globalPeakHour)   Peak day: $($report.globalPeakDay)" -ForegroundColor White
  if ($quietHours.Count -gt 0 -and $quietHours.Count -le 20) {
    Write-Host "  Quiet hours ($($quietHours.Count)): $($quietHours -join ', ')" -ForegroundColor DarkGray
  }
  elseif ($quietHours.Count -gt 20) {
    Write-Host "  Quiet hours: $($quietHours.Count) of 24 (low activity)" -ForegroundColor DarkGray
  }
  Write-Host ''

  # --- text heatmap: hour-of-day bar chart ---
  Write-Host '  Hour-of-day (all partners):' -ForegroundColor White
  $maxHourVal = ($globalHourSorted.Values | Measure-Object -Maximum).Maximum
  if ($maxHourVal -eq 0) { $maxHourVal = 1 }
  $barMax = 30
  for ($h = 0; $h -lt 24; $h++) {
    $val = $globalHourSorted["$h"]
    $barLen = [math]::Floor($val / $maxHourVal * $barMax)
    $bar = '#' * $barLen
    $hourLabel = '{0,2}:00' -f $h
    $valLabel = if ($val -gt 0) { " $val" } else { '' }
    $color = if ($val -eq $maxHourVal -and $val -gt 0) { 'Green' } elseif ($val -gt 0) { 'Cyan' } else { 'DarkGray' }
    Write-Host "    $hourLabel |$bar$valLabel" -ForegroundColor $color
  }
  Write-Host ''

  # --- day-of-week chart ---
  Write-Host '  Day-of-week (all partners):' -ForegroundColor White
  $maxDowVal = ($globalDowSorted.Values | Measure-Object -Maximum).Maximum
  if ($maxDowVal -eq 0) { $maxDowVal = 1 }
  foreach ($day in $dowLabels) {
    $val = $globalDowSorted[$day]
    $barLen = [math]::Floor($val / $maxDowVal * $barMax)
    $bar = '#' * $barLen
    $valLabel = if ($val -gt 0) { " $val" } else { '' }
    $color = if ($val -eq $maxDowVal -and $val -gt 0) { 'Green' } elseif ($val -gt 0) { 'Cyan' } else { 'DarkGray' }
    Write-Host "    $day |$bar$valLabel" -ForegroundColor $color
  }
  Write-Host ''

  # --- per-partner summary ---
  Write-Host '  Per partner:' -ForegroundColor White
  foreach ($pr in $partnerResults) {
    if (-not $pr.reachable) {
      Write-Host "    $($pr.email): [unreachable]" -ForegroundColor Red
      continue
    }
    Write-Host "    $($pr.email): $($pr.totalEvents) events, $($pr.sessions) sessions" -ForegroundColor White
    Write-Host "      Peak: $($pr.peakHour) ($($pr.peakDay))  Activity: $($pr.firstEventAt) .. $($pr.latestEventAt)" -ForegroundColor Gray

    # Mini hour sparkline for this partner
    $partnerMaxH = 0
    for ($h = 0; $h -lt 24; $h++) {
      $v = $pr.hourBreakdown["$h"]
      if ($v -gt $partnerMaxH) { $partnerMaxH = $v }
    }
    if ($partnerMaxH -gt 0) {
      $sparkChars = @()
      for ($h = 0; $h -lt 24; $h++) {
        $v = $pr.hourBreakdown["$h"]
        $level = [math]::Floor($v / $partnerMaxH * 4)
        $ch = switch ($level) {
          0 { '.' }
          1 { '_' }
          2 { '=' }
          3 { '#' }
          4 { '#' }
          default { '.' }
        }
        $sparkChars += $ch
      }
      $sparkStr = $sparkChars -join ''
      Write-Host "      Hours: [$sparkStr]" -ForegroundColor Cyan
      Write-Host '              0         12        23' -ForegroundColor DarkGray
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
  $peakStr = "$($report.globalPeakHour) $($report.globalPeakDay)"
  $perPartnerStr = ($partnerResults | ForEach-Object {
    "$($_.email)=$($_.totalEvents)ev"
  }) -join ', '

  $line = "$stamp Engagement heatmap: events=$totalEvents, peak=$peakStr, quietHours=$($quietHours.Count), partners=($perPartnerStr)"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Engagement heatmap:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted engagement heatmap in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] Engagement heatmap complete.' -ForegroundColor Green
