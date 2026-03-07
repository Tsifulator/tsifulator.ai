param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [switch]$Json,
  [switch]$AppendToTriage,
  [string]$OutputPath = "",
  [double]$Threshold = 20   # % change to flag as drift
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath  = "docs/daily-triage/$today.md"
$historyPath = "docs/reports/partner-compare-history.csv"

# Normalize emails
$normalizedEmails = @()
foreach ($entry in $Emails) {
  $normalizedEmails += ($entry -split '[,;\s]+' | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
}
$Emails = $normalizedEmails | Select-Object -Unique

Write-Host '== tsifulator.ai daily drift detector ==' -ForegroundColor Cyan
Write-Host "Date: $today"
Write-Host "Threshold: $Threshold%"
Write-Host "Partners: $($Emails.Count)"

# ---------- load history ----------
if (-not (Test-Path $historyPath)) {
  Write-Host '[warn] No rolling history found — cannot detect drift.' -ForegroundColor DarkYellow
  $summary = [ordered]@{
    date       = $today
    status     = 'no-history'
    threshold  = $Threshold
    partners   = @()
    drifts     = 0
    health     = 'unknown'
  }
  if ($Json) { $summary | ConvertTo-Json -Depth 5; exit 0 }
  Write-Host 'done.' -ForegroundColor Green
  exit 0
}

$allRows = Import-Csv -Path $historyPath
$dates = @($allRows | Select-Object -ExpandProperty snapshotDate -Unique | Sort-Object)

# For each date we take the LATEST snapshot (by snapshotAt) per partner
function Get-LatestSnapshot {
  param([object[]]$Rows, [string]$TargetDate, [string]$Email)
  $matching = @($Rows | Where-Object { $_.snapshotDate -eq $TargetDate -and $_.email -eq $Email } |
    Sort-Object -Property snapshotAt -Descending)
  if ($matching.Count -gt 0) { return $matching[0] }
  return $null
}

# Find today and previous date
$todayIdx = [array]::IndexOf($dates, $today)
$prevDate = $null
if ($todayIdx -gt 0) {
  $prevDate = $dates[$todayIdx - 1]
} elseif ($todayIdx -eq -1 -and $dates.Count -ge 2) {
  # today not in history yet — compare last two dates
  $prevDate = $dates[$dates.Count - 2]
  $today = $dates[$dates.Count - 1]
  Write-Host "[info] Target date not in history; comparing $prevDate -> $today" -ForegroundColor DarkYellow
}

$metrics = @('prompts24h', 'proposed', 'confirmed', 'blocked', 'streamRequests', 'streamCompletions', 'medianChatLatencyMs', 'medianStreamFirstTokenLatencyMs')

# ---------- analyse ----------
$partnerDrifts = @()
$totalDrifts = 0

foreach ($email in $Emails) {
  $todaySnap = Get-LatestSnapshot -Rows $allRows -TargetDate $today -Email $email
  $prevSnap  = if ($prevDate) { Get-LatestSnapshot -Rows $allRows -TargetDate $prevDate -Email $email } else { $null }

  $driftItems = @()

  if ($null -eq $todaySnap) {
    $partnerDrifts += [ordered]@{
      email     = $email
      status    = 'no-today-data'
      drifts    = @()
      driftCount = 0
    }
    continue
  }

  if ($null -eq $prevSnap) {
    # First day — no comparison baseline
    $partnerDrifts += [ordered]@{
      email     = $email
      status    = 'first-day'
      today     = [ordered]@{}
      drifts    = @()
      driftCount = 0
    }
    # Still populate today values
    $todayVals = [ordered]@{}
    foreach ($m in $metrics) {
      $todayVals[$m] = [double]$todaySnap.$m
    }
    $partnerDrifts[-1]['today'] = $todayVals
    continue
  }

  # Compare each metric
  $todayVals = [ordered]@{}
  $prevVals = [ordered]@{}
  foreach ($m in $metrics) {
    $tVal = [double]$todaySnap.$m
    $pVal = [double]$prevSnap.$m
    $todayVals[$m] = $tVal
    $prevVals[$m]  = $pVal

    # Calculate % change
    $delta = $tVal - $pVal
    $pctChange = if ($pVal -ne 0) {
      [math]::Round(($delta / [math]::Abs($pVal)) * 100, 1)
    } elseif ($tVal -ne 0) {
      # From 0 to non-zero = 100% spike
      100.0
    } else {
      0.0
    }

    $direction = if ($delta -gt 0) { 'up' } elseif ($delta -lt 0) { 'down' } else { 'flat' }

    if ([math]::Abs($pctChange) -ge $Threshold -and $direction -ne 'flat') {
      $severity = if ([math]::Abs($pctChange) -ge 50) { 'major' } else { 'minor' }
      $driftItems += [ordered]@{
        metric    = $m
        prev      = $pVal
        current   = $tVal
        delta     = $delta
        pctChange = $pctChange
        direction = $direction
        severity  = $severity
      }
      $totalDrifts++
    }
  }

  $partnerDrifts += [ordered]@{
    email      = $email
    status     = 'compared'
    prevDate   = $prevDate
    today      = $todayVals
    prev       = $prevVals
    drifts     = $driftItems
    driftCount = $driftItems.Count
  }
}

# ---------- fleet summary ----------
$comparedPartners = @($partnerDrifts | Where-Object { $_['status'] -eq 'compared' })
$partnersWithDrift = @($comparedPartners | Where-Object { $_['driftCount'] -gt 0 }).Count

$health = if ($comparedPartners.Count -eq 0) { 'no-baseline' }
          elseif ($totalDrifts -eq 0) { 'stable' }
          elseif ($totalDrifts -le 2) { 'minor-drift' }
          elseif ($totalDrifts -le 5) { 'moderate-drift' }
          else { 'significant-drift' }

$summary = [ordered]@{
  date              = $today
  prevDate          = $prevDate
  threshold         = $Threshold
  status            = if ($prevDate) { 'compared' } else { 'no-baseline' }
  totalDrifts       = $totalDrifts
  partnersWithDrift = $partnersWithDrift
  partnersCompared  = $comparedPartners.Count
  health            = $health
  partners          = [array]$partnerDrifts
}

# ---------- output ----------
if ($Json) {
  $summary | ConvertTo-Json -Depth 6
  exit 0
}

if ($OutputPath) {
  $summary | ConvertTo-Json -Depth 6 | Set-Content -Path $OutputPath -Encoding utf8
  Write-Host "[saved] $OutputPath"
}

# ---------- display ----------
Write-Host ''
Write-Host '--- Fleet Drift Summary ---' -ForegroundColor Cyan
if (-not $prevDate) {
  Write-Host '  No previous date available — first day, no drift comparison.' -ForegroundColor DarkYellow
} else {
  Write-Host "  Comparing: $prevDate -> $today"
  Write-Host "  Partners compared: $($comparedPartners.Count)"
  Write-Host "  Total drifts: $totalDrifts"
  Write-Host "  Partners with drift: $partnersWithDrift"
  $healthStr = "  Health: $health"
  if ($health -eq 'stable') {
    Write-Host $healthStr -ForegroundColor Green
  } elseif ($health -eq 'significant-drift') {
    Write-Host $healthStr -ForegroundColor Red
  } else {
    Write-Host $healthStr -ForegroundColor Yellow
  }
}

Write-Host ''
Write-Host '--- Per-Partner Detail ---' -ForegroundColor Cyan
foreach ($p in $partnerDrifts) {
  Write-Host "  $($p['email'])" -ForegroundColor White
  if ($p['status'] -eq 'first-day') {
    Write-Host '    First day — no baseline for comparison' -ForegroundColor DarkYellow
    foreach ($m in $metrics) {
      if ($p['today'].Contains($m)) {
        Write-Host "    $m = $($p['today'][$m])"
      }
    }
  } elseif ($p['status'] -eq 'no-today-data') {
    Write-Host '    No data for today' -ForegroundColor DarkYellow
  } elseif ($p['status'] -eq 'compared') {
    if ($p['driftCount'] -eq 0) {
      Write-Host '    No significant drift detected' -ForegroundColor Green
    } else {
      Write-Host "    Drifts: $($p['driftCount'])" -ForegroundColor Yellow
      foreach ($d in $p['drifts']) {
        $arrow = if ($d['direction'] -eq 'up') { '^' } else { 'v' }
        $color = if ($d['severity'] -eq 'major') { 'Red' } else { 'Yellow' }
        $sign = if ($d['pctChange'] -gt 0) { '+' } else { '' }
        Write-Host "      $arrow $($d['metric']): $($d['prev']) -> $($d['current']) ($sign$($d['pctChange'])%) [$($d['severity'])]" -ForegroundColor $color
      }
    }
  }
  Write-Host ''
}

# ---------- triage ----------
if ($AppendToTriage) {
  if (-not (Test-Path (Split-Path $triagePath -Parent))) {
    New-Item -ItemType Directory -Path (Split-Path $triagePath -Parent) -Force | Out-Null
  }
  $statusStr = if ($prevDate) { "compared=$prevDate" } else { 'no-baseline' }
  $triageLine = "Daily drift detector: $statusStr drifts=$totalDrifts partnersWithDrift=$partnersWithDrift health=$health threshold=$Threshold%"
  if (Test-Path $triagePath) {
    $existing = Get-Content -Path $triagePath -Raw
    $cleaned = ($existing -split "`n" | Where-Object { $_ -notmatch '^Daily drift detector:' }) -join "`n"
    Set-Content -Path $triagePath -Value $cleaned.TrimEnd() -Encoding utf8
  }
  Add-Content -Path $triagePath -Value $triageLine -Encoding utf8
  Write-Host "[triage] appended to $triagePath" -ForegroundColor Green
}

Write-Host 'done.' -ForegroundColor Green
