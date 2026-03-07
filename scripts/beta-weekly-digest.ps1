param(
  [string]$HistoryPath = "docs/reports/partner-compare-history.csv",
  [string]$Date,
  [switch]$Json,
  [switch]$AppendToTriage,
  [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"
$todayDate = [datetime]::ParseExact($today, "yyyy-MM-dd", $null)

# Current week = last 7 days ending at $today (inclusive)
$weekStart = $todayDate.AddDays(-6)
$weekStartStr = $weekStart.ToString("yyyy-MM-dd")

# Previous week = 7 days before that
$prevWeekEnd = $todayDate.AddDays(-7)
$prevWeekStart = $todayDate.AddDays(-13)
$prevWeekEndStr = $prevWeekEnd.ToString("yyyy-MM-dd")
$prevWeekStartStr = $prevWeekStart.ToString("yyyy-MM-dd")

Write-Host '== tsifulator.ai weekly digest ==' -ForegroundColor Cyan
Write-Host "Date:         $today"
Write-Host "Current week: $weekStartStr .. $today"
Write-Host "Prior week:   $prevWeekStartStr .. $prevWeekEndStr"

if (-not (Test-Path $HistoryPath)) {
  Write-Host ('  [warn] History file not found: ' + $HistoryPath) -ForegroundColor DarkYellow
  Write-Host 'Run: npm run beta:compare:partners:csv:append' -ForegroundColor DarkYellow
  exit 1
}

$allRows = Import-Csv -Path $HistoryPath
if (-not $allRows -or $allRows.Count -eq 0) {
  Write-Host ('  [warn] History file is empty: ' + $HistoryPath) -ForegroundColor DarkYellow
  exit 1
}

$toInt = {
  param($value)
  if ($null -eq $value -or $value -eq '') { return 0 }
  return [int]$value
}

$toDouble = {
  param($value)
  if ($null -eq $value -or $value -eq '') { return 0.0 }
  return [double]$value
}

# Filter rows by week
$currentWeekRows = $allRows | Where-Object { $_.snapshotDate -ge $weekStartStr -and $_.snapshotDate -le $today }
$prevWeekRows = $allRows | Where-Object { $_.snapshotDate -ge $prevWeekStartStr -and $_.snapshotDate -le $prevWeekEndStr }

# Dedupe: keep only latest snapshot per date+email in each week
function Get-LatestPerDateEmail {
  param($rows)
  $grouped = $rows | Group-Object -Property { "$($_.snapshotDate)|$($_.email)" }
  $deduped = @()
  foreach ($g in $grouped) {
    $latest = $g.Group | Sort-Object @{ Expression = { [datetime]$_.snapshotAt }; Descending = $true } | Select-Object -First 1
    $deduped += $latest
  }
  return $deduped
}

$currentDeduped = Get-LatestPerDateEmail $currentWeekRows
$prevDeduped = Get-LatestPerDateEmail $prevWeekRows

# Aggregate per partner
function Get-PartnerWeekStats {
  param($rows, $intFn, $dblFn)
  $grouped = $rows | Group-Object -Property email
  $stats = @{}
  foreach ($g in $grouped) {
    $email = $g.Name
    $partnerRows = $g.Group
    $daysActive = ($partnerRows | Select-Object -ExpandProperty snapshotDate -Unique).Count
    $totalPrompts = ($partnerRows | ForEach-Object { & $intFn $_.prompts24h } | Measure-Object -Sum).Sum
    $totalConfirmed = ($partnerRows | ForEach-Object { & $intFn $_.confirmed } | Measure-Object -Sum).Sum
    $totalProposed = ($partnerRows | ForEach-Object { & $intFn $_.proposed } | Measure-Object -Sum).Sum
    $totalBlocked = ($partnerRows | ForEach-Object { & $intFn $_.blocked } | Measure-Object -Sum).Sum
    $avgChatMs = [math]::Round(($partnerRows | ForEach-Object { & $dblFn $_.medianChatLatencyMs } | Measure-Object -Average).Average, 1)
    $approvalRate = if ($totalProposed -gt 0) { [math]::Round($totalConfirmed / $totalProposed, 2) } else { 0 }

    $stats[$email] = [ordered]@{
      email = $email
      daysActive = $daysActive
      totalPrompts = $totalPrompts
      totalConfirmed = $totalConfirmed
      totalProposed = $totalProposed
      totalBlocked = $totalBlocked
      avgChatMs = $avgChatMs
      approvalRate = $approvalRate
    }
  }
  return $stats
}

$currentStats = Get-PartnerWeekStats $currentDeduped $toInt $toDouble
$prevStats = Get-PartnerWeekStats $prevDeduped $toInt $toDouble

# Build digest per partner
$allEmails = @($currentStats.Keys) + @($prevStats.Keys) | Select-Object -Unique | Sort-Object
$digest = @()

foreach ($email in $allEmails) {
  $cur = if ($currentStats.ContainsKey($email)) { $currentStats[$email] } else { $null }
  $prv = if ($prevStats.ContainsKey($email)) { $prevStats[$email] } else { $null }

  $curPrompts = if ($cur) { $cur.totalPrompts } else { 0 }
  $prvPrompts = if ($prv) { $prv.totalPrompts } else { 0 }
  $curConfirmed = if ($cur) { $cur.totalConfirmed } else { 0 }
  $prvConfirmed = if ($prv) { $prv.totalConfirmed } else { 0 }
  $curBlocked = if ($cur) { $cur.totalBlocked } else { 0 }
  $prvBlocked = if ($prv) { $prv.totalBlocked } else { 0 }
  $curChatMs = if ($cur) { $cur.avgChatMs } else { 0 }
  $prvChatMs = if ($prv) { $prv.avgChatMs } else { 0 }
  $curDays = if ($cur) { $cur.daysActive } else { 0 }
  $prvDays = if ($prv) { $prv.daysActive } else { 0 }
  $curApproval = if ($cur) { $cur.approvalRate } else { 0 }
  $prvApproval = if ($prv) { $prv.approvalRate } else { 0 }

  $digest += [ordered]@{
    email = $email
    currentWeek = [ordered]@{
      range = "$weekStartStr..$today"
      daysActive = $curDays
      totalPrompts = $curPrompts
      totalConfirmed = $curConfirmed
      totalBlocked = $curBlocked
      avgChatMs = $curChatMs
      approvalRate = $curApproval
    }
    priorWeek = [ordered]@{
      range = "$prevWeekStartStr..$prevWeekEndStr"
      daysActive = $prvDays
      totalPrompts = $prvPrompts
      totalConfirmed = $prvConfirmed
      totalBlocked = $prvBlocked
      avgChatMs = $prvChatMs
      approvalRate = $prvApproval
    }
    weekOverWeek = [ordered]@{
      deltaPrompts = $curPrompts - $prvPrompts
      deltaConfirmed = $curConfirmed - $prvConfirmed
      deltaBlocked = $curBlocked - $prvBlocked
      deltaChatMs = [math]::Round($curChatMs - $prvChatMs, 1)
      deltaDaysActive = $curDays - $prvDays
      deltaApprovalRate = [math]::Round($curApproval - $prvApproval, 2)
    }
  }
}

$summary = [ordered]@{
  generatedAt = (Get-Date).ToString("o")
  date = $today
  currentWeekRange = "$weekStartStr..$today"
  priorWeekRange = "$prevWeekStartStr..$prevWeekEndStr"
  partnerCount = $allEmails.Count
  currentWeekDataPoints = $currentDeduped.Count
  priorWeekDataPoints = $prevDeduped.Count
  partners = $digest
}

# ---------- output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host ('  [ok] Digest written: ' + $OutputPath) -ForegroundColor Green
}

Write-Host ''
if ($Json) {
  $summary | ConvertTo-Json -Depth 8
}
else {
  $display = $digest | ForEach-Object {
    $d = $_
    $wow = $d.weekOverWeek
    [pscustomobject]@{
      Email = $d.email
      CurDays = $d.currentWeek.daysActive
      CurPrompts = $d.currentWeek.totalPrompts
      CurConfirmed = $d.currentWeek.totalConfirmed
      CurBlocked = $d.currentWeek.totalBlocked
      CurChatMs = $d.currentWeek.avgChatMs
      CurApproval = $d.currentWeek.approvalRate
      WoWPrompts = if ($wow.deltaPrompts -ge 0) { "+$($wow.deltaPrompts)" } else { "$($wow.deltaPrompts)" }
      WoWConfirmed = if ($wow.deltaConfirmed -ge 0) { "+$($wow.deltaConfirmed)" } else { "$($wow.deltaConfirmed)" }
      WoWBlocked = if ($wow.deltaBlocked -ge 0) { "+$($wow.deltaBlocked)" } else { "$($wow.deltaBlocked)" }
    }
  }

  $display | Format-Table -AutoSize | Out-String | Write-Host
}

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host ('  [warn] Triage file not found: ' + $triagePath) -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $parts = @()
  foreach ($d in $digest) {
    $wow = $d.weekOverWeek
    $dP = if ($wow.deltaPrompts -ge 0) { "+$($wow.deltaPrompts)" } else { "$($wow.deltaPrompts)" }
    $dC = if ($wow.deltaConfirmed -ge 0) { "+$($wow.deltaConfirmed)" } else { "$($wow.deltaConfirmed)" }
    $parts += "$($d.email) prompts=$($d.currentWeek.totalPrompts) (wow=$dP), confirmed=$($d.currentWeek.totalConfirmed) (wow=$dC), days=$($d.currentWeek.daysActive), approval=$($d.currentWeek.approvalRate)"
  }

  $line = "$stamp Weekly digest ($($summary.currentWeekRange)): " + ($parts -join ' | ')

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Weekly digest \(' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted weekly digest in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] Weekly digest complete.' -ForegroundColor Green
