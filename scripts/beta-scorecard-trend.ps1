param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string]$Email = "partner.a@company.com",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [string]$HistoryPath = "docs/reports/scorecard-history.csv",
  [switch]$Json,
  [switch]$AppendToTriage,
  [int]$ShowDays = 7
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

Write-Host '== tsifulator.ai scorecard trend ==' -ForegroundColor Cyan
Write-Host "Date: $today"

# ---------- run scorecard to collect today's data ----------
$scorecardScript = Join-Path $PSScriptRoot "beta-daily-scorecard.ps1"
if (-not (Test-Path $scorecardScript)) {
  throw "Missing required script: $scorecardScript"
}

# Capture scorecard JSON
$scorecardJson = & powershell -ExecutionPolicy Bypass -File $scorecardScript -BaseUrl $BaseUrl -Email $Email -Emails ($Emails -join ',') -Date $today -Json 2>&1 | Where-Object { $_ -notmatch '^\s*==' -and $_ -notmatch '^\s*Date:' -and $_ -notmatch '^\[done\]' }
$scorecardText = ($scorecardJson -join "`n").Trim()
$sc = $null
try {
  $sc = $scorecardText | ConvertFrom-Json
}
catch {
  Write-Host ('  [warn] Could not parse scorecard JSON') -ForegroundColor DarkYellow
  Write-Host $scorecardText
  exit 1
}

# ---------- build today's row ----------
$row = [ordered]@{
  snapshotDate    = $today
  snapshotAt      = (Get-Date).ToString("o")
  grade           = $sc.grade
  score           = $sc.score
  maxScore        = $sc.maxScore
  pct             = $sc.pct
  promptsSent24h  = $sc.kpi.promptsSent24h
  approvalRate    = $sc.kpi.approvalRate
  blockedAttempts = $sc.kpi.blockedAttempts
  streamRate      = $sc.kpi.streamRate
  historyRows     = $sc.history.rows
  historyDays     = $sc.history.days
  feedbackTotal   = $sc.feedback.total
  partnersActive  = $sc.feedback.partnersReporting
  partnersTotal   = $sc.feedback.partnersTotal
  dedupeOk        = $sc.triage.dedupeOk
  apiHealthy      = $sc.api.healthy
}

# ---------- upsert into history CSV ----------
$historyDir = Split-Path -Path $HistoryPath -Parent
if ($historyDir -and -not (Test-Path $historyDir)) {
  New-Item -ItemType Directory -Path $historyDir -Force | Out-Null
}

$existingRows = @()
if (Test-Path $HistoryPath) {
  $existingRows = @(Import-Csv -Path $HistoryPath)
}

# Remove any existing row for today (upsert)
$filteredRows = @($existingRows | Where-Object { $_.snapshotDate -ne $today })
$filteredRows += [PSCustomObject]$row

# Sort by date
$sortedRows = $filteredRows | Sort-Object { $_.snapshotDate }
$sortedRows | Export-Csv -Path $HistoryPath -NoTypeInformation -Encoding UTF8

Write-Host ('  [ok] Upserted scorecard for ' + $today + ' in ' + $HistoryPath) -ForegroundColor Green

# ---------- load history for display ----------
$allRows = @(Import-Csv -Path $HistoryPath | Sort-Object { $_.snapshotDate })
$recentRows = @($allRows | Select-Object -Last $ShowDays)

Write-Host ''
Write-Host "  Trend (last $($recentRows.Count) entries):" -ForegroundColor White
Write-Host ''

if ($Json) {
  $trendData = [ordered]@{
    generatedAt = (Get-Date).ToString("o")
    date        = $today
    showDays    = $ShowDays
    totalEntries = $allRows.Count
    today       = $row
    trend       = @()
  }

  for ($i = 0; $i -lt $recentRows.Count; $i++) {
    $curr = $recentRows[$i]
    $prev = if ($i -gt 0) { $recentRows[$i - 1] } else { $null }

    $entry = [ordered]@{
      date           = $curr.snapshotDate
      grade          = $curr.grade
      score          = [int]$curr.score
      pct            = [int]$curr.pct
      prompts24h     = [int]$curr.promptsSent24h
      approvalRate   = [double]$curr.approvalRate
      feedback       = [int]$curr.feedbackTotal
      historyRows    = [int]$curr.historyRows
    }

    if ($prev) {
      $entry['d_score'] = [int]$curr.score - [int]$prev.score
      $entry['d_pct'] = [int]$curr.pct - [int]$prev.pct
      $entry['d_prompts'] = [int]$curr.promptsSent24h - [int]$prev.promptsSent24h
      $entry['d_approval'] = [math]::Round([double]$curr.approvalRate - [double]$prev.approvalRate, 2)
      $entry['d_feedback'] = [int]$curr.feedbackTotal - [int]$prev.feedbackTotal
    }

    $trendData.trend += $entry
  }

  $trendData | ConvertTo-Json -Depth 8
}
else {
  # Table display
  $tableRows = @()
  for ($i = 0; $i -lt $recentRows.Count; $i++) {
    $curr = $recentRows[$i]
    $prev = if ($i -gt 0) { $recentRows[$i - 1] } else { $null }

    $dScore = ''
    $dPrompts = ''
    $dApproval = ''
    $dFeedback = ''

    if ($prev) {
      $ds = [int]$curr.score - [int]$prev.score
      $dp = [int]$curr.promptsSent24h - [int]$prev.promptsSent24h
      $da = [math]::Round([double]$curr.approvalRate - [double]$prev.approvalRate, 2)
      $df = [int]$curr.feedbackTotal - [int]$prev.feedbackTotal

      $dScore = if ($ds -gt 0) { "+$ds" } elseif ($ds -lt 0) { "$ds" } else { '=' }
      $dPrompts = if ($dp -gt 0) { "+$dp" } elseif ($dp -lt 0) { "$dp" } else { '=' }
      $dApproval = if ($da -gt 0) { "+$da" } elseif ($da -lt 0) { "$da" } else { '=' }
      $dFeedback = if ($df -gt 0) { "+$df" } elseif ($df -lt 0) { "$df" } else { '=' }
    }
    else {
      $dScore = 'n/a'
      $dPrompts = 'n/a'
      $dApproval = 'n/a'
      $dFeedback = 'n/a'
    }

    $tableRows += [PSCustomObject]@{
      Date      = $curr.snapshotDate
      Grade     = $curr.grade
      Score     = "$($curr.score)/$($curr.maxScore)"
      Pct       = "$($curr.pct)%"
      dScore    = $dScore
      Prompts   = $curr.promptsSent24h
      dPrompts  = $dPrompts
      Approval  = $curr.approvalRate
      dApproval = $dApproval
      Feedback  = $curr.feedbackTotal
      dFeedback = $dFeedback
      API       = $curr.apiHealthy
      Dedupe    = $curr.dedupeOk
    }
  }

  $tableRows | Format-Table -AutoSize
}

# ---------- grade change alert ----------
if ($allRows.Count -ge 2) {
  $prevRow = $allRows[$allRows.Count - 2]
  $currRow = $allRows[$allRows.Count - 1]
  if ($currRow.grade -ne $prevRow.grade) {
    $direction = if ([int]$currRow.pct -gt [int]$prevRow.pct) { 'improved' } else { 'regressed' }
    Write-Host "  ** Grade $direction" -NoNewline -ForegroundColor $(if ($direction -eq 'improved') { 'Green' } else { 'Red' })
    Write-Host ": $($prevRow.grade) -> $($currRow.grade)  ($($prevRow.pct)% -> $($currRow.pct)%)" -ForegroundColor $(if ($direction -eq 'improved') { 'Green' } else { 'Red' })
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
  $line = "$stamp Scorecard trend: grade=$($row.grade), score=$($row.score)/$($row.maxScore) ($($row.pct)%), entries=$($allRows.Count), prompts=$($row.promptsSent24h), approval=$($row.approvalRate), feedback=$($row.feedbackTotal)"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Scorecard trend:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted scorecard trend in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] Scorecard trend complete.' -ForegroundColor Green
