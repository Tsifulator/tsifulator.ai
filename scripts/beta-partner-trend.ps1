param(
  [string]$HistoryPath = "docs/reports/partner-compare-history.csv",
  [switch]$Json,
  [switch]$AppendToTriage,
  [string]$Date
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"

Write-Host "== tsifulator.ai partner trend report ==" -ForegroundColor Cyan

if (-not (Test-Path $HistoryPath)) {
  Write-Host "[warn] History file not found: $HistoryPath" -ForegroundColor DarkYellow
  Write-Host "Run: npm run beta:compare:partners:csv:append" -ForegroundColor DarkYellow
  exit 1
}

$rows = Import-Csv -Path $HistoryPath
if (-not $rows -or $rows.Count -eq 0) {
  Write-Host "[warn] History file is empty: $HistoryPath" -ForegroundColor DarkYellow
  exit 1
}

$toInt = {
  param($value)
  if ($null -eq $value -or $value -eq "") { return 0 }
  return [int]$value
}

$grouped = $rows | Group-Object -Property email
$results = @()

foreach ($group in $grouped) {
  $email = $group.Name
  $partnerRows = $group.Group | Sort-Object @{ Expression = { [datetime]$_.snapshotAt }; Descending = $true }
  $latest = $partnerRows[0]
  $previousDateRow = $partnerRows | Where-Object { $_.snapshotDate -ne $latest.snapshotDate } | Select-Object -First 1

  $latestPrompts = & $toInt $latest.prompts24h
  $latestConfirmed = & $toInt $latest.confirmed
  $latestBlocked = & $toInt $latest.blocked
  $latestChatMs = & $toInt $latest.medianChatLatencyMs

  $deltaPrompts = if ($null -ne $previousDateRow) { $latestPrompts - (& $toInt $previousDateRow.prompts24h) } else { $null }
  $deltaConfirmed = if ($null -ne $previousDateRow) { $latestConfirmed - (& $toInt $previousDateRow.confirmed) } else { $null }
  $deltaBlocked = if ($null -ne $previousDateRow) { $latestBlocked - (& $toInt $previousDateRow.blocked) } else { $null }
  $deltaChatMs = if ($null -ne $previousDateRow) { $latestChatMs - (& $toInt $previousDateRow.medianChatLatencyMs) } else { $null }

  $results += [ordered]@{
    email = $email
    latestDate = $latest.snapshotDate
    latestAt = $latest.snapshotAt
    prompts24h = $latestPrompts
    confirmed = $latestConfirmed
    blocked = $latestBlocked
    streamRatio = $latest.streamRatio
    medianChatLatencyMs = $latestChatMs
    deltaPrompts24h = $deltaPrompts
    deltaConfirmed = $deltaConfirmed
    deltaBlocked = $deltaBlocked
    deltaMedianChatLatencyMs = $deltaChatMs
  }
}

$results = $results | Sort-Object email

Write-Host ""
if ($Json) {
  $results | ConvertTo-Json -Depth 6
}
else {
  $display = $results | ForEach-Object {
    [pscustomobject]@{
      Email = $_.email
      Date = $_.latestDate
      Prompts24h = $_.prompts24h
      Confirmed = $_.confirmed
      Blocked = $_.blocked
      Stream = $_.streamRatio
      ChatMs = $_.medianChatLatencyMs
      DeltaPrompts = if ($null -eq $_.deltaPrompts24h) { "n/a" } else { $_.deltaPrompts24h }
      DeltaConfirmed = if ($null -eq $_.deltaConfirmed) { "n/a" } else { $_.deltaConfirmed }
      DeltaBlocked = if ($null -eq $_.deltaBlocked) { "n/a" } else { $_.deltaBlocked }
      DeltaChatMs = if ($null -eq $_.deltaMedianChatLatencyMs) { "n/a" } else { $_.deltaMedianChatLatencyMs }
    }
  }

  $display | Format-Table -AutoSize | Out-String | Write-Host
}

if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host "[warn] Triage file not found: $triagePath" -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $parts = @()
  foreach ($result in $results) {
    $deltaPromptsText = if ($null -eq $result.deltaPrompts24h) { "n/a" } else { [string]$result.deltaPrompts24h }
    $deltaConfirmedText = if ($null -eq $result.deltaConfirmed) { "n/a" } else { [string]$result.deltaConfirmed }
    $deltaBlockedText = if ($null -eq $result.deltaBlocked) { "n/a" } else { [string]$result.deltaBlocked }
    $deltaChatText = if ($null -eq $result.deltaMedianChatLatencyMs) { "n/a" } else { [string]$result.deltaMedianChatLatencyMs }

    $parts += "$($result.email) prompts24h=$($result.prompts24h) (d=$deltaPromptsText), confirmed=$($result.confirmed) (d=$deltaConfirmedText), blocked=$($result.blocked) (d=$deltaBlockedText), stream=$($result.streamRatio), medChatMs=$($result.medianChatLatencyMs) (d=$deltaChatText)"
  }

  $line = "[$stamp] Partner trend snapshot: " + ($parts -join " | ")
  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch "^- \[[^\]]+\] Partner trend snapshot:" }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host "[ok] Upserted partner trend snapshot in triage: $triagePath" -ForegroundColor Green
}

Write-Host ""
Write-Host "[done] Partner trend report complete." -ForegroundColor Green
