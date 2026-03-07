param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [switch]$Json,
  [switch]$Csv,
  [string]$CsvPath = "",
  [switch]$AppendCsvDaily,
  [switch]$AllowDuplicateDaily,
  [string]$RollingCsvPath = "docs/reports/partner-compare-history.csv",
  [switch]$AppendToTriage,
  [string]$Date
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$snapshotAt = (Get-Date).ToString("o")
$triagePath = "docs/daily-triage/$today.md"

Write-Host "== tsifulator.ai partner comparison ==" -ForegroundColor Cyan

try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health"
  Write-Host "[ok] API healthy: $($health.status)" -ForegroundColor Green
}
catch {
  Write-Host "[warn] API not reachable at $BaseUrl. Start it with: npm run dev" -ForegroundColor DarkYellow
  exit 1
}

$normalizedEmails = @()
foreach ($entry in $Emails) {
  $normalizedEmails += ($entry -split '[,;\s]+' | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
}

$normalizedEmails = $normalizedEmails | Select-Object -Unique
if (-not $normalizedEmails -or $normalizedEmails.Count -eq 0) {
  $normalizedEmails = @("partner.a@company.com", "partner.b@company.com")
}

$rows = @()

foreach ($email in $normalizedEmails) {
  $loginBody = @{ email = $email } | ConvertTo-Json
  $login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType "application/json" -Body $loginBody
  $headers = @{ authorization = "Bearer $($login.token)" }

  $kpi = Invoke-RestMethod -Method Get -Uri "$BaseUrl/telemetry/counters" -Headers $headers
  $c = $kpi.counters

  $rows += [ordered]@{
    email = $email
    prompts24h = $c.promptsSent24h
    proposed = $c.applyActionsProposed
    confirmed = $c.applyActionsConfirmed
    blocked = $c.blockedCommandAttempts
    streamRequests = $c.streamRequests
    streamCompletions = $c.streamCompletions
    streamRatio = "$($c.streamCompletions)/$($c.streamRequests)"
    medianChatLatencyMs = $c.medianChatLatencyMs
    medianStreamFirstTokenLatencyMs = $c.medianStreamFirstTokenLatencyMs
  }
}

Write-Host ""
if ($Json) {
  $rows | ConvertTo-Json -Depth 6
}
elseif ($Csv) {
  $targetPath = if ($CsvPath) {
    $CsvPath
  }
  else {
    "docs/reports/partner-compare-$today.csv"
  }

  $targetDir = Split-Path -Path $targetPath -Parent
  if ($targetDir -and -not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
  }

  $rows | ForEach-Object { [pscustomobject]$_ } | Export-Csv -Path $targetPath -NoTypeInformation -Encoding UTF8
  Write-Host "[ok] CSV exported: $targetPath" -ForegroundColor Green
}
elseif ($AppendCsvDaily) {
  $rollingPath = if ($RollingCsvPath) {
    $RollingCsvPath
  }
  else {
    "docs/reports/partner-compare-history.csv"
  }

  $rollingDir = Split-Path -Path $rollingPath -Parent
  if ($rollingDir -and -not (Test-Path $rollingDir)) {
    New-Item -ItemType Directory -Path $rollingDir -Force | Out-Null
  }

  $dailyRows = $rows | ForEach-Object {
    [pscustomobject]@{
      snapshotDate = $today
      snapshotAt = $snapshotAt
      email = $_.email
      prompts24h = $_.prompts24h
      proposed = $_.proposed
      confirmed = $_.confirmed
      blocked = $_.blocked
      streamRequests = $_.streamRequests
      streamCompletions = $_.streamCompletions
      streamRatio = $_.streamRatio
      medianChatLatencyMs = $_.medianChatLatencyMs
      medianStreamFirstTokenLatencyMs = $_.medianStreamFirstTokenLatencyMs
    }
  }

  if ($AllowDuplicateDaily) {
    if (Test-Path $rollingPath) {
      $dailyRows | Export-Csv -Path $rollingPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
      $dailyRows | Export-Csv -Path $rollingPath -NoTypeInformation -Encoding UTF8
    }

    Write-Host "[ok] Daily rows appended (duplicates allowed) in rolling CSV: $rollingPath" -ForegroundColor Green
  }
  else {
    $mergedRows = @()

    if (Test-Path $rollingPath) {
      $existingRows = Import-Csv -Path $rollingPath
      $existingRowsFiltered = $existingRows | Where-Object {
        -not ($_.snapshotDate -eq $today -and $Emails -contains $_.email)
      }

      $mergedRows += $existingRowsFiltered
    }

    $mergedRows += $dailyRows
    $mergedRows | Export-Csv -Path $rollingPath -NoTypeInformation -Encoding UTF8

    Write-Host "[ok] Daily rows upserted in rolling CSV: $rollingPath" -ForegroundColor Green
  }
}
else {
  $displayRows = $rows | ForEach-Object {
    [pscustomobject]@{
      Email = $_.email
      Prompts24h = $_.prompts24h
      Proposed = $_.proposed
      Confirmed = $_.confirmed
      Blocked = $_.blocked
      Stream = $_.streamRatio
      ChatMs = $_.medianChatLatencyMs
      StreamFirstMs = $_.medianStreamFirstTokenLatencyMs
    }
  }

  $displayRows | Format-Table -AutoSize | Out-String | Write-Host
}

if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host "[warn] Triage file not found: $triagePath" -ForegroundColor DarkYellow
    exit 1
  }

  if ($rows.Count -lt 2) {
    Write-Host "[warn] Need at least two partner emails to append comparison." -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $parts = @()
  foreach ($row in $rows) {
    $parts += "$($row.email) prompts24h=$($row.prompts24h), proposed=$($row.proposed), confirmed=$($row.confirmed), blocked=$($row.blocked), stream=$($row.streamCompletions)/$($row.streamRequests), medChatMs=$($row.medianChatLatencyMs), medStreamFirstMs=$($row.medianStreamFirstTokenLatencyMs)"
  }

  $line = "[$stamp] Partner comparison snapshot: " + ($parts -join " | ")
  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch "Partner comparison snapshot:" }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host "[ok] Upserted partner comparison snapshot in triage: $triagePath" -ForegroundColor Green
}

Write-Host ""
Write-Host "[done] Partner comparison complete." -ForegroundColor Green
