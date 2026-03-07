param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [switch]$Json,
  [switch]$AppendToTriage,
  [string]$OutputPath = "",
  [int]$Days = 30
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { [DateTime]::Parse($Date) } else { Get-Date }
$triagePath = "docs/daily-triage/$($today.ToString('yyyy-MM-dd')).md"

# Normalize emails
$normalizedEmails = @()
foreach ($entry in $Emails) {
  $normalizedEmails += ($entry -split '[,;\s]+' | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
}
$Emails = $normalizedEmails | Select-Object -Unique

Write-Host '== tsifulator.ai partner session calendar ==' -ForegroundColor Cyan
Write-Host "Date: $($today.ToString('yyyy-MM-dd'))"
Write-Host "Partners: $($Emails.Count)"
Write-Host "Days: $Days"

$startDate = $today.AddDays(-$Days+1).Date
$endDate = $today.Date

$results = @()

foreach ($email in $Emails) {
  $loginBody = @{ email = $email } | ConvertTo-Json
  try {
    $login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $loginBody
    $headers = @{ authorization = "Bearer $($login.token)" }
    $sessions = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=100" -Headers $headers
    # Build per-day session count
    $sessionCounts = @{}
    foreach ($s in $sessions.sessions) {
      try {
        $dt = [DateTime]::Parse($s.created_at)
        $dateKey = $dt.ToString('yyyy-MM-dd')
        if (-not $sessionCounts.ContainsKey($dateKey)) { $sessionCounts[$dateKey] = 0 }
        $sessionCounts[$dateKey]++
      } catch {}
    }
    # Build calendar string
    $calendar = ''
    $streak = 0
    $maxStreak = 0
    $gap = 0
    $maxGap = 0
    for ($d = 0; $d -lt $Days; $d++) {
      $dateKey = $startDate.AddDays($d).ToString('yyyy-MM-dd')
      $hasSession = $sessionCounts.ContainsKey($dateKey) -and $sessionCounts[$dateKey] -gt 0
      if ($hasSession) {
        $calendar += '█'
        $streak++
        if ($streak -gt $maxStreak) { $maxStreak = $streak }
        $gap = 0
      } else {
        $calendar += '·'
        $gap++
        if ($gap -gt $maxGap) { $maxGap = $gap }
        $streak = 0
      }
    }
    $results += [ordered]@{
      email = $email
      calendar = $calendar
      sessionCounts = $sessionCounts
      maxStreak = $maxStreak
      maxGap = $maxGap
    }
  } catch {
    $results += [ordered]@{ email = $email; error = $_.Exception.Message }
  }
}

$summary = [ordered]@{
  date = $today.ToString('yyyy-MM-dd')
  days = $Days
  partners = $results
}

if ($Json) {
  $summary | ConvertTo-Json -Depth 8
  exit 0
}

if ($OutputPath) {
  $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding utf8
  Write-Host "[saved] $OutputPath"
}

# Display
Write-Host ''
Write-Host '--- Partner Session Calendar ---' -ForegroundColor Cyan
foreach ($p in $results) {
  Write-Host "  $($p.email)" -ForegroundColor White
  if ($p['error'] -ne $null) {
    Write-Host "    Error: $($p.error)" -ForegroundColor Red
    continue
  }
  Write-Host "    $($p.calendar)"
  Write-Host "    Max streak: $($p.maxStreak)  Max gap: $($p.maxGap)"
}

# Triage
if ($AppendToTriage) {
  if (-not (Test-Path (Split-Path $triagePath -Parent))) {
    New-Item -ItemType Directory -Path (Split-Path $triagePath -Parent) -Force | Out-Null
  }
  $triageLine = "Partner session calendar: days=$Days"
  if (Test-Path $triagePath) {
    $existing = Get-Content -Path $triagePath -Raw
    $cleaned = ($existing -split "`n" | Where-Object { $_ -notmatch '^Partner session calendar:' }) -join "`n"
    Set-Content -Path $triagePath -Value $cleaned.TrimEnd() -Encoding utf8
  }
  Add-Content -Path $triagePath -Value $triageLine -Encoding utf8
  Write-Host "[triage] appended to $triagePath" -ForegroundColor Green
}

Write-Host 'done.' -ForegroundColor Green
