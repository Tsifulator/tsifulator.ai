param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [switch]$Json,
  [switch]$AppendToTriage,
  [string]$OutputPath = "",
  [int]$Days = 7
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

Write-Host '== tsifulator.ai partner activity timeline ==' -ForegroundColor Cyan
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
    $allEvents = @()
    foreach ($s in $sessions.sessions) {
      $ev = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($s.id)/events" -Headers $headers
      $allEvents += $ev.events | Where-Object { $_.createdAt }
    }
    # Filter events to last $Days
    $recentEvents = $allEvents | Where-Object {
      try { $dt = [DateTime]::Parse($_.createdAt); $dt -ge $startDate -and $dt -le $endDate } catch { $false }
    }
    # Build per-day, per-hour matrix
    $timeline = @{}
    for ($d = 0; $d -lt $Days; $d++) {
      $dateKey = $startDate.AddDays($d).ToString('yyyy-MM-dd')
      $timeline[$dateKey] = @{}
      for ($h = 0; $h -lt 24; $h++) { $timeline[$dateKey]["$h"] = 0 }
    }
    foreach ($ev in $recentEvents) {
      try {
        $dt = [DateTime]::Parse($ev.createdAt)
        $dateKey = $dt.ToString('yyyy-MM-dd')
        $hour = $dt.Hour.ToString()
        if ($timeline.ContainsKey($dateKey)) {
          $timeline[$dateKey][$hour]++
        }
      } catch {}
    }
    # Daily session count
    $sessionCounts = @{}
    foreach ($s in $sessions.sessions) {
      try {
        $dt = [DateTime]::Parse($s.created_at)
        $dateKey = $dt.ToString('yyyy-MM-dd')
        if (-not $sessionCounts.ContainsKey($dateKey)) { $sessionCounts[$dateKey] = 0 }
        $sessionCounts[$dateKey]++
      } catch {}
    }
    $results += [ordered]@{
      email = $email
      timeline = $timeline
      sessionCounts = $sessionCounts
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
Write-Host '--- Partner Activity Timeline ---' -ForegroundColor Cyan
foreach ($p in $results) {
  Write-Host "  $($p.email)" -ForegroundColor White
  if ($p['error'] -ne $null) {
    Write-Host "    Error: $($p.error)" -ForegroundColor Red
    continue
  }
  foreach ($dateKey in ($p.timeline.Keys | Sort-Object)) {
    $line = "$($dateKey): "
    $activeHours = 0
    for ($h = 0; $h -lt 24; $h++) {
      $cnt = $p.timeline[$dateKey]["$h"]
      if ($cnt -gt 0) { $line += '█'; $activeHours++ } else { $line += '·' }
    }
    $sess = if ($p.sessionCounts.ContainsKey($dateKey)) { $p.sessionCounts[$dateKey] } else { 0 }
    $line += "  ($activeHours h, $sess sessions)"
    Write-Host "    $line"
  }
  Write-Host ''
}

# Triage
if ($AppendToTriage) {
  if (-not (Test-Path (Split-Path $triagePath -Parent))) {
    New-Item -ItemType Directory -Path (Split-Path $triagePath -Parent) -Force | Out-Null
  }
  $triageLine = "Partner activity timeline: days=$Days"
  if (Test-Path $triagePath) {
    $existing = Get-Content -Path $triagePath -Raw
    $cleaned = ($existing -split "`n" | Where-Object { $_ -notmatch '^Partner activity timeline:' }) -join "`n"
    Set-Content -Path $triagePath -Value $cleaned.TrimEnd() -Encoding utf8
  }
  Add-Content -Path $triagePath -Value $triageLine -Encoding utf8
  Write-Host "[triage] appended to $triagePath" -ForegroundColor Green
}

Write-Host 'done.' -ForegroundColor Green
