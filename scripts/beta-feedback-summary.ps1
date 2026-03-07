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

# Normalize emails: split comma/space-delimited entries
$normalizedEmails = @()
foreach ($entry in $Emails) {
  $normalizedEmails += ($entry -split '[,;\s]+' | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
}
$Emails = $normalizedEmails | Select-Object -Unique

Write-Host '== tsifulator.ai feedback summary ==' -ForegroundColor Cyan
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

# ---------- collect feedback per partner ----------
$allPartnerFeedback = @()

foreach ($email in $Emails) {
  $emailClean = $email.Trim()
  if (-not $emailClean) { continue }

  try {
    $loginBody = @{ email = $emailClean } | ConvertTo-Json
    $login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $loginBody
    $headers = @{ authorization = "Bearer $($login.token)" }
  }
  catch {
    Write-Host ('  [warn] Could not login as ' + $emailClean) -ForegroundColor DarkYellow
    continue
  }

  # get sessions
  try {
    $sessionsResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=$SessionLimit" -Headers $headers
    $sessions = $sessionsResp.sessions
  }
  catch {
    Write-Host ('  [warn] Could not fetch sessions for ' + $emailClean) -ForegroundColor DarkYellow
    continue
  }

  if (-not $sessions -or $sessions.Count -eq 0) {
    continue
  }

  $feedbackItems = @()

  foreach ($session in $sessions) {
    try {
      $eventsResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($session.id)/events" -Headers $headers
      $events = $eventsResp.events
    }
    catch {
      continue
    }

    if (-not $events) { continue }

    foreach ($evt in $events) {
      if ($evt.type -ne 'user_feedback') { continue }

      $payload = $null
      try {
        $payload = $evt.payload | ConvertFrom-Json -ErrorAction SilentlyContinue
      }
      catch {
        try {
          # payload may already be an object
          $payload = $evt.payload
        }
        catch { continue }
      }

      if (-not $payload) { continue }

      $feedbackText = ''
      if ($payload -is [string]) {
        $feedbackText = $payload
      }
      elseif ($null -ne $payload.text) {
        $feedbackText = [string]$payload.text
      }

      if (-not $feedbackText) { continue }

      $feedbackItems += [ordered]@{
        sessionId = $session.id
        createdAt = $evt.createdAt
        text      = $feedbackText
      }
    }
  }

  $allPartnerFeedback += [ordered]@{
    email         = $emailClean
    sessionCount  = $sessions.Count
    feedbackCount = $feedbackItems.Count
    feedback      = $feedbackItems
    latestFeedback = if ($feedbackItems.Count -gt 0) {
      ($feedbackItems | Sort-Object { $_.createdAt } | Select-Object -Last 1).text
    } else { $null }
  }
}

# ---------- aggregate ----------
$totalFeedback = ($allPartnerFeedback | ForEach-Object { $_.feedbackCount } | Measure-Object -Sum).Sum
$totalSessions = ($allPartnerFeedback | ForEach-Object { $_.sessionCount } | Measure-Object -Sum).Sum
$partnersWithFeedback = ($allPartnerFeedback | Where-Object { $_.feedbackCount -gt 0 }).Count

$summary = [ordered]@{
  generatedAt          = (Get-Date).ToString("o")
  date                 = $today
  partnerCount         = $Emails.Count
  partnersWithFeedback = $partnersWithFeedback
  totalSessions        = $totalSessions
  totalFeedbackItems   = $totalFeedback
  partners             = $allPartnerFeedback
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host ('  [ok] Feedback summary written: ' + $OutputPath) -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $summary | ConvertTo-Json -Depth 8
}
else {
  Write-Host "Partners: $($Emails.Count), with feedback: $partnersWithFeedback, sessions: $totalSessions, feedback items: $totalFeedback"
  Write-Host ''

  foreach ($p in $allPartnerFeedback) {
    $marker = if ($p.feedbackCount -gt 0) { '+' } else { '-' }
    Write-Host "  $marker $($p.email): $($p.feedbackCount) feedback item(s), $($p.sessionCount) session(s)" -ForegroundColor $(if ($p.feedbackCount -gt 0) { 'Green' } else { 'DarkGray' })

    if ($p.feedbackCount -gt 0) {
      foreach ($fb in $p.feedback) {
        $shortDate = $fb.createdAt
        if ($shortDate.Length -gt 19) { $shortDate = $shortDate.Substring(0, 19) }
        Write-Host "      $shortDate  $($fb.text)" -ForegroundColor Gray
      }
    }
  }
  Write-Host ''
}

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host ('  [warn] Triage file not found: ' + $triagePath) -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $perPartner = ($allPartnerFeedback | ForEach-Object {
    "$($_.email)=$($_.feedbackCount)"
  }) -join ', '

  $latestText = ''
  $latestAll = $allPartnerFeedback | Where-Object { $_.feedbackCount -gt 0 } | ForEach-Object {
    $_.feedback | Sort-Object { $_.createdAt } | Select-Object -Last 1
  }
  if ($latestAll) {
    $newest = $latestAll | Sort-Object { $_.createdAt } | Select-Object -Last 1
    $latestText = if ($newest.text.Length -gt 80) { $newest.text.Substring(0, 77) + '...' } else { $newest.text }
  }

  $latestSuffix = if ($latestText) { ", latest=""$latestText""" } else { '' }
  $line = "$stamp Feedback summary: total=$totalFeedback, partners=$partnersWithFeedback/$($Emails.Count), items=($perPartner)$latestSuffix"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Feedback summary:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted feedback summary in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] Feedback summary complete.' -ForegroundColor Green
