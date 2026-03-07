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

Write-Host '== tsifulator.ai session depth analyzer ==' -ForegroundColor Cyan
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

# ---------- depth categories ----------
function Get-DepthLabel {
  param([double]$msgsPerSession)
  if ($msgsPerSession -ge 6)  { return 'deep' }
  if ($msgsPerSession -ge 3)  { return 'moderate' }
  if ($msgsPerSession -ge 1)  { return 'shallow' }
  return 'empty'
}

# ---------- collect per-partner session data ----------
$partnerResults = @()
$globalSessionDetails = @()

foreach ($email in $Emails) {
  $emailClean = $email.Trim()
  if (-not $emailClean) { continue }

  $sessionDetails = @()

  try {
    $loginBody = @{ email = $emailClean } | ConvertTo-Json
    $login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $loginBody
    $headers = @{ authorization = "Bearer $($login.token)" }
  }
  catch {
    Write-Host ('  [warn] Could not login as ' + $emailClean) -ForegroundColor DarkYellow
    $partnerResults += [ordered]@{
      email       = $emailClean
      reachable   = $false
      sessions    = 0
      depthLabel  = 'unknown'
      avgMsgsPerSession     = 0
      avgProposalsPerSession = 0
      avgApprovalsPerSession = 0
      avgFeedbackPerSession  = 0
      avgDurationSec        = 0
      sessionDetails        = @()
    }
    continue
  }

  try {
    $sessionsResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=$SessionLimit" -Headers $headers
    $sessions = if ($sessionsResp.sessions) { $sessionsResp.sessions } else { @() }
  }
  catch { $sessions = @() }

  foreach ($session in $sessions) {
    try {
      $eventsResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($session.id)/events" -Headers $headers
      $events = if ($eventsResp.events) { $eventsResp.events } else { @() }
    }
    catch { $events = @(); continue }

    $userMsgs = 0; $assistantMsgs = 0; $proposals = 0; $approvals = 0
    $blocked = 0; $feedback = 0; $streams = 0; $totalEvents = $events.Count
    $timestamps = @()

    foreach ($evt in $events) {
      if ($evt.createdAt) { $timestamps += $evt.createdAt }
      switch ($evt.type) {
        'chat_user_message'      { $userMsgs++ }
        'chat_assistant_message' { $assistantMsgs++ }
        'action_proposed'        { $proposals++ }
        'action_executed'        { $approvals++ }
        'action_blocked'         { $blocked++ }
        'user_feedback'          { $feedback++ }
        'chat_stream_completed'  { $streams++ }
      }
    }

    # Duration: first event to last event
    $durationSec = 0
    if ($timestamps.Count -ge 2) {
      try {
        $sorted = $timestamps | Sort-Object
        $first = [DateTime]::Parse($sorted[0])
        $last = [DateTime]::Parse($sorted[-1])
        $durationSec = [math]::Max(0, ($last - $first).TotalSeconds)
      } catch {}
    }

    $detail = [ordered]@{
      sessionId     = $session.id
      totalEvents   = $totalEvents
      userMessages  = $userMsgs
      assistantMsgs = $assistantMsgs
      proposals     = $proposals
      approvals     = $approvals
      blocked       = $blocked
      feedback      = $feedback
      streams       = $streams
      durationSec   = [math]::Round($durationSec, 1)
    }
    $sessionDetails += $detail
    $globalSessionDetails += $detail
  }

  $sessionCount = $sessionDetails.Count

  # Compute averages
  $avgMsgs = 0; $avgProposals = 0; $avgApprovals = 0; $avgFeedback = 0; $avgDuration = 0
  if ($sessionCount -gt 0) {
    $avgMsgs = [math]::Round(($sessionDetails | ForEach-Object { $_.userMessages } | Measure-Object -Average).Average, 1)
    $avgProposals = [math]::Round(($sessionDetails | ForEach-Object { $_.proposals } | Measure-Object -Average).Average, 1)
    $avgApprovals = [math]::Round(($sessionDetails | ForEach-Object { $_.approvals } | Measure-Object -Average).Average, 1)
    $avgFeedback = [math]::Round(($sessionDetails | ForEach-Object { $_.feedback } | Measure-Object -Average).Average, 1)
    $avgDuration = [math]::Round(($sessionDetails | ForEach-Object { $_.durationSec } | Measure-Object -Average).Average, 1)
  }

  $depthLabel = Get-DepthLabel -msgsPerSession $avgMsgs

  # Find deepest and shallowest sessions
  $sorted = $sessionDetails | Sort-Object { $_.userMessages } -Descending
  $deepest = if ($sorted.Count -gt 0) { $sorted[0] } else { $null }
  $shallowest = if ($sorted.Count -gt 0) { $sorted[-1] } else { $null }

  $partnerResults += [ordered]@{
    email                   = $emailClean
    reachable               = $true
    sessions                = $sessionCount
    depthLabel              = $depthLabel
    avgMsgsPerSession       = $avgMsgs
    avgProposalsPerSession  = $avgProposals
    avgApprovalsPerSession  = $avgApprovals
    avgFeedbackPerSession   = $avgFeedback
    avgDurationSec          = $avgDuration
    deepestSession          = if ($deepest) { [ordered]@{ id = $deepest.sessionId; msgs = $deepest.userMessages; events = $deepest.totalEvents } } else { $null }
    shallowestSession       = if ($shallowest -and $shallowest.sessionId -ne $deepest.sessionId) { [ordered]@{ id = $shallowest.sessionId; msgs = $shallowest.userMessages; events = $shallowest.totalEvents } } else { $null }
    sessionDetails          = $sessionDetails
  }
}

# ---------- aggregate ----------
$totalSessions = $globalSessionDetails.Count
$globalAvgMsgs = 0; $globalAvgProposals = 0; $globalAvgDuration = 0
if ($totalSessions -gt 0) {
  $globalAvgMsgs = [math]::Round(($globalSessionDetails | ForEach-Object { $_.userMessages } | Measure-Object -Average).Average, 1)
  $globalAvgProposals = [math]::Round(($globalSessionDetails | ForEach-Object { $_.proposals } | Measure-Object -Average).Average, 1)
  $globalAvgDuration = [math]::Round(($globalSessionDetails | ForEach-Object { $_.durationSec } | Measure-Object -Average).Average, 1)
}
$globalDepth = Get-DepthLabel -msgsPerSession $globalAvgMsgs

# Depth distribution
$depthDist = [ordered]@{ deep = 0; moderate = 0; shallow = 0; empty = 0 }
foreach ($sd in $globalSessionDetails) {
  $label = Get-DepthLabel -msgsPerSession $sd.userMessages
  $depthDist[$label]++
}

$report = [ordered]@{
  generatedAt       = (Get-Date).ToString("o")
  date              = $today
  totalSessions     = $totalSessions
  totalPartners     = $Emails.Count
  globalDepth       = $globalDepth
  globalAvgMsgsPerSession     = $globalAvgMsgs
  globalAvgProposalsPerSession = $globalAvgProposals
  globalAvgDurationSec        = $globalAvgDuration
  depthDistribution = $depthDist
  partners          = $partnerResults
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $report | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host ('  [ok] Report written: ' + $OutputPath) -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $report | ConvertTo-Json -Depth 8
}
else {
  $depthColor = switch ($globalDepth) {
    'deep'     { 'Green' }
    'moderate' { 'Yellow' }
    'shallow'  { 'DarkYellow' }
    default    { 'Red' }
  }
  Write-Host "  Global depth: $globalDepth ($totalSessions sessions)" -ForegroundColor $depthColor
  Write-Host "  Avg per session: msgs=$globalAvgMsgs  proposals=$globalAvgProposals  duration=$($globalAvgDuration)s" -ForegroundColor White
  Write-Host "  Distribution: deep=$($depthDist.deep) moderate=$($depthDist.moderate) shallow=$($depthDist.shallow) empty=$($depthDist.empty)" -ForegroundColor Gray
  Write-Host ''

  # Per-partner table
  Write-Host '  Per partner:' -ForegroundColor White
  $headerLine = '  {0,-25} {1,4} {2,4} {3,4} {4,4} {5,3} {6,6} {7,8}' -f 'EMAIL','SESS','MSGS','PROP','APPR','FB','DUR(s)','DEPTH'
  Write-Host $headerLine -ForegroundColor DarkGray

  foreach ($pr in $partnerResults) {
    if (-not $pr.reachable) {
      Write-Host "  $($pr.email): [unreachable]" -ForegroundColor Red
      continue
    }

    $dColor = switch ($pr.depthLabel) {
      'deep'     { 'Green' }
      'moderate' { 'Yellow' }
      'shallow'  { 'DarkYellow' }
      default    { 'Red' }
    }

    $shortEmail = $pr.email
    if ($shortEmail.Length -gt 25) { $shortEmail = $shortEmail.Substring(0,22) + '...' }
    $line = '  {0,-25} {1,4} {2,4} {3,4} {4,4} {5,3} {6,6} {7,8}' -f $shortEmail, $pr.sessions, $pr.avgMsgsPerSession, $pr.avgProposalsPerSession, $pr.avgApprovalsPerSession, $pr.avgFeedbackPerSession, $pr.avgDurationSec, $pr.depthLabel
    Write-Host $line -ForegroundColor $dColor

    # Session breakdown mini-chart
    if ($pr.sessionDetails.Count -gt 0) {
      $maxMsgs = ($pr.sessionDetails | ForEach-Object { $_.userMessages } | Measure-Object -Maximum).Maximum
      if ($maxMsgs -eq 0) { $maxMsgs = 1 }
      $barMax = 20

      foreach ($sd in $pr.sessionDetails) {
        $barLen = [math]::Floor($sd.userMessages / $maxMsgs * $barMax)
        $bar = '#' * $barLen
        $shortId = $sd.sessionId.Substring(0, [math]::Min(8, $sd.sessionId.Length))
        $sColor = switch (Get-DepthLabel -msgsPerSession $sd.userMessages) {
          'deep'     { 'Green' }
          'moderate' { 'Yellow' }
          'shallow'  { 'DarkYellow' }
          default    { 'DarkGray' }
        }
        Write-Host "    $shortId |$bar $($sd.userMessages)m $($sd.proposals)p $($sd.approvals)a $($sd.durationSec)s" -ForegroundColor $sColor
      }
    }
    Write-Host ''
  }

  # Insights
  $shallowPct = if ($totalSessions -gt 0) { [math]::Round(($depthDist.shallow + $depthDist.empty) / $totalSessions * 100) } else { 0 }
  $deepPct = if ($totalSessions -gt 0) { [math]::Round($depthDist.deep / $totalSessions * 100) } else { 0 }

  Write-Host '  Insights:' -ForegroundColor White
  if ($deepPct -ge 50) {
    Write-Host '    + Majority of sessions are deep (6+ messages) - strong engagement quality' -ForegroundColor Green
  }
  elseif ($shallowPct -ge 50) {
    Write-Host '    ! Majority of sessions are shallow (< 3 messages) - investigate barriers' -ForegroundColor DarkYellow
  }
  else {
    Write-Host '    ~ Mixed session depth - engagement quality is developing' -ForegroundColor Yellow
  }

  $noFeedbackSessions = @($globalSessionDetails | Where-Object { $_.feedback -eq 0 }).Count
  if ($noFeedbackSessions -gt 0 -and $totalSessions -gt 0) {
    $noFbPct = [math]::Round($noFeedbackSessions / $totalSessions * 100)
    if ($noFbPct -gt 50) {
      Write-Host "    ! $noFbPct% of sessions have no feedback - consider adding prompts" -ForegroundColor DarkYellow
    }
  }

  $noProposalSessions = @($globalSessionDetails | Where-Object { $_.proposals -eq 0 }).Count
  if ($noProposalSessions -gt 0 -and $totalSessions -gt 0) {
    $noPropPct = [math]::Round($noProposalSessions / $totalSessions * 100)
    if ($noPropPct -gt 30) {
      Write-Host "    ~ $noPropPct% of sessions had no proposals - users may need guidance" -ForegroundColor Yellow
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
  $perPartnerStr = ($partnerResults | ForEach-Object {
    "$($_.email)=$($_.depthLabel)($($_.avgMsgsPerSession)m)"
  }) -join ', '

  $line = "$stamp Session depth: globalDepth=$globalDepth, sessions=$totalSessions, avgMsgs=$globalAvgMsgs, avgProposals=$globalAvgProposals, avgDur=$($globalAvgDuration)s, partners=($perPartnerStr)"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Session depth:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted session depth in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] Session depth analysis complete.' -ForegroundColor Green
