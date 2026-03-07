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

Write-Host '== tsifulator.ai partner onboarding ==' -ForegroundColor Cyan
Write-Host "Date:     $today"
Write-Host "Partners: $($Emails -join ', ')"

# ---------- milestone definitions ----------
# Each milestone: name, check function (receives $data hashtable), weight
$milestones = @(
  @{ name = 'logged_in';      label = 'Logged in';          weight = 10; check = { param($d) $d.sessionCount -gt 0 } }
  @{ name = 'first_chat';     label = 'First chat';         weight = 15; check = { param($d) $d.chatMessages -gt 0 } }
  @{ name = 'first_proposal'; label = 'First proposal';     weight = 15; check = { param($d) $d.actionsProposed -gt 0 } }
  @{ name = 'first_approval'; label = 'First approval';     weight = 15; check = { param($d) $d.actionsExecuted -gt 0 } }
  @{ name = 'safety_tested';  label = 'Safety tested';      weight = 10; check = { param($d) $d.actionsBlocked -gt 0 } }
  @{ name = 'first_stream';   label = 'First stream';       weight = 10; check = { param($d) $d.streamsCompleted -gt 0 } }
  @{ name = 'first_feedback';  label = 'First feedback';    weight = 15; check = { param($d) $d.feedbackCount -gt 0 } }
  @{ name = 'multi_session';  label = 'Multi-session (3+)'; weight = 10; check = { param($d) $d.sessionCount -ge 3 } }
)

$maxWeight = ($milestones | ForEach-Object { $_.weight } | Measure-Object -Sum).Sum

# ---------- health check ----------
try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health" -TimeoutSec 5
  Write-Host ('[ok] API healthy: ' + $health.status) -ForegroundColor Green
}
catch {
  Write-Host '[error] API unreachable' -ForegroundColor Red
  exit 1
}

# ---------- collect per-partner data ----------
$partnerResults = @()

foreach ($email in $Emails) {
  $emailClean = $email.Trim()
  if (-not $emailClean) { continue }

  $data = @{
    sessionCount     = 0
    chatMessages     = 0
    actionsProposed  = 0
    actionsExecuted  = 0
    actionsBlocked   = 0
    streamsCompleted = 0
    feedbackCount    = 0
    firstEventAt     = $null
    latestEventAt    = $null
  }

  try {
    $loginBody = @{ email = $emailClean } | ConvertTo-Json
    $login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $loginBody
    $headers = @{ authorization = "Bearer $($login.token)" }
  }
  catch {
    Write-Host ('  [warn] Could not login as ' + $emailClean) -ForegroundColor DarkYellow
    $partnerResults += [ordered]@{
      email      = $emailClean
      reachable  = $false
      data       = $data
      milestones = @()
      completed  = 0
      total      = $milestones.Count
      pct        = 0
      score      = 0
      maxScore   = $maxWeight
    }
    continue
  }

  try {
    $sessionsResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=$SessionLimit" -Headers $headers
    $sessions = $sessionsResp.sessions
    $data.sessionCount = if ($sessions) { $sessions.Count } else { 0 }
  }
  catch {
    $sessions = @()
  }

  # Collect all events
  $allTimestamps = @()
  foreach ($session in $sessions) {
    try {
      $eventsResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($session.id)/events" -Headers $headers
      $events = $eventsResp.events
    }
    catch { continue }

    if (-not $events) { continue }

    foreach ($evt in $events) {
      if ($evt.createdAt) { $allTimestamps += $evt.createdAt }

      switch ($evt.type) {
        'chat_user_message'      { $data.chatMessages++ }
        'action_proposed'        { $data.actionsProposed++ }
        'action_executed'        { $data.actionsExecuted++ }
        'action_blocked'         { $data.actionsBlocked++ }
        'chat_stream_completed'  { $data.streamsCompleted++ }
        'user_feedback'          { $data.feedbackCount++ }
      }
    }
  }

  if ($allTimestamps.Count -gt 0) {
    $sorted = $allTimestamps | Sort-Object
    $data.firstEventAt = $sorted[0]
    $data.latestEventAt = $sorted[-1]
  }

  # Evaluate milestones
  $milestoneResults = @()
  $completedCount = 0
  $earnedScore = 0

  foreach ($ms in $milestones) {
    $passed = & $ms.check $data
    if ($passed) {
      $completedCount++
      $earnedScore += $ms.weight
    }
    $milestoneResults += [ordered]@{
      name     = $ms.name
      label    = $ms.label
      weight   = $ms.weight
      passed   = [bool]$passed
    }
  }

  $pctComplete = if ($milestones.Count -gt 0) { [math]::Round($completedCount / $milestones.Count * 100) } else { 0 }

  $partnerResults += [ordered]@{
    email      = $emailClean
    reachable  = $true
    data       = [ordered]@{
      sessions         = $data.sessionCount
      chatMessages     = $data.chatMessages
      actionsProposed  = $data.actionsProposed
      actionsExecuted  = $data.actionsExecuted
      actionsBlocked   = $data.actionsBlocked
      streamsCompleted = $data.streamsCompleted
      feedbackCount    = $data.feedbackCount
      firstEventAt     = $data.firstEventAt
      latestEventAt    = $data.latestEventAt
    }
    milestones = $milestoneResults
    completed  = $completedCount
    total      = $milestones.Count
    pct        = $pctComplete
    score      = $earnedScore
    maxScore   = $maxWeight
  }
}

# ---------- aggregate ----------
$totalPartners = $partnerResults.Count
$fullyOnboarded = @($partnerResults | Where-Object { $_.pct -eq 100 }).Count
$avgPct = if ($totalPartners -gt 0) {
  [math]::Round(($partnerResults | ForEach-Object { $_.pct } | Measure-Object -Average).Average)
} else { 0 }

$overallStatus = if ($fullyOnboarded -eq $totalPartners -and $totalPartners -gt 0) { 'all_onboarded' }
  elseif ($fullyOnboarded -gt 0) { 'partial' }
  else { 'none_onboarded' }

# Per-milestone cross-partner completion
$milestoneOverview = @()
foreach ($ms in $milestones) {
  $passedCount = @($partnerResults | Where-Object {
    ($_.milestones | Where-Object { $_.name -eq $ms.name -and $_.passed }).Count -gt 0
  }).Count
  $milestoneOverview += [ordered]@{
    name        = $ms.name
    label       = $ms.label
    partnersCompleted = $passedCount
    partnersTotal     = $totalPartners
    rate        = if ($totalPartners -gt 0) { [math]::Round($passedCount / $totalPartners, 2) } else { 0 }
  }
}

$report = [ordered]@{
  generatedAt      = (Get-Date).ToString("o")
  date             = $today
  totalPartners    = $totalPartners
  fullyOnboarded   = $fullyOnboarded
  avgCompletion    = $avgPct
  status           = $overallStatus
  milestoneOverview = $milestoneOverview
  partners         = $partnerResults
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $report | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host ('  [ok] Onboarding report written: ' + $OutputPath) -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $report | ConvertTo-Json -Depth 8
}
else {
  Write-Host "  Partners: $totalPartners   Fully onboarded: $fullyOnboarded   Avg completion: $avgPct%" -ForegroundColor White
  Write-Host "  Status: $overallStatus" -ForegroundColor $(if ($overallStatus -eq 'all_onboarded') { 'Green' } elseif ($overallStatus -eq 'partial') { 'Yellow' } else { 'Red' })
  Write-Host ''

  # Milestone overview
  Write-Host '  Milestone adoption:' -ForegroundColor White
  foreach ($mo in $milestoneOverview) {
    $icon = if ($mo.partnersCompleted -eq $totalPartners) { '+' } elseif ($mo.partnersCompleted -gt 0) { '~' } else { '-' }
    $color = if ($mo.partnersCompleted -eq $totalPartners) { 'Green' } elseif ($mo.partnersCompleted -gt 0) { 'Yellow' } else { 'DarkGray' }
    Write-Host "    $icon $($mo.label): $($mo.partnersCompleted)/$totalPartners" -ForegroundColor $color
  }
  Write-Host ''

  # Per-partner detail
  Write-Host '  Per partner:' -ForegroundColor White
  foreach ($pr in $partnerResults) {
    $pColor = if ($pr.pct -eq 100) { 'Green' } elseif ($pr.pct -ge 50) { 'Yellow' } else { 'Red' }
    Write-Host "    $($pr.email): $($pr.completed)/$($pr.total) milestones ($($pr.pct)%)" -ForegroundColor $pColor

    if (-not $pr.reachable) {
      Write-Host '      [unreachable]' -ForegroundColor Red
      continue
    }

    # Show the progress bar
    $barLen = 20
    $filled = [math]::Floor($pr.pct / 100 * $barLen)
    $empty = $barLen - $filled
    $bar = ('#' * $filled) + ('.' * $empty)
    Write-Host "      [$bar]" -ForegroundColor $pColor

    # List individual milestones
    foreach ($ms in $pr.milestones) {
      $msIcon = if ($ms.passed) { '+' } else { ' ' }
      $msColor = if ($ms.passed) { 'Green' } else { 'DarkGray' }
      Write-Host "        $msIcon $($ms.label)" -ForegroundColor $msColor
    }

    # Activity window
    if ($pr.data.firstEventAt) {
      $first = $pr.data.firstEventAt
      $latest = $pr.data.latestEventAt
      if ($first.Length -gt 19) { $first = $first.Substring(0, 19) }
      if ($latest.Length -gt 19) { $latest = $latest.Substring(0, 19) }
      Write-Host "      Activity: $first .. $latest" -ForegroundColor Gray
    }
    Write-Host ''
  }

  # Blockers: milestones with 0% adoption
  $blockers = @($milestoneOverview | Where-Object { $_.partnersCompleted -eq 0 })
  if ($blockers.Count -gt 0) {
    Write-Host '  Blockers (0% adoption):' -ForegroundColor Red
    foreach ($b in $blockers) {
      Write-Host "    ! $($b.label)" -ForegroundColor Red
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
  $perPartnerStr = ($partnerResults | ForEach-Object {
    "$($_.email)=$($_.completed)/$($_.total)"
  }) -join ', '

  $line = "$stamp Partner onboarding: status=$overallStatus, fullyOnboarded=$fullyOnboarded/$totalPartners, avg=$avgPct%, partners=($perPartnerStr)"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Partner onboarding:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted partner onboarding in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] Partner onboarding complete.' -ForegroundColor Green
