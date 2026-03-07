param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string]$Email = "partner.a@company.com",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [switch]$Json,
  [switch]$AppendToTriage,
  [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"
$historyPath = "docs/reports/partner-compare-history.csv"

# Normalize emails
$normalizedEmails = @()
foreach ($entry in $Emails) {
  $normalizedEmails += ($entry -split '[,;\s]+' | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
}
$Emails = $normalizedEmails | Select-Object -Unique

Write-Host '== tsifulator.ai daily scorecard ==' -ForegroundColor Cyan
Write-Host "Date: $today"

# ---------- collect signals ----------

# 1. API health
$apiOk = $false
$kpi = $null
try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health" -TimeoutSec 5
  $apiOk = ($health.status -eq 'ok')
  $loginBody = @{ email = $Email } | ConvertTo-Json
  $login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $loginBody
  $headers = @{ authorization = "Bearer $($login.token)" }
  $kpiResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/telemetry/counters" -Headers $headers
  $kpi = $kpiResp.counters
}
catch {
  Write-Host '  [warn] API unreachable' -ForegroundColor DarkYellow
}

# 2. Triage dedupe
$triageExists = Test-Path $triagePath
$snapshotCounts = [ordered]@{}
$dedupeOk = $true
if ($triageExists) {
  $patterns = [ordered]@{
    kpi        = '^- \[[^\]]+\] KPI refresh:'
    comparison = 'Partner comparison snapshot:'
    trend      = 'Partner trend snapshot:'
    status     = 'Daily status report:'
    alerts     = 'Daily alerts:'
    feedback   = 'Feedback summary:'
    weekly     = 'Weekly digest'
  }
  foreach ($key in $patterns.Keys) {
    $count = (Select-String -Path $triagePath -Pattern $patterns[$key] -AllMatches).Matches.Count
    $snapshotCounts[$key] = $count
    if ($count -gt 1) { $dedupeOk = $false }
  }
}

# 3. Rolling history
$historyExists = Test-Path $historyPath
$historyRows = 0
$partnerDays = 0
if ($historyExists) {
  $rows = Import-Csv -Path $historyPath
  $historyRows = if ($rows) { $rows.Count } else { 0 }
  $partnerDays = @($rows | Select-Object -ExpandProperty snapshotDate -Unique).Count
}

# 4. Feedback counts + 5. Onboarding milestones (from API)
$totalFeedback = 0
$partnersWithFeedback = 0
$onboardingMilestoneTotal = 8
$onboardingPcts = @()
$fullyOnboarded = 0
$activeHours = @{}
$totalEventCount = 0
foreach ($em in $Emails) {
  try {
    $lb = @{ email = $em } | ConvertTo-Json
    $lg = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $lb
    $hd = @{ authorization = "Bearer $($lg.token)" }
    $sess = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=100" -Headers $hd
    $partnerFb = 0
    $partnerSessions = if ($sess.sessions) { $sess.sessions.Count } else { 0 }
    $hasChatMsg = $false; $hasProposal = $false; $hasApproval = $false
    $hasBlocked = $false; $hasStream = $false; $hasFeedback = $false
    foreach ($s in $sess.sessions) {
      try {
        $ev = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($s.id)/events" -Headers $hd
        $partnerFb += @($ev.events | Where-Object { $_.type -eq 'user_feedback' }).Count
        foreach ($e in $ev.events) {
          $totalEventCount++
          if ($e.createdAt) {
            try { $activeHours[([DateTime]::Parse($e.createdAt)).Hour] = $true } catch {}
          }
          switch ($e.type) {
            'chat_user_message'     { $hasChatMsg = $true }
            'action_proposed'       { $hasProposal = $true }
            'action_executed'       { $hasApproval = $true }
            'action_blocked'        { $hasBlocked = $true }
            'chat_stream_completed' { $hasStream = $true }
            'user_feedback'         { $hasFeedback = $true }
          }
        }
      } catch {}
    }
    $totalFeedback += $partnerFb
    if ($partnerFb -gt 0) { $partnersWithFeedback++ }
    # Count milestones hit
    $msHit = 0
    if ($partnerSessions -gt 0) { $msHit++ }  # logged_in
    if ($hasChatMsg)  { $msHit++ }  # first_chat
    if ($hasProposal) { $msHit++ }  # first_proposal
    if ($hasApproval) { $msHit++ }  # first_approval
    if ($hasBlocked)  { $msHit++ }  # safety_tested
    if ($hasStream)   { $msHit++ }  # first_stream
    if ($hasFeedback) { $msHit++ }  # first_feedback
    if ($partnerSessions -ge 3) { $msHit++ }  # multi_session
    $pctHit = [math]::Round($msHit / $onboardingMilestoneTotal * 100)
    $onboardingPcts += $pctHit
    if ($pctHit -eq 100) { $fullyOnboarded++ }
  } catch {
    $onboardingPcts += 0
  }
}
$avgOnboarding = if ($onboardingPcts.Count -gt 0) { [math]::Round(($onboardingPcts | Measure-Object -Average).Average) } else { 0 }
$uniqueActiveHours = $activeHours.Count

# ---------- compute grade ----------
$score = 0
$maxScore = 0
$checks = @()

# API health (20 pts)
$maxScore += 20
if ($apiOk) { $score += 20; $checks += [ordered]@{ name='api_health'; pts=20; max=20; status='pass' } }
else { $checks += [ordered]@{ name='api_health'; pts=0; max=20; status='fail' } }

# Triage exists (10 pts)
$maxScore += 10
if ($triageExists) { $score += 10; $checks += [ordered]@{ name='triage_exists'; pts=10; max=10; status='pass' } }
else { $checks += [ordered]@{ name='triage_exists'; pts=0; max=10; status='fail' } }

# Dedupe healthy (15 pts)
$maxScore += 15
if ($dedupeOk) { $score += 15; $checks += [ordered]@{ name='dedupe_healthy'; pts=15; max=15; status='pass' } }
else { $checks += [ordered]@{ name='dedupe_healthy'; pts=0; max=15; status='fail' } }

# KPI available (10 pts)
$maxScore += 10
if ($null -ne $kpi) { $score += 10; $checks += [ordered]@{ name='kpi_available'; pts=10; max=10; status='pass' } }
else { $checks += [ordered]@{ name='kpi_available'; pts=0; max=10; status='fail' } }

# Prompts in 24h > 0 (10 pts)
$maxScore += 10
$p24 = if ($null -ne $kpi) { $kpi.promptsSent24h } else { 0 }
if ($p24 -gt 0) { $score += 10; $checks += [ordered]@{ name='prompts_24h'; pts=10; max=10; status='pass'; value=$p24 } }
else { $checks += [ordered]@{ name='prompts_24h'; pts=0; max=10; status='fail'; value=$p24 } }

# Approval rate >= 50% (10 pts)
$maxScore += 10
$proposed = if ($null -ne $kpi) { $kpi.applyActionsProposed } else { 0 }
$confirmed = if ($null -ne $kpi) { $kpi.applyActionsConfirmed } else { 0 }
$approvalRate = if ($proposed -gt 0) { [math]::Round($confirmed / $proposed, 2) } else { 0 }
if ($approvalRate -ge 0.5) { $score += 10; $checks += [ordered]@{ name='approval_rate'; pts=10; max=10; status='pass'; value=$approvalRate } }
else { $checks += [ordered]@{ name='approval_rate'; pts=0; max=10; status='fail'; value=$approvalRate } }

# History has data (10 pts)
$maxScore += 10
if ($historyRows -gt 0) { $score += 10; $checks += [ordered]@{ name='history_data'; pts=10; max=10; status='pass'; value=$historyRows } }
else { $checks += [ordered]@{ name='history_data'; pts=0; max=10; status='fail'; value=0 } }

# Feedback collected (15 pts)
$maxScore += 15
if ($totalFeedback -gt 0) { $score += 15; $checks += [ordered]@{ name='feedback_collected'; pts=15; max=15; status='pass'; value=$totalFeedback } }
else { $checks += [ordered]@{ name='feedback_collected'; pts=0; max=15; status='fail'; value=0 } }

# Onboarding progress >= 75% avg (10 pts)
$maxScore += 10
if ($avgOnboarding -ge 75) { $score += 10; $checks += [ordered]@{ name='onboarding_progress'; pts=10; max=10; status='pass'; value=$avgOnboarding } }
else { $checks += [ordered]@{ name='onboarding_progress'; pts=0; max=10; status='fail'; value=$avgOnboarding } }

# Engagement spread: activity spans 2+ hours (10 pts)
$maxScore += 10
if ($uniqueActiveHours -ge 2) { $score += 10; $checks += [ordered]@{ name='engagement_spread'; pts=10; max=10; status='pass'; value=$uniqueActiveHours } }
else { $checks += [ordered]@{ name='engagement_spread'; pts=0; max=10; status='fail'; value=$uniqueActiveHours } }

$pct = if ($maxScore -gt 0) { [math]::Round(($score / $maxScore) * 100) } else { 0 }
$grade = switch ($pct) {
  { $_ -ge 90 } { 'A'; break }
  { $_ -ge 80 } { 'B'; break }
  { $_ -ge 70 } { 'C'; break }
  { $_ -ge 60 } { 'D'; break }
  default { 'F' }
}

$scorecard = [ordered]@{
  generatedAt = (Get-Date).ToString("o")
  date        = $today
  grade       = $grade
  score       = $score
  maxScore    = $maxScore
  pct         = $pct
  api         = [ordered]@{ healthy = $apiOk }
  kpi         = [ordered]@{
    available       = ($null -ne $kpi)
    promptsSent24h  = $p24
    approvalRate    = $approvalRate
    blockedAttempts = if ($null -ne $kpi) { $kpi.blockedCommandAttempts } else { 0 }
    streamRate      = if ($null -ne $kpi) { $kpi.streamSuccessRate } else { 0 }
  }
  triage      = [ordered]@{
    exists       = $triageExists
    dedupeOk     = $dedupeOk
    snapshots    = $snapshotCounts
  }
  history     = [ordered]@{
    exists    = $historyExists
    rows      = $historyRows
    days      = $partnerDays
  }
  feedback    = [ordered]@{
    total              = $totalFeedback
    partnersReporting  = $partnersWithFeedback
    partnersTotal      = $Emails.Count
  }
  onboarding  = [ordered]@{
    avgCompletion    = $avgOnboarding
    fullyOnboarded   = $fullyOnboarded
    partnersTotal    = $Emails.Count
  }
  engagement  = [ordered]@{
    totalEvents      = $totalEventCount
    uniqueActiveHours = $uniqueActiveHours
  }
  checks      = $checks
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $scorecard | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host ('  [ok] Scorecard written: ' + $OutputPath) -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $scorecard | ConvertTo-Json -Depth 8
}
else {
  $gradeColor = switch ($grade) {
    'A' { 'Green' }
    'B' { 'Green' }
    'C' { 'Yellow' }
    'D' { 'DarkYellow' }
    default { 'Red' }
  }
  Write-Host "  Grade: $grade ($pct%  $score/$maxScore)" -ForegroundColor $gradeColor
  Write-Host ''
  Write-Host '  Checks:' -ForegroundColor White
  foreach ($ch in $checks) {
    $icon = if ($ch.status -eq 'pass') { '+' } else { '-' }
    $color = if ($ch.status -eq 'pass') { 'Green' } else { 'DarkYellow' }
    $valStr = if ($null -ne $ch.value) { "  ($($ch.value))" } else { '' }
    Write-Host "    $icon $($ch.name): $($ch.pts)/$($ch.max)$valStr" -ForegroundColor $color
  }
  Write-Host ''
  Write-Host '  KPI:' -ForegroundColor White
  Write-Host "    prompts24h=$p24  approvalRate=$approvalRate  blocked=$($scorecard.kpi.blockedAttempts)  streamRate=$($scorecard.kpi.streamRate)"
  Write-Host ''
  Write-Host '  Triage:' -ForegroundColor White
  Write-Host "    exists=$triageExists  dedupeOk=$dedupeOk  snapshots=$($snapshotCounts.Count)"
  Write-Host ''
  Write-Host '  History:' -ForegroundColor White
  Write-Host "    rows=$historyRows  days=$partnerDays"
  Write-Host ''
  Write-Host '  Feedback:' -ForegroundColor White
  Write-Host "    total=$totalFeedback  partnersReporting=$partnersWithFeedback/$($Emails.Count)"
  Write-Host ''
  Write-Host '  Onboarding:' -ForegroundColor White
  Write-Host "    avgCompletion=$avgOnboarding%  fullyOnboarded=$fullyOnboarded/$($Emails.Count)"
  Write-Host ''
  Write-Host '  Engagement:' -ForegroundColor White
  Write-Host "    totalEvents=$totalEventCount  uniqueActiveHours=$uniqueActiveHours"
  Write-Host ''
}

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host ('  [warn] Triage file not found: ' + $triagePath) -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $line = "$stamp Daily scorecard: grade=$grade, score=$score/$maxScore ($pct%), prompts24h=$p24, approval=$approvalRate, feedback=$totalFeedback, history=$historyRows, dedupe=$dedupeOk, onboarding=$avgOnboarding%, activeHours=$uniqueActiveHours"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Daily scorecard:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted daily scorecard in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] Daily scorecard complete.' -ForegroundColor Green
