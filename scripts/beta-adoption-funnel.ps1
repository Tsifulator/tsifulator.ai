param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
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

Write-Host '== tsifulator.ai feature adoption funnel ==' -ForegroundColor Cyan
Write-Host "Date: $today"
Write-Host "Partners: $($Emails.Count)"

# ---------- funnel stages ----------
# Stage 1: logged_in  (has >= 1 session)
# Stage 2: chatted    (sent a chat_user_message)
# Stage 3: proposed   (received an action_proposed event)
# Stage 4: approved   (has an action_executed event)
# Stage 5: feedback   (submitted user_feedback)
# Stage 6: multi_sess (has >= 3 sessions)

$stageNames = @('logged_in', 'chatted', 'proposed', 'approved', 'feedback', 'multi_session')
$stageLabels = @('Logged In', 'Sent Chat', 'Saw Proposal', 'Approved Action', 'Gave Feedback', 'Multi-Session (3+)')

$partnerDetails = @()
$stageTotals = @(0, 0, 0, 0, 0, 0)

foreach ($em in $Emails) {
  $stages = @($false, $false, $false, $false, $false, $false)
  $sessionCount = 0
  $eventTypes = @{}
  $dropOff = 'none'

  try {
    $lb = @{ email = $em } | ConvertTo-Json
    $lg = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $lb
    $hd = @{ authorization = "Bearer $($lg.token)" }
    $sess = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=100" -Headers $hd
    $sessionCount = if ($sess.sessions) { $sess.sessions.Count } else { 0 }

    # Stage 1: logged in
    if ($sessionCount -gt 0) { $stages[0] = $true }

    # Scan events
    foreach ($s in $sess.sessions) {
      try {
        $ev = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($s.id)/events" -Headers $hd
        foreach ($e in $ev.events) {
          if ($e.type) { $eventTypes[$e.type] = $true }
        }
      } catch {}
    }

    # Stage 2: chatted
    if ($eventTypes['chat_user_message']) { $stages[1] = $true }

    # Stage 3: proposal
    if ($eventTypes['action_proposed']) { $stages[2] = $true }

    # Stage 4: approved
    if ($eventTypes['action_executed']) { $stages[3] = $true }

    # Stage 5: feedback
    if ($eventTypes['user_feedback']) { $stages[4] = $true }

    # Stage 6: multi-session
    if ($sessionCount -ge 3) { $stages[5] = $true }
  }
  catch {
    Write-Host "  [warn] Failed to query $em" -ForegroundColor DarkYellow
  }

  # Find drop-off point: first stage that is false after a true
  $reachedStage = -1
  for ($i = 0; $i -lt $stages.Count; $i++) {
    if ($stages[$i]) {
      $reachedStage = $i
      $stageTotals[$i]++
    }
  }
  # Drop-off = first false stage (or 'complete' if all true)
  $allPassed = ($stages | Where-Object { -not $_ }).Count -eq 0
  if ($allPassed) {
    $dropOff = 'complete'
  } else {
    for ($i = 0; $i -lt $stages.Count; $i++) {
      if (-not $stages[$i]) {
        $dropOff = $stageNames[$i]
        break
      }
    }
  }

  $partnerDetails += [ordered]@{
    email        = $em
    sessions     = $sessionCount
    stages       = $stages
    stageMap     = [ordered]@{
      logged_in     = $stages[0]
      chatted       = $stages[1]
      proposed      = $stages[2]
      approved      = $stages[3]
      feedback      = $stages[4]
      multi_session = $stages[5]
    }
    dropOff      = $dropOff
    reachedStage = $reachedStage
  }
}

# ---------- compute funnel metrics ----------
$totalPartners = $Emails.Count
$conversionRates = @()
$dropOffRates = @()
for ($i = 0; $i -lt $stageNames.Count; $i++) {
  $rate = if ($totalPartners -gt 0) { [math]::Round($stageTotals[$i] / $totalPartners * 100) } else { 0 }
  $conversionRates += $rate
  if ($i -eq 0) {
    $dropOff = $totalPartners - $stageTotals[$i]
  } else {
    $dropOff = $stageTotals[$i - 1] - $stageTotals[$i]
  }
  $dropOffRates += $dropOff
}

# Find biggest bottleneck (largest drop from previous stage)
$bottleneckIdx = 0
$maxDrop = 0
for ($i = 0; $i -lt $stageNames.Count; $i++) {
  if ($dropOffRates[$i] -gt $maxDrop) {
    $maxDrop = $dropOffRates[$i]
    $bottleneckIdx = $i
  }
}
$bottleneck = if ($maxDrop -gt 0) { $stageNames[$bottleneckIdx] } else { 'none' }

# Overall health
$fullConversion = $conversionRates[-1]
$health = if ($fullConversion -ge 80) { 'excellent' }
           elseif ($fullConversion -ge 60) { 'good' }
           elseif ($fullConversion -ge 40) { 'moderate' }
           elseif ($fullConversion -ge 20) { 'weak' }
           else { 'critical' }

$funnel = [ordered]@{
  generatedAt    = (Get-Date).ToString("o")
  date           = $today
  totalPartners  = $totalPartners
  health         = $health
  bottleneck     = $bottleneck
  fullConversion = $fullConversion
  stages         = @()
  partners       = @()
}

for ($i = 0; $i -lt $stageNames.Count; $i++) {
  $funnel.stages += [ordered]@{
    name           = $stageNames[$i]
    label          = $stageLabels[$i]
    reached        = $stageTotals[$i]
    total          = $totalPartners
    conversionPct  = $conversionRates[$i]
    droppedOff     = $dropOffRates[$i]
  }
}

foreach ($pd in $partnerDetails) {
  $funnel.partners += [ordered]@{
    email    = $pd.email
    sessions = $pd.sessions
    stages   = $pd.stageMap
    dropOff  = $pd.dropOff
  }
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $funnel | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host "  [ok] Funnel written: $OutputPath" -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $funnel | ConvertTo-Json -Depth 8
}
else {
  # Funnel visualization
  $maxBar = 30
  Write-Host '  Adoption Funnel:' -ForegroundColor White
  Write-Host ''
  for ($i = 0; $i -lt $stageNames.Count; $i++) {
    $pct = $conversionRates[$i]
    $barLen = [math]::Max(1, [math]::Round($pct / 100 * $maxBar))
    $bar = ('=' * $barLen).PadRight($maxBar)
    $color = if ($pct -ge 80) { 'Green' } elseif ($pct -ge 50) { 'Yellow' } else { 'Red' }
    $label = $stageLabels[$i].PadRight(22)
    $line = "    $label |$bar| $($stageTotals[$i])/$totalPartners ($pct%)"
    Write-Host $line -ForegroundColor $color
  }

  Write-Host ''
  Write-Host '  Summary:' -ForegroundColor White
  Write-Host "    health=$health  fullConversion=$fullConversion%  bottleneck=$bottleneck"

  # Per-partner breakdown
  Write-Host ''
  Write-Host '  Per-Partner:' -ForegroundColor White
  $headerLine = "    {0,-28} {1,-8} {2,-7} {3,-7} {4,-7} {5,-7} {6,-7} {7,-7} {8}" -f 'Email', 'Sess', 'Login', 'Chat', 'Prop', 'Appr', 'Fdbk', 'Multi', 'DropOff'
  Write-Host $headerLine -ForegroundColor DarkGray
  foreach ($pd in $partnerDetails) {
    $marks = @()
    for ($i = 0; $i -lt 6; $i++) {
      $marks += if ($pd.stages[$i]) { '+' } else { '-' }
    }
    $line = "    {0,-28} {1,-8} {2,-7} {3,-7} {4,-7} {5,-7} {6,-7} {7,-7} {8}" -f $pd.email, $pd.sessions, $marks[0], $marks[1], $marks[2], $marks[3], $marks[4], $marks[5], $pd.dropOff
    $col = if ($pd.dropOff -eq 'complete') { 'Green' } else { 'DarkYellow' }
    Write-Host $line -ForegroundColor $col
  }
  Write-Host ''
}

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host "  [warn] Triage file not found: $triagePath" -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $stageStr = ($stageNames | ForEach-Object {
    $idx = [array]::IndexOf($stageNames, $_)
    "$_=$($conversionRates[$idx])%"
  }) -join ', '
  $line = "$stamp Adoption funnel: health=$health, fullConversion=$fullConversion%, bottleneck=$bottleneck, $stageStr"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Adoption funnel:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host "  [ok] Upserted adoption funnel in triage: $triagePath" -ForegroundColor Green
}

Write-Host '[done] Feature adoption funnel complete.' -ForegroundColor Green
