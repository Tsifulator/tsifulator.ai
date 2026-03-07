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

Write-Host '== tsifulator.ai engagement score ==' -ForegroundColor Cyan
Write-Host "Date: $today"
Write-Host "Partners: $($Emails.Count)"

# ---------- weights ----------
# Total = 100
$wSessions    = 20   # session count (more = better)
$wEvents      = 20   # event volume
$wFeedback    = 15   # feedback submissions
$wFeatures    = 20   # feature breadth (unique event types used)
$wRecency     = 15   # days since last session (lower = better)
$wDepth       = 10   # avg messages per session

# ---------- collect per-partner signals ----------
$partnerScores = @()
$now = Get-Date

foreach ($em in $Emails) {
  $sessionCount = 0
  $eventCount = 0
  $feedbackCount = 0
  $uniqueTypes = @{}
  $totalMessages = 0
  $lastSessionAt = $null

  try {
    $lb = @{ email = $em } | ConvertTo-Json
    $lg = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $lb
    $hd = @{ authorization = "Bearer $($lg.token)" }
    $sess = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=100" -Headers $hd
    $sessionCount = if ($sess.sessions) { $sess.sessions.Count } else { 0 }

    foreach ($s in $sess.sessions) {
      # Track most recent session
      if ($s.createdAt) {
        try {
          $sDate = [DateTime]::Parse($s.createdAt)
          if ($null -eq $lastSessionAt -or $sDate -gt $lastSessionAt) { $lastSessionAt = $sDate }
        } catch {}
      }

      try {
        $ev = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($s.id)/events" -Headers $hd
        $eventCount += $ev.events.Count
        foreach ($e in $ev.events) {
          if ($e.type) { $uniqueTypes[$e.type] = $true }
          if ($e.type -eq 'user_feedback') { $feedbackCount++ }
        }
      } catch {}

      try {
        $msgs = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($s.id)/messages" -Headers $hd
        $totalMessages += if ($msgs.messages) { $msgs.messages.Count } else { 0 }
      } catch {}
    }
  }
  catch {
    Write-Host "  [warn] Failed to query $em" -ForegroundColor DarkYellow
  }

  $avgMsgsPerSession = if ($sessionCount -gt 0) { [math]::Round($totalMessages / $sessionCount, 1) } else { 0 }
  $daysSinceLastSession = if ($null -ne $lastSessionAt) { [math]::Max(0, [math]::Round(($now - $lastSessionAt).TotalDays, 1)) } else { 999 }
  $featureCount = $uniqueTypes.Count

  # ---------- score each dimension (0-100 normalized, then weighted) ----------

  # Sessions: 1=10, 3=50, 5=80, 10+=100
  $sSessions = [math]::Min(100, [math]::Round($sessionCount / 10 * 100))

  # Events: 10=10, 50=50, 100=80, 200+=100
  $sEvents = [math]::Min(100, [math]::Round($eventCount / 200 * 100))

  # Feedback: 1=20, 3=50, 5=80, 10+=100
  $sFeedback = [math]::Min(100, [math]::Round($feedbackCount / 10 * 100))

  # Features: out of 11 known event types
  $maxTypes = 11
  $sFeatures = [math]::Min(100, [math]::Round($featureCount / $maxTypes * 100))

  # Recency: 0 days=100, 1=90, 3=70, 7=40, 14=10, 30+=0
  $sRecency = if ($daysSinceLastSession -le 0) { 100 }
              elseif ($daysSinceLastSession -le 1) { 90 }
              elseif ($daysSinceLastSession -le 3) { 70 }
              elseif ($daysSinceLastSession -le 7) { 40 }
              elseif ($daysSinceLastSession -le 14) { 10 }
              else { 0 }

  # Depth: 0=0, 2=20, 4=50, 6=70, 10+=100
  $sDepth = [math]::Min(100, [math]::Round($avgMsgsPerSession / 10 * 100))

  # Composite weighted score
  $composite = [math]::Round(
    ($sSessions * $wSessions +
     $sEvents * $wEvents +
     $sFeedback * $wFeedback +
     $sFeatures * $wFeatures +
     $sRecency * $wRecency +
     $sDepth * $wDepth) / 100
  )

  # Tier
  $tier = if ($composite -ge 80) { 'champion' }
          elseif ($composite -ge 60) { 'engaged' }
          elseif ($composite -ge 40) { 'moderate' }
          elseif ($composite -ge 20) { 'passive' }
          else { 'dormant' }

  $partnerScores += [ordered]@{
    email              = $em
    score              = $composite
    tier               = $tier
    dimensions         = [ordered]@{
      sessions  = [ordered]@{ raw = $sessionCount; normalized = $sSessions; weight = $wSessions }
      events    = [ordered]@{ raw = $eventCount; normalized = $sEvents; weight = $wEvents }
      feedback  = [ordered]@{ raw = $feedbackCount; normalized = $sFeedback; weight = $wFeedback }
      features  = [ordered]@{ raw = $featureCount; normalized = $sFeatures; weight = $wFeatures }
      recency   = [ordered]@{ raw = $daysSinceLastSession; normalized = $sRecency; weight = $wRecency }
      depth     = [ordered]@{ raw = $avgMsgsPerSession; normalized = $sDepth; weight = $wDepth }
    }
  }
}

# ---------- summary ----------
$scores = @($partnerScores | ForEach-Object { $_['score'] })
$avgScore = if ($scores.Count -gt 0) { [math]::Round(($scores | Measure-Object -Average).Average) } else { 0 }
$topPartner = if ($partnerScores.Count -gt 0) { ($partnerScores | Sort-Object { $_['score'] } -Descending)[0].email } else { 'none' }
$bottomPartner = if ($partnerScores.Count -gt 0) { ($partnerScores | Sort-Object { $_['score'] })[0].email } else { 'none' }

$overallTier = if ($avgScore -ge 80) { 'champion' }
               elseif ($avgScore -ge 60) { 'engaged' }
               elseif ($avgScore -ge 40) { 'moderate' }
               elseif ($avgScore -ge 20) { 'passive' }
               else { 'dormant' }

$result = [ordered]@{
  generatedAt   = (Get-Date).ToString("o")
  date          = $today
  avgScore      = $avgScore
  overallTier   = $overallTier
  topPartner    = $topPartner
  bottomPartner = $bottomPartner
  partners      = $partnerScores
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $result | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host "  [ok] Engagement scores written: $OutputPath" -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $result | ConvertTo-Json -Depth 8
}
else {
  $tierColor = switch ($overallTier) {
    'champion' { 'Green' }
    'engaged'  { 'Green' }
    'moderate' { 'Yellow' }
    'passive'  { 'DarkYellow' }
    'dormant'  { 'Red' }
    default    { 'White' }
  }
  Write-Host "  Average Score: $avgScore/100 ($overallTier)" -ForegroundColor $tierColor
  Write-Host "  Top: $topPartner  Bottom: $bottomPartner"
  Write-Host ''

  # Leaderboard
  Write-Host '  Leaderboard:' -ForegroundColor White
  $headerLine = "    {0,-3} {1,-28} {2,-6} {3,-10} {4,-5} {5,-5} {6,-5} {7,-5} {8,-5} {9,-5}" -f '#', 'Email', 'Score', 'Tier', 'Sess', 'Evts', 'Fdbk', 'Feat', 'Rcnc', 'Dpth'
  Write-Host $headerLine -ForegroundColor DarkGray
  $rank = 0
  $sorted = $partnerScores | Sort-Object { $_['score'] } -Descending
  foreach ($ps in $sorted) {
    $rank++
    $d = $ps.dimensions
    $color = switch ($ps.tier) {
      'champion' { 'Green' }
      'engaged'  { 'Green' }
      'moderate' { 'Yellow' }
      'passive'  { 'DarkYellow' }
      'dormant'  { 'Red' }
      default    { 'White' }
    }
    $line = "    {0,-3} {1,-28} {2,-6} {3,-10} {4,-5} {5,-5} {6,-5} {7,-5} {8,-5} {9,-5}" -f $rank, $ps.email, $ps.score, $ps.tier, $d.sessions.normalized, $d.events.normalized, $d.feedback.normalized, $d.features.normalized, $d.recency.normalized, $d.depth.normalized
    Write-Host $line -ForegroundColor $color
  }

  # Score bar chart
  Write-Host ''
  Write-Host '  Score Distribution:' -ForegroundColor White
  $maxBar = 30
  foreach ($ps in $sorted) {
    $barLen = [math]::Max(1, [math]::Round($ps.score / 100 * $maxBar))
    $bar = ('=' * $barLen).PadRight($maxBar)
    $color = switch ($ps.tier) {
      'champion' { 'Green' }
      'engaged'  { 'Green' }
      'moderate' { 'Yellow' }
      'passive'  { 'DarkYellow' }
      'dormant'  { 'Red' }
      default    { 'White' }
    }
    $emailShort = if ($ps.email.Length -gt 25) { $ps.email.Substring(0, 25) } else { $ps.email }
    $line = "    {0,-25} |{1}| {2}/100" -f $emailShort, $bar, $ps.score
    Write-Host $line -ForegroundColor $color
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
  $partnerStr = ($partnerScores | ForEach-Object { "$($_.email)=$($_.score)($($_.tier))" }) -join ', '
  $line = "$stamp Engagement score: avg=$avgScore, tier=$overallTier, top=$topPartner, $partnerStr"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Engagement score:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host "  [ok] Upserted engagement score in triage: $triagePath" -ForegroundColor Green
}

Write-Host '[done] Engagement score complete.' -ForegroundColor Green
