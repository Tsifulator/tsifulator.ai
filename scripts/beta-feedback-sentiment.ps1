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

Write-Host '== tsifulator.ai feedback sentiment ==' -ForegroundColor Cyan
Write-Host "Date:     $today"
Write-Host "Partners: $($Emails -join ', ')"

# ---------- keyword dictionaries ----------
$positiveKeywords = @(
  'good', 'great', 'excellent', 'safe', 'clear', 'helpful', 'easy',
  'intuitive', 'fast', 'smooth', 'nice', 'love', 'awesome', 'perfect',
  'works well', 'well done', 'useful', 'reliable', 'satisfied', 'happy',
  'impressed', 'efficient', 'convenient', 'comfortable', 'confident'
)

$negativeKeywords = @(
  'bad', 'broken', 'error', 'fail', 'crash', 'slow', 'bug', 'wrong',
  'terrible', 'awful', 'frustrat', 'annoying', 'disappoint', 'useless',
  'unusable', 'unreliable', 'difficult', 'painful', 'horrible', 'hate',
  'confus', 'unclear', 'mislead', 'unexpected', 'stuck'
)

$suggestionKeywords = @(
  'should', 'could', 'would be', 'suggest', 'wish', 'prefer', 'idea',
  'improve', 'shorter', 'longer', 'better if', 'consider', 'maybe',
  'feature request', 'add', 'change', 'tweak', 'option', 'able to',
  'point:', 'proposal', 'wording', 'how about', 'can be'
)

function Get-Sentiment {
  param([string]$Text)

  $lower = $Text.ToLower()

  $posHits = @($positiveKeywords | Where-Object { $lower -match [regex]::Escape($_) })
  $negHits = @($negativeKeywords | Where-Object { $lower -match [regex]::Escape($_) })
  $sugHits = @($suggestionKeywords | Where-Object { $lower -match [regex]::Escape($_) })

  $posCount = $posHits.Count
  $negCount = $negHits.Count
  $sugCount = $sugHits.Count

  # Suggestion takes priority when mixed with negative (constructive)
  if ($sugCount -gt 0 -and $sugCount -ge $negCount) {
    return [ordered]@{
      sentiment = 'suggestion'
      confidence = [math]::Round($sugCount / ($posCount + $negCount + $sugCount + 0.01), 2)
      keywords = $sugHits
    }
  }

  if ($posCount -gt $negCount) {
    return [ordered]@{
      sentiment = 'positive'
      confidence = [math]::Round($posCount / ($posCount + $negCount + $sugCount + 0.01), 2)
      keywords = $posHits
    }
  }

  if ($negCount -gt $posCount) {
    return [ordered]@{
      sentiment = 'negative'
      confidence = [math]::Round($negCount / ($posCount + $negCount + $sugCount + 0.01), 2)
      keywords = $negHits
    }
  }

  return [ordered]@{
    sentiment = 'neutral'
    confidence = 0
    keywords = @()
  }
}

# ---------- health check ----------
try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health" -TimeoutSec 5
  Write-Host ('[ok] API healthy: ' + $health.status) -ForegroundColor Green
}
catch {
  Write-Host '[error] API unreachable' -ForegroundColor Red
  exit 1
}

# ---------- collect and classify ----------
$allItems = @()

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

  try {
    $sessionsResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=$SessionLimit" -Headers $headers
    $sessions = $sessionsResp.sessions
  }
  catch { continue }

  if (-not $sessions -or $sessions.Count -eq 0) { continue }

  foreach ($session in $sessions) {
    try {
      $eventsResp = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($session.id)/events" -Headers $headers
      $events = $eventsResp.events
    }
    catch { continue }

    if (-not $events) { continue }

    foreach ($evt in $events) {
      if ($evt.type -ne 'user_feedback') { continue }

      $payload = $null
      try { $payload = $evt.payload | ConvertFrom-Json -ErrorAction SilentlyContinue } catch {}
      if (-not $payload) { try { $payload = $evt.payload } catch { continue } }
      if (-not $payload) { continue }

      $feedbackText = ''
      if ($payload -is [string]) { $feedbackText = $payload }
      elseif ($null -ne $payload.text) { $feedbackText = [string]$payload.text }
      if (-not $feedbackText) { continue }

      $result = Get-Sentiment -Text $feedbackText

      $allItems += [ordered]@{
        email      = $emailClean
        sessionId  = $session.id
        createdAt  = $evt.createdAt
        text       = $feedbackText
        sentiment  = $result.sentiment
        confidence = $result.confidence
        keywords   = $result.keywords
      }
    }
  }
}

# ---------- aggregate ----------
$totalItems = $allItems.Count
$posItems = @($allItems | Where-Object { $_.sentiment -eq 'positive' })
$negItems = @($allItems | Where-Object { $_.sentiment -eq 'negative' })
$sugItems = @($allItems | Where-Object { $_.sentiment -eq 'suggestion' })
$neuItems = @($allItems | Where-Object { $_.sentiment -eq 'neutral' })

$posCount = $posItems.Count
$negCount = $negItems.Count
$sugCount = $sugItems.Count
$neuCount = $neuItems.Count

$posRate = if ($totalItems -gt 0) { [math]::Round($posCount / $totalItems, 2) } else { 0 }
$negRate = if ($totalItems -gt 0) { [math]::Round($negCount / $totalItems, 2) } else { 0 }
$sugRate = if ($totalItems -gt 0) { [math]::Round($sugCount / $totalItems, 2) } else { 0 }

# Per-partner breakdown
$perPartner = @()
foreach ($email in $Emails) {
  $partnerItems = @($allItems | Where-Object { $_.email -eq $email })
  $pPos = @($partnerItems | Where-Object { $_.sentiment -eq 'positive' }).Count
  $pNeg = @($partnerItems | Where-Object { $_.sentiment -eq 'negative' }).Count
  $pSug = @($partnerItems | Where-Object { $_.sentiment -eq 'suggestion' }).Count
  $pNeu = @($partnerItems | Where-Object { $_.sentiment -eq 'neutral' }).Count

  $perPartner += [ordered]@{
    email    = $email
    total    = $partnerItems.Count
    positive = $pPos
    negative = $pNeg
    suggestion = $pSug
    neutral  = $pNeu
  }
}

# Top keywords across all feedback
$allKeywords = $allItems | ForEach-Object { $_.keywords } | Where-Object { $_ }
$keywordFreq = @{}
foreach ($kw in $allKeywords) {
  if ($keywordFreq.ContainsKey($kw)) { $keywordFreq[$kw]++ }
  else { $keywordFreq[$kw] = 1 }
}
$topKeywords = $keywordFreq.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10

$sentiment = [ordered]@{
  generatedAt = (Get-Date).ToString("o")
  date        = $today
  total       = $totalItems
  breakdown   = [ordered]@{
    positive   = $posCount
    negative   = $negCount
    suggestion = $sugCount
    neutral    = $neuCount
  }
  rates       = [ordered]@{
    positiveRate   = $posRate
    negativeRate   = $negRate
    suggestionRate = $sugRate
  }
  topKeywords = @($topKeywords | ForEach-Object { [ordered]@{ keyword = $_.Key; count = $_.Value } })
  perPartner  = $perPartner
  items       = $allItems
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $sentiment | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host ('  [ok] Sentiment report written: ' + $OutputPath) -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $sentiment | ConvertTo-Json -Depth 8
}
else {
  Write-Host "  Total feedback: $totalItems" -ForegroundColor White
  Write-Host ''
  Write-Host '  Sentiment breakdown:' -ForegroundColor White
  Write-Host "    + positive:   $posCount ($([math]::Round($posRate * 100))%)" -ForegroundColor Green
  Write-Host "    - negative:   $negCount ($([math]::Round($negRate * 100))%)" -ForegroundColor $(if ($negCount -gt 0) { 'Red' } else { 'Gray' })
  Write-Host "    ~ suggestion: $sugCount ($([math]::Round($sugRate * 100))%)" -ForegroundColor Yellow
  Write-Host "    . neutral:    $neuCount" -ForegroundColor Gray
  Write-Host ''

  if ($topKeywords.Count -gt 0) {
    Write-Host '  Top keywords:' -ForegroundColor White
    foreach ($kw in $topKeywords) {
      Write-Host "    $($kw.Key) ($($kw.Value)x)" -ForegroundColor Cyan
    }
    Write-Host ''
  }

  Write-Host '  Per partner:' -ForegroundColor White
  foreach ($pp in $perPartner) {
    Write-Host "    $($pp.email): total=$($pp.total), pos=$($pp.positive), neg=$($pp.negative), sug=$($pp.suggestion), neu=$($pp.neutral)" -ForegroundColor Gray
  }
  Write-Host ''

  # Show suggestion items specifically (actionable)
  if ($sugItems.Count -gt 0) {
    Write-Host '  Actionable suggestions:' -ForegroundColor Yellow
    foreach ($si in $sugItems) {
      $shortDate = $si.createdAt
      if ($shortDate.Length -gt 19) { $shortDate = $shortDate.Substring(0, 19) }
      Write-Host "    [$($si.email)] $shortDate" -ForegroundColor DarkYellow
      Write-Host "      $($si.text)" -ForegroundColor White
    }
    Write-Host ''
  }

  # Negative items (needs attention)
  if ($negItems.Count -gt 0) {
    Write-Host '  Negative feedback (needs attention):' -ForegroundColor Red
    foreach ($ni in $negItems) {
      $shortDate = $ni.createdAt
      if ($shortDate.Length -gt 19) { $shortDate = $shortDate.Substring(0, 19) }
      Write-Host "    [$($ni.email)] $shortDate" -ForegroundColor Red
      Write-Host "      $($ni.text)" -ForegroundColor White
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
  $topKwStr = if ($topKeywords.Count -gt 0) {
    ($topKeywords | Select-Object -First 5 | ForEach-Object { "$($_.Key)($($_.Value))" }) -join ', '
  } else { 'none' }

  $line = "$stamp Feedback sentiment: total=$totalItems, positive=$posCount($([math]::Round($posRate * 100))%), negative=$negCount($([math]::Round($negRate * 100))%), suggestion=$sugCount($([math]::Round($sugRate * 100))%), topKeywords=[$topKwStr]"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Feedback sentiment:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted feedback sentiment in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] Feedback sentiment complete.' -ForegroundColor Green
