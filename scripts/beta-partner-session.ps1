param(
  [string]$PartnerName = "Design Partner",
  [string]$PartnerEmail,
  [string]$SessionGoal = "Validate one real workflow end-to-end",
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string]$Owner = "",
  [switch]$DryRun
)

$ErrorActionPreference = "Stop"

function Get-SafeSlug {
  param([string]$Value)

  $slug = $Value.ToLowerInvariant() -replace "[^a-z0-9]+", "-"
  $slug = $slug.Trim("-")
  if (-not $slug) {
    return "partner"
  }

  return $slug
}

$today = Get-Date -Format "yyyy-MM-dd"
if (-not $PartnerEmail) {
  $PartnerEmail = "partner.$($today -replace '-', '')@tsifulator.ai"
}

$workspace = Get-Location
$triageScript = Join-Path $workspace "scripts\beta-triage-new.ps1"
$checkpointScript = Join-Path $workspace "scripts\beta-checkpoint-today.ps1"

if (-not (Test-Path $triageScript)) {
  throw "Missing script: $triageScript"
}

if (-not (Test-Path $checkpointScript)) {
  throw "Missing script: $checkpointScript"
}

$partnerSlug = Get-SafeSlug -Value $PartnerName
$sessionDir = Join-Path $workspace "docs\partner-sessions"
if (-not (Test-Path $sessionDir)) {
  New-Item -ItemType Directory -Path $sessionDir | Out-Null
}

$sessionPath = Join-Path $sessionDir "$today-$partnerSlug.md"
if (-not (Test-Path $sessionPath)) {
  $sessionTemplate = @"
# Design Partner Session - $PartnerName

## Meta
- Date: $today
- Partner: $PartnerName
- Partner email: $PartnerEmail
- Owner: $Owner
- Session goal: $SessionGoal

## Pre-session checklist
- [ ] API running (`npm run dev`)
- [ ] CLI running (`npm run cli`)
- [ ] Baseline checkpoint captured (`npm run beta:checkpoint:append`)

## During session
- [ ] Capture at least 3 `/feedback` entries from real user actions
- [ ] Run one safe apply action and confirm execution
- [ ] Run one blocked-command simulation to confirm policy behavior

## Outcomes
- Wins:
- Frictions:
- Top requests:

## Follow-ups (next 24h)
1.
2.
3.
"@

  $sessionTemplate | Out-File -FilePath $sessionPath -Encoding utf8
}

Write-Host "== tsifulator.ai design partner bootstrap ==" -ForegroundColor Cyan
Write-Host "Partner: $PartnerName <$PartnerEmail>"
Write-Host "Session goal: $SessionGoal"
Write-Host "Session notes: $sessionPath"

if ($DryRun) {
  Write-Host "[dry-run] would ensure today's triage exists" -ForegroundColor DarkYellow
  Write-Host "[dry-run] would append checkpoint KPI snapshot" -ForegroundColor DarkYellow
  Write-Host "[dry-run] would submit kickoff feedback note" -ForegroundColor DarkYellow
  Write-Host "[done] Dry-run complete." -ForegroundColor Green
  exit 0
}

$triageArgs = @("-ExecutionPolicy", "Bypass", "-File", $triageScript, "-Date", $today)
if ($Owner) {
  $triageArgs += @("-Owner", $Owner)
}
& powershell @triageArgs
if ($LASTEXITCODE -ne 0) {
  throw "Failed to create/verify daily triage file"
}

& powershell -ExecutionPolicy Bypass -File $checkpointScript -AppendToTriage -Date $today -Owner $Owner -Email $PartnerEmail -BaseUrl $BaseUrl
if ($LASTEXITCODE -ne 0) {
  throw "Failed to append daily checkpoint"
}

try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health"
  if ($health.status -ne "ok") {
    throw "API health is not ok"
  }
}
catch {
  throw "API not reachable at $BaseUrl. Start it with: npm run dev"
}

$loginBody = @{ email = $PartnerEmail } | ConvertTo-Json
$login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType "application/json" -Body $loginBody
$headers = @{ authorization = "Bearer $($login.token)" }

$kickoffText = "design-partner kickoff: partner='$PartnerName', goal='$SessionGoal'"
$feedbackBody = @{ text = $kickoffText } | ConvertTo-Json
$feedback = Invoke-RestMethod -Method Post -Uri "$BaseUrl/feedback" -Headers $headers -ContentType "application/json" -Body $feedbackBody

Write-Host "[ok] Kickoff feedback captured (session: $($feedback.sessionId))" -ForegroundColor Green
Write-Host "[next] Start live session: npm run cli (login as $PartnerEmail)" -ForegroundColor Cyan
Write-Host "[next] Capture 3+ user feedback entries with /feedback" -ForegroundColor Cyan
Write-Host "[done] Partner session bootstrap complete." -ForegroundColor Green
