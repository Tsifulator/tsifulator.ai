param(
  [string]$Email = "beta.user@tsifulator.ai",
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string]$Date,
  [string]$Owner = "",
  [switch]$AppendToTriage
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$planPath = "docs/mvp-30-day-plan.md"
$triagePath = "docs/daily-triage/$today.md"

Write-Host "== tsifulator.ai beta checkpoint ($today) ==" -ForegroundColor Cyan
Write-Host "Plan snapshot: $planPath"
Write-Host "Today triage:  $triagePath"

try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health"
  Write-Host "[ok] API healthy: $($health.status)" -ForegroundColor Green
}
catch {
  Write-Host "[warn] API not reachable at $BaseUrl. Start it with: npm run dev" -ForegroundColor DarkYellow
  exit 1
}

$loginBody = @{ email = $Email } | ConvertTo-Json
$login = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType "application/json" -Body $loginBody
$headers = @{ authorization = "Bearer $($login.token)" }
$kpi = Invoke-RestMethod -Method Get -Uri "$BaseUrl/telemetry/counters" -Headers $headers

Write-Host ""
Write-Host "KPI Snapshot:" -ForegroundColor Cyan
$kpi | ConvertTo-Json -Depth 6

if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host "[warn] Triage file not found, creating it first: $triagePath" -ForegroundColor DarkYellow

    $triageScript = Join-Path $PSScriptRoot "beta-triage-new.ps1"
    if (-not (Test-Path $triageScript)) {
      Write-Host "[warn] Missing triage generator script: $triageScript" -ForegroundColor DarkYellow
      exit 1
    }

    $triageArgs = @("-ExecutionPolicy", "Bypass", "-File", $triageScript, "-Date", $today)
    if ($Owner) {
      $triageArgs += @("-Owner", $Owner)
    }

    & powershell @triageArgs
    if ($LASTEXITCODE -ne 0) {
      Write-Host "[warn] Failed to create triage file; skipping append." -ForegroundColor DarkYellow
      exit 1
    }
  }

  if (Test-Path $triagePath) {
    $counters = $kpi.counters
    $stamp = $kpi.generatedAt
    $summaryLine = "[$stamp] KPI refresh: users7d=$($counters.newBetaUsers7d), dau=$($counters.dailyActiveUsers), prompts=$($counters.promptsSent), prompts24h=$($counters.promptsSent24h), proposed=$($counters.applyActionsProposed), confirmed=$($counters.applyActionsConfirmed), blocked=$($counters.blockedCommandAttempts), stream=$($counters.streamCompletions)/$($counters.streamRequests), streamRate=$($counters.streamSuccessRate), medChatMs=$($counters.medianChatLatencyMs), medStreamFirstMs=$($counters.medianStreamFirstTokenLatencyMs)"

    $triageContent = Get-Content $triagePath -Raw
    if ($triageContent -notmatch "## Checkpoint Log") {
      $triageContent = $triageContent.TrimEnd() + "`r`n`r`n## Checkpoint Log`r`n"
    }

    $triageLines = $triageContent -split "`r?`n"
    $filteredLines = $triageLines | Where-Object { $_ -notmatch "^- \[[^\]]+\] KPI refresh:" }
    $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $summaryLine`r`n"

    Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8
    Write-Host "[ok] Upserted KPI snapshot in triage: $triagePath" -ForegroundColor Green
  }
  else {
    Write-Host "[warn] Triage file still missing after creation attempt: $triagePath" -ForegroundColor DarkYellow
    exit 1
  }
}

Write-Host ""
Write-Host "[done] Checkpoint snapshot complete." -ForegroundColor Green
