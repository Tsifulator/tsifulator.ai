param(
  [string]$Email = "beta.user@tsifulator.ai",
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string]$Message = "hello from beta traffic warmup",
  [switch]$Json
)

$ErrorActionPreference = "Stop"

Write-Host "== tsifulator.ai beta traffic warmup ==" -ForegroundColor Cyan

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

Write-Host "[run] sending /chat request" -ForegroundColor Cyan
$chatBody = @{ message = $Message; cwd = (Get-Location).Path; lastOutput = "" } | ConvertTo-Json
$chat = Invoke-RestMethod -Method Post -Uri "$BaseUrl/chat" -ContentType "application/json" -Headers $headers -Body $chatBody
$sessionId = $chat.sessionId

if (-not $sessionId) {
  Write-Host "[warn] /chat did not return sessionId" -ForegroundColor DarkYellow
  exit 1
}

Write-Host "[run] sending /chat/stream request" -ForegroundColor Cyan
$streamMessage = "$Message stream"
$streamUri = "$BaseUrl/chat/stream?sessionId=$([uri]::EscapeDataString($sessionId))&message=$([uri]::EscapeDataString($streamMessage))&cwd=$([uri]::EscapeDataString((Get-Location).Path))&lastOutput="
$streamResponse = Invoke-WebRequest -Method Get -Uri $streamUri -Headers $headers -UseBasicParsing

if ($streamResponse.StatusCode -lt 200 -or $streamResponse.StatusCode -ge 300) {
  Write-Host "[warn] /chat/stream failed with status $($streamResponse.StatusCode)" -ForegroundColor DarkYellow
  exit 1
}

$kpi = Invoke-RestMethod -Method Get -Uri "$BaseUrl/telemetry/counters" -Headers $headers

Write-Host ""
if ($Json) {
  $kpi | ConvertTo-Json -Depth 6
}
else {
  $c = $kpi.counters
  $streamPct = [Math]::Round(($c.streamSuccessRate * 100), 1)

  Write-Host "KPI counters @ $($kpi.generatedAt)" -ForegroundColor Cyan
  Write-Host "- Prompts sent (all-time): $($c.promptsSent)"
  Write-Host "- Stream requests: $($c.streamRequests)"
  Write-Host "- Stream completions: $($c.streamCompletions)"
  Write-Host "- Stream success rate: $streamPct%"
  Write-Host "- Median /chat latency: $($c.medianChatLatencyMs) ms"
  Write-Host "- Median /chat/stream first-token latency: $($c.medianStreamFirstTokenLatencyMs) ms"
}

Write-Host ""
Write-Host "[done] Traffic warmup complete for $Email (session: $sessionId)" -ForegroundColor Green
