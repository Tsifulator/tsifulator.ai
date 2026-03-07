param(
  [string]$Email = "beta.user@tsifulator.ai",
  [string]$Owner = "",
  [switch]$SkipInstall,
  [switch]$StartServer,
  [switch]$RunTrafficWarmup
)

$ErrorActionPreference = "Stop"

Write-Host "== tsifulator.ai beta onboarding ==" -ForegroundColor Cyan

function Assert-Command {
  param(
    [string]$Name,
    [string]$VersionArg = "--version"
  )

  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Missing required command: $Name"
  }

  try {
    $version = & $Name $VersionArg 2>$null
    Write-Host "[ok] $Name $version" -ForegroundColor Green
  }
  catch {
    Write-Host "[ok] $Name" -ForegroundColor Green
  }
}

Assert-Command -Name "node" -VersionArg "-v"
Assert-Command -Name "npm" -VersionArg "-v"

if (-not (Test-Path ".env")) {
  if (Test-Path ".env.example") {
    Copy-Item ".env.example" ".env"
    Write-Host "[created] .env from .env.example" -ForegroundColor DarkYellow
  }
  else {
    throw "Missing .env and .env.example"
  }
} else {
  Write-Host "[ok] .env present" -ForegroundColor Green
}

if (-not $SkipInstall) {
  Write-Host "[run] npm install" -ForegroundColor Yellow
  npm install
  if ($LASTEXITCODE -ne 0) {
    throw "npm install failed"
  }
}
else {
  Write-Host "[skip] npm install" -ForegroundColor DarkYellow
}

Write-Host "[run] npx tsc --noEmit" -ForegroundColor Yellow
npx tsc --noEmit
if ($LASTEXITCODE -ne 0) {
  throw "Typecheck failed"
}

Write-Host "[run] npm test" -ForegroundColor Yellow
npm test
if ($LASTEXITCODE -ne 0) {
  throw "Tests failed"
}

Write-Host "\n[info] Quick beta start commands:" -ForegroundColor Cyan
Write-Host "1) Terminal A: npm run dev"
Write-Host "2) Terminal B: npm run cli"
Write-Host "3) In CLI login email: $Email"
Write-Host "4) In CLI try: /kpi"
Write-Host "5) In CLI try: /feedback onboarding works"

if ($RunTrafficWarmup) {
  $warmupScript = Join-Path $PSScriptRoot "beta-traffic-warmup.ps1"
  $checkpointScript = Join-Path $PSScriptRoot "beta-checkpoint-today.ps1"

  if (-not (Test-Path $warmupScript)) {
    throw "Missing warmup script: $warmupScript"
  }

  if (-not (Test-Path $checkpointScript)) {
    throw "Missing checkpoint script: $checkpointScript"
  }

  Write-Host "\n[run] traffic warmup KPI check" -ForegroundColor Yellow
  & powershell -ExecutionPolicy Bypass -File $warmupScript -Email $Email
  if ($LASTEXITCODE -ne 0) {
    throw "Traffic warmup failed"
  }

  Write-Host "[run] append onboarding checkpoint to daily triage" -ForegroundColor Yellow
  $checkpointArgs = @("-ExecutionPolicy", "Bypass", "-File", $checkpointScript, "-Email", $Email, "-AppendToTriage")
  if ($Owner) {
    $checkpointArgs += @("-Owner", $Owner)
  }

  & powershell @checkpointArgs
  if ($LASTEXITCODE -ne 0) {
    throw "Checkpoint append failed"
  }
}
else {
  Write-Host "[skip] traffic warmup (use -RunTrafficWarmup to enable)" -ForegroundColor DarkYellow
}

if ($StartServer) {
  Write-Host "\n[run] starting API server (foreground): npm run dev" -ForegroundColor Yellow
  npm run dev
}

Write-Host "\n[done] Beta onboarding checks passed." -ForegroundColor Green
