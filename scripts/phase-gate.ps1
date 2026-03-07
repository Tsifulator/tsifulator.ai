param(
  [string]$Phase = "phase"
)

$ErrorActionPreference = "Stop"

Write-Host "== phase gate: $Phase ==" -ForegroundColor Cyan

$scriptPath = Join-Path (Get-Location) "package.json"
if (-not (Test-Path $scriptPath)) {
  Write-Host "[fail] package.json not found" -ForegroundColor Red
  exit 1
}

$package = Get-Content package.json -Raw | ConvertFrom-Json
$scripts = @{}
if ($package.scripts) {
  $package.scripts.PSObject.Properties | ForEach-Object {
    $scripts[$_.Name] = $_.Value
  }
}

$checks = @("lint", "typecheck", "test", "build")
$ranAny = $false

foreach ($check in $checks) {
  if ($scripts.ContainsKey($check)) {
    $ranAny = $true
    Write-Host "[run] npm run $check" -ForegroundColor Yellow
    npm run $check
    if ($LASTEXITCODE -ne 0) {
      Write-Host "[fail] $check failed" -ForegroundColor Red
      exit 1
    }
    Write-Host "[ok] $check" -ForegroundColor Green
  } else {
    Write-Host "[skip] missing script: $check" -ForegroundColor DarkYellow
  }
}

if (-not $ranAny) {
  Write-Host "[warn] no quality scripts found yet; gate passed with skips" -ForegroundColor Yellow
}

Write-Host "[ok] phase gate complete: $Phase" -ForegroundColor Green
