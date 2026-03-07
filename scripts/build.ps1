#!/usr/bin/env pwsh
<#
.SYNOPSIS
  Reliable build script for tsifulator.ai
.DESCRIPTION
  Runs type-check, lint, format-check, and tests in sequence.
  Exits non-zero on first failure.
#>
param(
  [switch]$SkipTests,
  [switch]$Fix
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot

Write-Host "=== tsifulator.ai build ===" -ForegroundColor Cyan

# 1. Install deps if needed
if (-not (Test-Path "$root/node_modules")) {
  Write-Host "[1/5] Installing dependencies..." -ForegroundColor Yellow
  Push-Location $root
  pnpm install
  Pop-Location
} else {
  Write-Host "[1/5] Dependencies OK" -ForegroundColor Green
}

# 2. Type-check
Write-Host "[2/5] Type-checking..." -ForegroundColor Yellow
Push-Location $root
npx tsc --noEmit
if ($LASTEXITCODE -ne 0) { Write-Host "FAIL: Type errors" -ForegroundColor Red; exit 1 }
Pop-Location
Write-Host "  OK" -ForegroundColor Green

# 3. Lint
if ($Fix) {
  Write-Host "[3/5] Lint (auto-fix)..." -ForegroundColor Yellow
  Push-Location $root; pnpm run lint:fix; Pop-Location
} else {
  Write-Host "[3/5] Lint..." -ForegroundColor Yellow
  Push-Location $root; pnpm run lint; Pop-Location
}
if ($LASTEXITCODE -ne 0) { Write-Host "FAIL: Lint errors" -ForegroundColor Red; exit 1 }
Write-Host "  OK" -ForegroundColor Green

# 4. Format check
if ($Fix) {
  Write-Host "[4/5] Format (auto-fix)..." -ForegroundColor Yellow
  Push-Location $root; pnpm run format; Pop-Location
} else {
  Write-Host "[4/5] Format check..." -ForegroundColor Yellow
  Push-Location $root; pnpm run format:check; Pop-Location
}
if ($LASTEXITCODE -ne 0) { Write-Host "FAIL: Formatting issues" -ForegroundColor Red; exit 1 }
Write-Host "  OK" -ForegroundColor Green

# 5. Tests
if ($SkipTests) {
  Write-Host "[5/5] Tests SKIPPED" -ForegroundColor Yellow
} else {
  Write-Host "[5/5] Running tests..." -ForegroundColor Yellow
  Push-Location $root; pnpm test; Pop-Location
  if ($LASTEXITCODE -ne 0) { Write-Host "FAIL: Tests failed" -ForegroundColor Red; exit 1 }
  Write-Host "  OK" -ForegroundColor Green
}

Write-Host ""
Write-Host "=== BUILD PASSED ===" -ForegroundColor Green
