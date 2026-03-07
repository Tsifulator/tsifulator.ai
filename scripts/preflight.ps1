$ErrorActionPreference = "Stop"

Write-Host "== tsifulator.ai preflight ==" -ForegroundColor Cyan

function Check-Command {
  param(
    [string]$Name,
    [string]$VersionArg = "--version"
  )

  if (Get-Command $Name -ErrorAction SilentlyContinue) {
    try {
      $version = & $Name $VersionArg 2>$null
      Write-Host "[ok] $Name $version" -ForegroundColor Green
      return $true
    }
    catch {
      Write-Host "[ok] $Name (installed)" -ForegroundColor Green
      return $true
    }
  }

  Write-Host "[missing] $Name" -ForegroundColor Yellow
  return $false
}

$allGood = $true
$allGood = (Check-Command "git") -and $allGood
$allGood = (Check-Command "node" "-v") -and $allGood
$allGood = (Check-Command "npm" "-v") -and $allGood
$allGood = (Check-Command "openclaw" "--version") -and $allGood
$allGood = (Check-Command "python" "--version") -and $allGood

$requiredDirs = @(
  "server",
  "clients/terminal",
  "docs"
)

foreach ($dir in $requiredDirs) {
  if (-not (Test-Path $dir)) {
    New-Item -ItemType Directory -Path $dir | Out-Null
    Write-Host "[created] $dir" -ForegroundColor DarkYellow
  } else {
    Write-Host "[ok] $dir" -ForegroundColor Green
  }
}

if (-not (Test-Path ".env")) {
  Copy-Item ".env.example" ".env"
  Write-Host "[created] .env from .env.example" -ForegroundColor DarkYellow
} else {
  Write-Host "[ok] .env" -ForegroundColor Green
}

if (-not (Test-Path ".git")) {
  git init | Out-Null
  Write-Host "[created] git repository" -ForegroundColor DarkYellow
} else {
  Write-Host "[ok] git repository" -ForegroundColor Green
}

Write-Host "\nExtension check:" -ForegroundColor Cyan
$extensions = @(
  "github.copilot-chat",
  "dbaeumer.vscode-eslint",
  "esbenp.prettier-vscode"
)

$installed = ""
try {
  $installed = code --list-extensions
} catch {
  Write-Host "[warn] VS Code CLI 'code' not on PATH. Skip extension verification." -ForegroundColor Yellow
}

if ($installed) {
  foreach ($ext in $extensions) {
    if ($installed -match [regex]::Escape($ext)) {
      Write-Host "[ok] $ext" -ForegroundColor Green
    } else {
      Write-Host "[missing] $ext" -ForegroundColor Yellow
    }
  }
}

Write-Host "\nReady to run OpenClaw prompt files in docs/." -ForegroundColor Cyan
if (-not $allGood) {
  Write-Host "One or more tools are missing. Install missing tools and re-run preflight." -ForegroundColor Yellow
  exit 1
}
