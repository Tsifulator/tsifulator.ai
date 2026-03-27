param(
  [string]$JsonFile = ".\ollama-last-answer.json"
)

cd "C:\Users\ntsif\OneDrive\Tsifulator.ai"

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$resultFile = ".\agent-last-result.txt"

function Write-Result([string]$Text) {
  [System.IO.File]::WriteAllText((Resolve-Path $resultFile), $Text, $utf8NoBom)
}

function Run-CheckedCommand([string]$Command) {
  Write-Host "Running: $Command" -ForegroundColor Cyan
  cmd /c $Command
  if ($LASTEXITCODE -ne 0) {
    throw "Command failed: $Command"
  }
}

try {
  if (-not (Test-Path $JsonFile)) {
    throw "JSON file not found: $JsonFile"
  }

  $json = Get-Content $JsonFile -Raw | ConvertFrom-Json

  if ($json.status -eq "no_safe_change") {
    $msg = "NO_SAFE_CHANGE`n`nSummary:`n$($json.summary)"
    Write-Result $msg
    Write-Host $msg -ForegroundColor Yellow
    exit 0
  }

  if ($json.status -ne "ok") {
    throw "Model status was not ok."
  }

  foreach ($file in $json.files) {
    $relativePath = [string]$file.path
    $content = [string]$file.content

    if ([string]::IsNullOrWhiteSpace($relativePath)) {
      throw "A file entry had an empty path."
    }

    if ($relativePath.Contains("..") -or [System.IO.Path]::IsPathRooted($relativePath)) {
      throw "Unsafe file path: $relativePath"
    }

    $fullPath = Join-Path (Get-Location) $relativePath
    $parent = Split-Path $fullPath -Parent

    if (-not (Test-Path $parent)) {
      New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    [System.IO.File]::WriteAllText($fullPath, $content, $utf8NoBom)
    Write-Host "Wrote $relativePath" -ForegroundColor Green
  }

  $commands = @()
  foreach ($cmd in $json.commands) {
    if (-not [string]::IsNullOrWhiteSpace([string]$cmd)) {
      $commands += [string]$cmd
    }
  }

  if ($commands.Count -eq 0) {
    $commands = @("npm run build", "npm test")
  }

  foreach ($cmd in $commands) {
    Run-CheckedCommand $cmd
  }

  $changedFiles = @()
  foreach ($file in $json.files) {
    $changedFiles += [string]$file.path
  }

  $result = @"
SUCCESS

Summary:
$($json.summary)

Files changed:
$($changedFiles -join "`n")

Checks passed:
$($commands -join "`n")
"@

  Write-Result $result
  Write-Host $result -ForegroundColor Green
  exit 0
}
catch {
  $msg = @"
FAILED

Error:
$($_.Exception.Message)
"@
  Write-Result $msg
  Write-Host $msg -ForegroundColor Red
  exit 1
}
