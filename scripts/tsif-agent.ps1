param(
  [string]$Model = "kimi-k2.5:cloud",
  [string]$TaskFile = ".\current-task.txt"
)

$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path $PSScriptRoot -Parent

if (-not (Test-Path $TaskFile)) {
  throw "Task file not found: $TaskFile"
}

$task = Get-Content $TaskFile -Raw
if ([string]::IsNullOrWhiteSpace($task)) {
  throw "Task file is empty: $TaskFile"
}

function Read-RepoFile([string]$RelativePath) {
  $fullPath = Join-Path $RepoRoot $RelativePath
  if (-not (Test-Path $fullPath)) {
    return ""
  }

  $content = Get-Content $fullPath -Raw
  return "=== $RelativePath ===`r`n$content`r`n"
}

function Extract-JsonObject([string]$Text) {
  $parsedText = $Text.Trim()

  if ([string]::IsNullOrWhiteSpace($parsedText)) {
    throw "Model returned empty text."
  }

  if ($parsedText.StartsWith('```json')) {
    $parsedText = $parsedText.Substring(7).Trim()
  }

  if ($parsedText.StartsWith('```')) {
    $parsedText = $parsedText.Substring(3).Trim()
  }

  if ($parsedText.EndsWith('```')) {
    $parsedText = $parsedText.Substring(0, $parsedText.Length - 3).Trim()
  }

  $start = $parsedText.IndexOf("{")
  $end = $parsedText.LastIndexOf("}")

  if ($start -lt 0 -or $end -lt 0 -or $end -lt $start) {
    throw "Model did not return a parseable JSON object."
  }

  return $parsedText.Substring($start, $end - $start + 1)
}

$contextFiles = @(
  "server/src/adapters/terminal-adapter.ts",
  "server/src/adapters/contract.ts",
  "server/src/risk.ts",
  "server/src/types.ts",
  "server/src/shared-types.ts",
  "server/src/chat-engine.ts"
)

$context = ($contextFiles | ForEach-Object { Read-RepoFile $_ }) -join "`r`n"

$prompt = @"
You are editing an existing TypeScript codebase for Tsifulator.ai.

Return JSON only. No markdown. No explanation.

Task:
$task

Hard constraints:
- edit only ONE small file
- terminal only
- no Excel changes
- no RStudio changes
- no new features
- no broad rewrites
- do not edit db.ts
- preserve behavior exactly
- build and tests must still pass

Allowed files:
- server/src/adapters/terminal-adapter.ts
- server/src/adapters/contract.ts
- server/src/risk.ts
- server/src/types.ts
- server/src/shared-types.ts
- server/src/chat-engine.ts

JSON schema:
{
  "status": "ok" | "no_safe_block_found",
  "files": [
    {
      "path": "relative/path",
      "content": "full final file content"
    }
  ]
}

Rules:
- If no safe change exists, return {"status":"no_safe_block_found","files":[]}
- If status is ok, files must contain exactly one file
- content must be the full final content of that file
- do not include any text outside the JSON object

Code context:
$context
"@

Write-Host ""
Write-Host "=== STEP 1: MODEL CODEGEN ==="
Write-Host ""

$rawText = $prompt | & ollama run $Model | Out-String
$rawPath = Join-Path $RepoRoot "ollama-last-answer-raw.txt"
$rawText | Set-Content $rawPath -Encoding utf8
Write-Host "Saved raw model output to: $rawPath"

$jsonText = Extract-JsonObject $rawText
$jsonPath = Join-Path $RepoRoot "ollama-last-answer.json"
$jsonText | Set-Content $jsonPath -Encoding utf8
Write-Host "Saved parsed JSON to: $jsonPath"

$json = $jsonText | ConvertFrom-Json

if ($json.status -eq "no_safe_block_found") {
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.MessageBox]::Show("No safe block found.","Tsifulator Agent")
  Write-Host ""
  Write-Host "=== DONE: NO SAFE BLOCK FOUND ==="
  exit 0
}

if ($json.status -ne "ok") {
  throw "Unexpected status: $($json.status)"
}

if (-not $json.files -or $json.files.Count -ne 1) {
  throw "Expected exactly one file in JSON response."
}

Write-Host ""
Write-Host "=== STEP 2: APPLY FILE ==="
Write-Host ""

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

foreach ($file in $json.files) {
  $targetPath = Join-Path $RepoRoot $file.path
  $targetDir = Split-Path $targetPath -Parent

  if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
  }

  [System.IO.File]::WriteAllText($targetPath, [string]$file.content, $utf8NoBom)
  Write-Host "Wrote $($file.path)"
}

Write-Host ""
Write-Host "=== STEP 3: RUN CHECKS ==="
Write-Host ""

Push-Location $RepoRoot
try {
  Write-Host "Running: npm run build"
  npm run build
  if ($LASTEXITCODE -ne 0) {
    throw "Build failed."
  }

  Write-Host "Running: npm test"
  npm test
  if ($LASTEXITCODE -ne 0) {
    throw "Tests failed."
  }
}
finally {
  Pop-Location
}

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("Coding block finished successfully.","Tsifulator Agent")

Write-Host ""
Write-Host "=== DONE ==="
Write-Host "Review changes with: git --no-pager diff -- .\server\src"
