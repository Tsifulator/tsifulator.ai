param(
  [string]$TaskFile = ".\current-task.txt",
  [string]$ContextFile = ".\current-context.txt",
  [string]$PlanFile = ".\ollama-last-plan.txt",
  [string]$ClaudeFile = ".\claude-last-answer.json",
  [string]$OllamaModel = "llama3.1",
  [string]$ClaudeModel = "claude-sonnet-4-5",
  [switch]$Deep,
  [switch]$Apply,
  [switch]$RunChecks
)

$ErrorActionPreference = "Stop"

if ($Deep) {
  $ClaudeModel = "claude-sonnet-4-5"
}

if (-not $env:ANTHROPIC_API_KEY) {
  throw "ANTHROPIC_API_KEY is not set."
}

function Clean-Text([string]$Text) {
  if ($null -eq $Text) { return "" }

  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $Text.ToCharArray()) {
    $code = [int][char]$ch
    if (
      ($code -eq 9) -or
      ($code -eq 10) -or
      ($code -eq 13) -or
      ($code -ge 32 -and $code -le 55295) -or
      ($code -ge 57344 -and $code -le 65533)
    ) {
      [void]$sb.Append($ch)
    }
    else {
      [void]$sb.Append("?")
    }
  }
  return $sb.ToString()
}

function Extract-JsonObject([string]$Text) {
  $parsedText = (Clean-Text $Text).Trim()

  if ([string]::IsNullOrWhiteSpace($parsedText)) {
    throw "Claude returned empty text."
  }

  if ($parsedText.StartsWith('`json')) {
    $parsedText = $parsedText.Substring(7).Trim()
  }

  if ($parsedText.StartsWith('`')) {
    $parsedText = $parsedText.Substring(3).Trim()
  }

  if ($parsedText.EndsWith('`')) {
    $parsedText = $parsedText.Substring(0, $parsedText.Length - 3).Trim()
  }

  $start = $parsedText.IndexOf("{")
  $end = $parsedText.LastIndexOf("}")

  if ($start -lt 0 -or $end -lt 0 -or $end -lt $start) {
    throw "Claude did not return a parseable JSON object."
  }

  return $parsedText.Substring($start, $end - $start + 1)
}

if (-not (Test-Path $TaskFile)) {
  throw "Task file not found: $TaskFile"
}

$task = Clean-Text (Get-Content $TaskFile -Raw)
if ([string]::IsNullOrWhiteSpace($task)) {
  throw "Task file is empty: $TaskFile"
}

$context = ""
if (Test-Path $ContextFile) {
  $context = Clean-Text (Get-Content $ContextFile -Raw)
}

Write-Host ""
Write-Host "=== STEP 1: OLLAMA PLAN ===" -ForegroundColor Cyan

$plannerPrompt = @"
You are preparing a very small coding task for Claude.

Return exactly these sections:
TITLE
GOAL
FILES TO CHANGE
RULES
DEFINITION OF DONE

Requirements:
- Keep it surgical.
- 1 to 2 files max.
- No invented files.
- No feature creep.
- Only use files present in the provided context.
- If unclear, say NO SAFE BLOCK FOUND.
"@

$ollamaInput = Clean-Text ($plannerPrompt + "`n`nTASK:`n" + $task + "`n`nCODE CONTEXT:`n" + $context)
$plan = $ollamaInput | ollama run $OllamaModel
$plan = Clean-Text $plan
$plan | Set-Content $PlanFile -Encoding utf8

Write-Host ""
Write-Host "=== STEP 2: CLAUDE CODEGEN ===" -ForegroundColor Cyan

$systemPrompt = @"
You are an expert TypeScript coding assistant editing an existing repository.

Rules:
- Make the smallest safe change only.
- Use only files actually present in the provided context.
- Do not invent files, functions, or abstractions.
- Return JSON only.
- No markdown fences.
- No explanations.
- If no safe grounded change exists, return exactly:
{"status":"no_safe_block_found","files":[],"commands":[]}

Otherwise return exactly this schema:
{
  "status": "ok",
  "files": [
    {
      "path": "relative/path/to/file.ts",
      "content": "FULL FINAL FILE CONTENT HERE"
    }
  ],
  "commands": [
    "npm run build",
    "npm test"
  ]
}
"@

$userPrompt = Clean-Text @"
TASK
$task

OLLAMA PLAN
$plan

CODE CONTEXT
$context
"@

$bodyObject = @{
  model = $ClaudeModel
  max_tokens = 4000
  system = (Clean-Text $systemPrompt)
  messages = @(
    @{
      role = "user"
      content = $userPrompt
    }
  )
}

$body = $bodyObject | ConvertTo-Json -Depth 10 -Compress
$body = Clean-Text $body
$body | Set-Content ".\claude-request-body.json" -Encoding utf8

try {
  $utf8 = New-Object System.Text.UTF8Encoding($false)
  $bytes = $utf8.GetBytes($body)

  $response = Invoke-RestMethod `
    -Method Post `
    -Uri "https://api.anthropic.com/v1/messages" `
    -Headers @{
      "x-api-key" = $env:ANTHROPIC_API_KEY
      "anthropic-version" = "2023-06-01"
    } `
    -ContentType "application/json; charset=utf-8" `
    -Body $bytes

  $textParts = @()
  foreach ($item in $response.content) {
    if ($item.type -eq "text") {
      $textParts += $item.text
    }
  }

  $finalText = Clean-Text (($textParts -join "`r`n").Trim())
  $finalText | Set-Content ".\claude-last-answer-raw.txt" -Encoding utf8

  $jsonText = Extract-JsonObject $finalText
  $jsonText | Set-Content ".\claude-last-answer-parsed.json" -Encoding utf8
  $jsonText | Set-Content $ClaudeFile -Encoding utf8

  Write-Host "Saved Claude raw text to: .\claude-last-answer-raw.txt" -ForegroundColor Yellow
  Write-Host "Saved Claude parsed JSON to: $ClaudeFile" -ForegroundColor Yellow

  $json = $jsonText | ConvertFrom-Json

  if ($json.status -eq "no_safe_block_found") {
    Write-Host ""
    Write-Host "Claude reported: NO SAFE BLOCK FOUND" -ForegroundColor Yellow
    exit 0
  }

  if ($Apply) {
    Write-Host ""
    Write-Host "=== STEP 3: APPLY FILES ===" -ForegroundColor Cyan

    foreach ($file in $json.files) {
      $path = $file.path
      $dir = Split-Path $path -Parent
      if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
      }
      (Clean-Text $file.content) | Set-Content $path -Encoding utf8
      Write-Host "Wrote $path" -ForegroundColor Green
    }
  }

  if ($RunChecks) {
    Write-Host ""
    Write-Host "=== STEP 4: RUN CHECKS ===" -ForegroundColor Cyan
    foreach ($cmd in $json.commands) {
      Write-Host "Running: $cmd" -ForegroundColor Yellow
      Invoke-Expression $cmd
      if ($LASTEXITCODE -ne 0) {
        throw "Command failed: $cmd"
      }
    }
  }

  Write-Host ""
  Write-Host "=== DONE ===" -ForegroundColor Green
  Write-Host "Review changes with: git diff -- .\server\src"
}
catch {
  Write-Host ""
  Write-Host "=== ERROR ===" -ForegroundColor Red
  Write-Host $_.Exception.Message -ForegroundColor Red

  if ($_.Exception.Response) {
    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $errorBody = $reader.ReadToEnd()
    $errorBody | Set-Content ".\claude-error.txt" -Encoding utf8
    Write-Host "Saved API error to .\claude-error.txt" -ForegroundColor Yellow
    Write-Host $errorBody
  }

  throw
}

