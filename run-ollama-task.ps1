param(
  [Parameter(Mandatory=$true)]
  [string]$Task
)

cd "C:\Users\ntsif\OneDrive\Tsifulator.ai"
$targetFileMatch = [regex]::Match($Task, "([a-zA-Z0-9\/\._-]+\.ts)")

$targetFileContent = ""

if ($targetFileMatch.Success) {
  $targetPath = $targetFileMatch.Value
  if (Test-Path $targetPath) {
    $targetFileContent = Get-Content $targetPath -Raw
  }
}
$contextFiles = Get-ChildItem -Recurse -Include *.ts,*.js,*.json -File |
  Where-Object { $_.FullName -notmatch "node_modules|dist|build|\.git" } |
  Select-Object -First 8

$contextText = ""

foreach ($file in $contextFiles) {
  $path = $file.FullName.Replace((Get-Location).Path + "\", "")
  $content = [System.IO.File]::ReadAllText($file.FullName)

  if ($content.Length -gt 2500) {
    $content = $content.Substring(0, 2500)
  }

  $contextText += "`nFILE: $path`n$content`n"
}

$fullPrompt = @"
You are a coding agent.

Return ONLY valid JSON.

Do NOT include:
- explanations
- markdown
- backticks
- thinking

If no safe change:
return:
{"status":"no_safe_change","summary":"reason","files":[],"commands":[]}

Otherwise return:
{"status":"ok","summary":"...","files":[{"path":"...","content":"..."}],"commands":["npm run build","npm test"]}

TASK:
$Task
"@
$raw = $fullPrompt | ollama run qwen2.5-coder:7b
$raw | Set-Content .\ollama-last-answer-raw.txt -Encoding utf8

$rawText = [string]::Join("`n", ($raw | ForEach-Object { $_.ToString() }))

if ([string]::IsNullOrWhiteSpace($rawText)) {
  if ([string]::IsNullOrWhiteSpace($rawText)) {
  $fallback = '{"status":"no_safe_change","summary":"empty_output","files":[],"commands":[]}'
  Set-Content .\ollama-last-answer.json $fallback -Encoding utf8
  Write-Host "Fallback JSON written due to empty output" -ForegroundColor Yellow
  exit 0
}
}

$start = $rawText.IndexOf("{")
$end = $rawText.LastIndexOf("}")

if ($start -lt 0 -or $end -lt 0 -or $end -le $start) {
  throw "Ollama did not return parseable JSON."
}

$json = $rawText.Substring($start, $end - $start + 1)

try {
  $parsed = $json | ConvertFrom-Json
}
catch {
  $json | Set-Content .\ollama-invalid-json.txt -Encoding utf8
  throw "Ollama returned invalid JSON. Saved raw candidate to .\ollama-invalid-json.txt"
}
$json | Set-Content .\ollama-last-answer.json -Encoding utf8
$json | Set-Content .\ollama-last-answer.txt -Encoding utf8

Write-Host $json
