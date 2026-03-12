param(
[string]$PromptFile
)

if (-not (Test-Path $PromptFile)) {
Write-Error "Prompt file missing."
exit 1
}

$prompt = Get-Content $PromptFile -Raw

$payload = @{
model = "qwen2.5-coder:7b"
messages = @(
@{
role = "user"
content = $prompt
}
)
stream = $false
}

$json = $payload | ConvertTo-Json -Depth 20 -Compress

try {
$response = Invoke-WebRequest `
-Uri "http://localhost:11434/api/chat" `
-Method Post `
-ContentType "application/json" `
-Body $json

$body = $response.Content | ConvertFrom-Json
$body.message.content
}
catch {
Write-Error "Ollama request failed:"
Write-Error $_
exit 1
}


