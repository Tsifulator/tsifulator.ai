param(
  [string]$Prompt
)

$body = @{
  model = "qwen2.5-coder:7b"
  messages = @(
    @{
      role = "user"
      content = $Prompt
    }
  )
  stream = $false
} | ConvertTo-Json -Depth 6

$response = Invoke-RestMethod -Uri "http://localhost:11434/api/chat" -Method Post -ContentType "application/json" -Body $body
$response.message.content
