param(
  [string]$PromptFile = ".\current-task.txt"
)

$ErrorActionPreference = "Stop"

if (-not $env:ANTHROPIC_API_KEY) {
  throw "ANTHROPIC_API_KEY is not set."
}

if (-not (Test-Path $PromptFile)) {
  throw "Prompt file not found: $PromptFile"
}

$task = Get-Content $PromptFile -Raw
if ([string]::IsNullOrWhiteSpace($task)) {
  throw "Prompt file is empty: $PromptFile"
}

@"
=== server/src/adapters/terminal-adapter.ts ===
$(Get-Content .\server\src\adapters\terminal-adapter.ts -Raw)

=== server/src/chat-engine.ts ===
$(Get-Content .\server\src\chat-engine.ts -Raw)

=== server/src/adapters/contract.ts ===
$(Get-Content .\server\src\adapters\contract.ts -Raw)

=== server/src/shared-types.ts ===
$(Get-Content .\server\src\shared-types.ts -Raw)

=== server/src/risk.ts ===
$(Get-Content .\server\src\risk.ts -Raw)

=== server/src/types.ts ===
$(Get-Content .\server\src\types.ts -Raw)

=== server/src/db.ts ===
$(Get-Content .\server\src\db.ts -Raw)
"@ | Set-Content .\current-context.txt -Encoding utf8

powershell -ExecutionPolicy Bypass -File .\scripts\clawdbot.ps1 -Deep -Apply -RunChecks

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("Claude finished coding this block.","Clawdbot")
