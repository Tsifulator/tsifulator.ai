cd "C:\Users\ntsif\OneDrive\Tsifulator.ai"
powershell -ExecutionPolicy Bypass -File .\run-ollama-task.ps1 -Task (Get-Content .\current-task.txt -Raw)
