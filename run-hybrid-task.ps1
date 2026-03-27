param(
  [Parameter(Mandatory=$true)]
  [string]$Task
)

cd "C:\Users\ntsif\OneDrive\Tsifulator.ai"
Set-Content .\current-task.txt $Task -Encoding utf8
powershell -ExecutionPolicy Bypass -File .\run-tsif-agent-once.ps1
