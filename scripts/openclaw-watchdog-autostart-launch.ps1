Set-Location -LiteralPath 'C:\Users\ntsif\OneDrive\Tsifulator.ai'
& 'C:\Users\ntsif\OneDrive\Tsifulator.ai\scripts\openclaw-watchdog-start.ps1' -PromptFile '.\docs\openclaw-cross-app-mvp-prompt.md' -SessionId 'tsif-phase1' -AgentId 'main' -IntervalSeconds 180 -LockStaleMinutes 2
