param(
  [string]$PromptFile = ".\docs\openclaw-autopilot-prompt.md",
  [string]$AgentId = "main",
  [string]$SessionId = "tsif-phase1",
  [int]$MaxTurns = 120,
  [int]$AgentTimeoutSeconds = 120,
  [int]$TurnDelaySeconds = 20,
  [int]$RateLimitCooldownSeconds = 180,
  [int]$QueueWaitTimeoutSeconds = 600,
  [int]$MaxPromptChars = 8000,
  [switch]$CompactResumePrompt,
  [switch]$DryRun
)

$ErrorActionPreference = "Continue"
$PSNativeCommandUseErrorActionPreference = $false

$openclawCmd = Get-Command openclaw.cmd -ErrorAction SilentlyContinue
if (-not $openclawCmd) {
  Write-Host "[fail] openclaw.cmd not found on PATH" -ForegroundColor Red
  exit 1
}
$openclawExe = $openclawCmd.Source

$runDir = Join-Path (Get-Location) "docs\phase-reports\autopilot-runs"
if (-not (Test-Path $runDir)) {
  New-Item -ItemType Directory -Path $runDir | Out-Null
}

$runnerLockPath = Join-Path $runDir "agent-run-slot.lock"
$lockAcquired = $false

function Acquire-RunSlot {
  param(
    [int]$TimeoutSeconds
  )

  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  while ((Get-Date) -lt $deadline) {
    if (-not (Test-Path $runnerLockPath)) {
      @{ pid = $PID; startedAt = (Get-Date).ToString("o") } | ConvertTo-Json | Out-File -FilePath $runnerLockPath -Encoding utf8
      return $true
    }

    try {
      $existing = Get-Content $runnerLockPath -Raw | ConvertFrom-Json
      if ($existing.pid) {
        $proc = Get-Process -Id $existing.pid -ErrorAction SilentlyContinue
        if (-not $proc) {
          Remove-Item $runnerLockPath -Force -ErrorAction SilentlyContinue
          continue
        }
      }
    } catch {
      Remove-Item $runnerLockPath -Force -ErrorAction SilentlyContinue
      continue
    }

    Start-Sleep -Seconds (Get-Random -Minimum 2 -Maximum 6)
  }

  return $false
}

function Release-RunSlot {
  if (-not (Test-Path $runnerLockPath)) {
    return
  }

  try {
    $existing = Get-Content $runnerLockPath -Raw | ConvertFrom-Json
    if ($existing.pid -eq $PID) {
      Remove-Item $runnerLockPath -Force -ErrorAction SilentlyContinue
    }
  } catch {
    Remove-Item $runnerLockPath -Force -ErrorAction SilentlyContinue
  }
}

function Clear-StaleSessionLock {
  $sessionsDir = Join-Path $env:USERPROFILE ".openclaw\agents\$AgentId\sessions"
  $lockPath = Join-Path $sessionsDir "$SessionId.jsonl.lock"

  if (-not (Test-Path $lockPath)) {
    return
  }

  $raw = Get-Content $lockPath -Raw -ErrorAction SilentlyContinue
  $pidMatch = [regex]::Match($raw, "(\d+)")

  if ($pidMatch.Success) {
    $lockPid = [int]$pidMatch.Groups[1].Value
    $proc = Get-Process -Id $lockPid -ErrorAction SilentlyContinue
    if (-not $proc) {
      Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
      Write-Host "[warn] Removed stale lock file for dead PID $lockPid" -ForegroundColor DarkYellow
    }
  } else {
    Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
    Write-Host "[warn] Removed unreadable lock file: $lockPath" -ForegroundColor DarkYellow
  }
}

if (-not (Test-Path $PromptFile)) {
  Write-Host "[fail] Prompt file not found: $PromptFile" -ForegroundColor Red
  exit 1
}

if (-not (Acquire-RunSlot -TimeoutSeconds $QueueWaitTimeoutSeconds)) {
  Write-Host "[fail] Timed out waiting for run slot after $QueueWaitTimeoutSeconds seconds." -ForegroundColor Red
  exit 1
}
$lockAcquired = $true

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$logPath = Join-Path $runDir "run-$timestamp.log"

$initialPromptRaw = Get-Content $PromptFile -Raw
$initialPrompt = if ($initialPromptRaw.Length -gt $MaxPromptChars) {
  $initialPromptRaw.Substring(0, $MaxPromptChars)
} else {
  $initialPromptRaw
}

$sessionFilePath = Join-Path $env:USERPROFILE ".openclaw\agents\$AgentId\sessions\$SessionId.jsonl"
$shouldUseCompactPrompt = $CompactResumePrompt -or (Test-Path $sessionFilePath)

$firstMessage = if ($shouldUseCompactPrompt) {
  "Resume from current session checkpoint. Emit only next smallest code delta (max 2 files), then run validation. Keep response compact."
} else {
  $initialPrompt
}

$messages = @(
  $firstMessage,
  "Continue from the latest checkpoint. Output next concrete implementation step only.",
  "Continue. Keep strict format and avoid prose.",
  "Continue. Run validation and fix failures automatically.",
  "Continue to next remaining task."
)

$consecutiveRateLimitHits = 0
$consecutiveNoReplyHits = 0
$maxBackoffSeconds = 900

function Invoke-AgentTurn {
  param(
    [string]$Message
  )

  if ($DryRun) {
    return "[dry-run] $Message"
  }

  $output = & $openclawExe agent --local --agent $AgentId --session-id $SessionId --timeout $AgentTimeoutSeconds --message $Message 2>&1
  $exitCode = $LASTEXITCODE
  return "EXIT_CODE=$exitCode`n" + ($output | Out-String)
}

"=== OpenClaw Autopilot Run ===" | Out-File -FilePath $logPath -Encoding utf8
"Started: $(Get-Date -Format o)" | Out-File -FilePath $logPath -Append -Encoding utf8
"PromptFile: $PromptFile" | Out-File -FilePath $logPath -Append -Encoding utf8
"AgentId: $AgentId" | Out-File -FilePath $logPath -Append -Encoding utf8
"SessionId: $SessionId" | Out-File -FilePath $logPath -Append -Encoding utf8

Write-Host "[info] Log file: $logPath" -ForegroundColor Cyan
Write-Host "[info] Starting autopilot with session '$SessionId'" -ForegroundColor Cyan

if ($shouldUseCompactPrompt) {
  Write-Host "[info] Using compact resume prompt to reduce TPM usage." -ForegroundColor Cyan
}

try {
  for ($turn = 1; $turn -le $MaxTurns; $turn++) {
    $message = if ($turn -le $messages.Count) { $messages[$turn - 1] } else { $messages[$messages.Count - 1] }

    Clear-StaleSessionLock

    Write-Host "[turn $turn/$MaxTurns] Sending message..." -ForegroundColor Yellow
    "\n--- TURN $turn ---" | Out-File -FilePath $logPath -Append -Encoding utf8
    "MESSAGE:" | Out-File -FilePath $logPath -Append -Encoding utf8
    $message | Out-File -FilePath $logPath -Append -Encoding utf8

    $result = Invoke-AgentTurn -Message $message
    $resultTrimmed = $result.Trim()

    "RESPONSE:" | Out-File -FilePath $logPath -Append -Encoding utf8
    $resultTrimmed | Out-File -FilePath $logPath -Append -Encoding utf8

    if ($resultTrimmed -match "API rate limit reached|rate limit|\b429\b") {
      $consecutiveRateLimitHits++
      $consecutiveNoReplyHits = 0
      $backoffMultiplier = [Math]::Pow(2, [Math]::Min($consecutiveRateLimitHits - 1, 5))
      $cooldown = [Math]::Min($maxBackoffSeconds, [int]($RateLimitCooldownSeconds * $backoffMultiplier))
      $jitterLow = [Math]::Max(5, [int]($cooldown * 0.10))
      $jitterHigh = [Math]::Max($jitterLow + 1, [int]($cooldown * 0.35))
      $jitter = Get-Random -Minimum $jitterLow -Maximum $jitterHigh
      $sleepFor = [Math]::Min($maxBackoffSeconds, $cooldown + $jitter)
      Write-Host "[warn] Rate limit/429 hit x$consecutiveRateLimitHits. Cooling down for $sleepFor seconds..." -ForegroundColor DarkYellow
      Start-Sleep -Seconds $sleepFor
      continue
    }

    $consecutiveRateLimitHits = 0

    if ($resultTrimmed -match "session file locked") {
      $lockPidMatch = [regex]::Match($resultTrimmed, "pid=(\d+)")
      if ($lockPidMatch.Success) {
        $lockPid = [int]$lockPidMatch.Groups[1].Value
        if ($lockPid -ne $PID) {
          $lockProc = Get-Process -Id $lockPid -ErrorAction SilentlyContinue
          if ($lockProc) {
            try {
              Stop-Process -Id $lockPid -Force -ErrorAction Stop
              Write-Host "[warn] Stopped lock-holder PID $lockPid ($($lockProc.ProcessName))." -ForegroundColor DarkYellow
            } catch {
              Write-Host "[warn] Could not stop lock-holder PID $lockPid." -ForegroundColor DarkYellow
            }
          }
        }
      }

      $lockPathMatch = [regex]::Match($resultTrimmed, '([A-Za-z]:\\[^\r\n"]+\.jsonl\.lock)')
      if ($lockPathMatch.Success) {
        $lockPath = $lockPathMatch.Groups[1].Value
        if (Test-Path $lockPath) {
          Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
          Write-Host "[warn] Removed session lock file: $lockPath" -ForegroundColor DarkYellow
        }
      } else {
        Clear-StaleSessionLock
      }

      Start-Sleep -Seconds 15
      continue
    }

    if ($resultTrimmed -match "No reply from agent") {
      $consecutiveNoReplyHits++
      $backoffMultiplier = [Math]::Pow(2, [Math]::Min($consecutiveNoReplyHits - 1, 4))
      $baseDelay = [Math]::Max(45, [int]($RateLimitCooldownSeconds * 0.75))
      $retryDelay = [Math]::Min($maxBackoffSeconds, [int]($baseDelay * $backoffMultiplier) + (Get-Random -Minimum 4 -Maximum 20))
      Write-Host "[warn] No reply from agent x$consecutiveNoReplyHits. Cooling down for $retryDelay seconds..." -ForegroundColor DarkYellow
      Start-Sleep -Seconds $retryDelay
      continue
    }

    $consecutiveNoReplyHits = 0

    if ($resultTrimmed -match "OFF_FORMAT") {
      Write-Host "[warn] OFF_FORMAT detected, sending format reset next turn." -ForegroundColor DarkYellow
      $messages = @(
        "Regenerate now. Output ONLY SECTION 2: COMMANDS for the next implementation step. No commentary.",
        "Now output ONLY SECTION 3: FILES for next 2 files using FILE: <path> ... END_FILE.",
        "Continue from latest checkpoint."
      )
    }

    if ((-not $DryRun) -and ($resultTrimmed -match "PHASE_COMPLETE|ACCEPTANCE_STATUS:\s*PASS|ALL_CHECKS_PASS")) {
      Write-Host "[ok] Completion signal detected. Stopping autopilot." -ForegroundColor Green
      break
    }

    Start-Sleep -Seconds ($TurnDelaySeconds + (Get-Random -Minimum 1 -Maximum 6))
  }
}
finally {
  if ($lockAcquired) {
    Release-RunSlot
  }
}

"Finished: $(Get-Date -Format o)" | Out-File -FilePath $logPath -Append -Encoding utf8
Write-Host "[done] Autopilot finished. Review: $logPath" -ForegroundColor Green
