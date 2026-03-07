param(
  [string]$Action = "register",
  [string]$Time = "07:00",
  [string]$TaskName = "TsifulatorDailyBundle",
  [string]$ProjectDir = "",
  [string]$Email = "partner.a@company.com",
  [string]$PartnerEmailsCsv = "partner.a@company.com,partner.b@company.com",
  [string]$BaseUrl = "http://127.0.0.1:4000"
)

$ErrorActionPreference = "Stop"

Write-Host '== tsifulator.ai daily scheduler ==' -ForegroundColor Cyan

# ---------- resolve paths ----------
if (-not $ProjectDir) {
  $ProjectDir = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

$bundleScript = Join-Path $ProjectDir "scripts\beta-daily-bundle.ps1"
if (-not (Test-Path $bundleScript)) {
  throw "Bundle script not found: $bundleScript"
}

# ---------- log file ----------
$logDir = Join-Path $ProjectDir "docs\logs"
if (-not (Test-Path $logDir)) {
  New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$logFile = Join-Path $logDir "daily-bundle.log"

switch ($Action.ToLower()) {

  "register" {
    Write-Host "Task:      $TaskName"
    Write-Host "Time:      $Time daily"
    Write-Host "Project:   $ProjectDir"
    Write-Host "Log:       $logFile"
    Write-Host ''

    # Build the PowerShell command that will run inside the scheduled task
    $psCommand = @(
      "Set-Location -Path '$ProjectDir';",
      "`$date = Get-Date -Format 'yyyy-MM-dd';",
      "Start-Transcript -Path '$logFile' -Append;",
      "Write-Host `"[`$date] Daily bundle starting...`";",
      "try {",
      "  & powershell -ExecutionPolicy Bypass -File '$bundleScript'",
      "    -Email '$Email'",
      "    -PartnerEmailsCsv '$PartnerEmailsCsv'",
      "    -BaseUrl '$BaseUrl';",
      "  Write-Host `"[`$date] Daily bundle completed (exit: `$LASTEXITCODE)`";",
      "} catch {",
      "  Write-Host `"[`$date] Daily bundle FAILED: `$_`";",
      "}",
      "Stop-Transcript;"
    ) -join ' '

    $taskAction = New-ScheduledTaskAction `
      -Execute "powershell.exe" `
      -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command `"$psCommand`"" `
      -WorkingDirectory $ProjectDir

    $taskTrigger = New-ScheduledTaskTrigger -Daily -At $Time

    $taskSettings = New-ScheduledTaskSettingsSet `
      -AllowStartIfOnBatteries `
      -DontStopIfGoingOnBatteries `
      -StartWhenAvailable `
      -RunOnlyIfNetworkAvailable `
      -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

    # Check if task already exists
    $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($existing) {
      Write-Host '  [info] Task already exists, updating...' -ForegroundColor DarkYellow
      Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }

    Register-ScheduledTask `
      -TaskName $TaskName `
      -Action $taskAction `
      -Trigger $taskTrigger `
      -Settings $taskSettings `
      -Description "Tsifulator.ai daily bundle: checkpoint + comparison + trend + validation + report + alerts + feedback + digest + scorecard + trend + rotation" `
      | Out-Null

    Write-Host ('  [ok] Scheduled task registered: ' + $TaskName) -ForegroundColor Green
    Write-Host "  Runs daily at $Time"
    Write-Host ''

    # Show confirmation
    $task = Get-ScheduledTask -TaskName $TaskName
    $info = $task | Get-ScheduledTaskInfo
    Write-Host '  Task details:' -ForegroundColor White
    Write-Host "    State:    $($task.State)"
    Write-Host "    NextRun:  $($info.NextRunTime)"
    Write-Host ''
  }

  "unregister" {
    $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $existing) {
      Write-Host ('  [info] Task not found: ' + $TaskName) -ForegroundColor DarkYellow
      Write-Host '[done] Nothing to unregister.' -ForegroundColor Green
      exit 0
    }
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host ('  [ok] Task unregistered: ' + $TaskName) -ForegroundColor Green
    Write-Host ''
  }

  "status" {
    $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $existing) {
      Write-Host ('  [info] Task not registered: ' + $TaskName) -ForegroundColor DarkYellow
      exit 0
    }

    $info = $existing | Get-ScheduledTaskInfo

    Write-Host '  Task details:' -ForegroundColor White
    Write-Host "    Name:     $TaskName"
    Write-Host "    State:    $($existing.State)"
    Write-Host "    NextRun:  $($info.NextRunTime)"
    Write-Host "    LastRun:  $($info.LastRunTime)"
    Write-Host "    LastResult: $($info.LastTaskResult)"
    Write-Host ''

    # Show recent log tail
    if (Test-Path $logFile) {
      Write-Host '  Recent log (last 15 lines):' -ForegroundColor White
      Get-Content -Path $logFile -Tail 15 | ForEach-Object { Write-Host "    $_" }
      Write-Host ''
    }
    else {
      Write-Host '  [info] No log file yet.' -ForegroundColor DarkYellow
    }
  }

  "run" {
    Write-Host '  [info] Running bundle now (manual trigger)...' -ForegroundColor Yellow
    $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $existing) {
      Write-Host ('  [warn] Task not registered. Registering first...') -ForegroundColor DarkYellow
      & $PSCommandPath -Action register -Time $Time -TaskName $TaskName -ProjectDir $ProjectDir -Email $Email -PartnerEmailsCsv $PartnerEmailsCsv -BaseUrl $BaseUrl
    }
    Start-ScheduledTask -TaskName $TaskName
    Write-Host ('  [ok] Task triggered: ' + $TaskName) -ForegroundColor Green
    Write-Host '  Check status with: npm run beta:schedule:status'
    Write-Host ''
  }

  "log" {
    if (-not (Test-Path $logFile)) {
      Write-Host '  [info] No log file yet.' -ForegroundColor DarkYellow
      exit 0
    }
    Write-Host '  Log file:' -ForegroundColor White
    Write-Host "    $logFile"
    Write-Host ''
    Get-Content -Path $logFile -Tail 50
  }

  default {
    Write-Host '  [error] Unknown action. Use: register, unregister, status, run, log' -ForegroundColor Red
    exit 1
  }
}

Write-Host '[done] Daily scheduler complete.' -ForegroundColor Green
