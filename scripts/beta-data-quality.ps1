param(
  [string]$DbPath = "./data/tsifulator.db",
  [string]$Date,
  [switch]$Json,
  [switch]$AppendToTriage,
  [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"

Write-Host '== tsifulator.ai data quality audit ==' -ForegroundColor Cyan
Write-Host "Date: $today"
Write-Host "DB: $DbPath"

if (-not (Test-Path $DbPath)) {
  Write-Host '  [error] Database not found' -ForegroundColor Red
  exit 1
}

# Helper: run sqlite query via node + better-sqlite3
function Invoke-Sql {
  param([string]$Query, [switch]$All)
  $dbNorm = ($DbPath -replace '\\','/')
  $tempJs = Join-Path (Split-Path $PSScriptRoot -Parent) "_dq_tmp_$([guid]::NewGuid().ToString('N').Substring(0,8)).js"
  $safeQuery = $Query -replace '"','\"'
  if ($All) {
    $lines = @(
      "const Database = require('better-sqlite3');"
      "const db = new Database('$dbNorm');"
      "try { const r = db.prepare(`"$safeQuery`").all(); console.log(JSON.stringify(r)); }"
      "catch(e) { console.log('[]'); }"
      "finally { db.close(); }"
    )
  } else {
    $lines = @(
      "const Database = require('better-sqlite3');"
      "const db = new Database('$dbNorm');"
      "try { const r = db.prepare(`"$safeQuery`").get(); console.log(JSON.stringify(r || {})); }"
      "catch(e) { console.log('{}'); }"
      "finally { db.close(); }"
    )
  }
  Set-Content -Path $tempJs -Value ($lines -join "`n") -Encoding UTF8
  $raw = node $tempJs 2>$null
  Remove-Item $tempJs -Force -ErrorAction SilentlyContinue
  if ($raw) { return ($raw | ConvertFrom-Json) } else { return $null }
}

$issues = @()
$checks = @()

# ---------- 1. Table row counts ----------
$tables = @('users', 'sessions', 'messages', 'action_proposals', 'approvals', 'action_executions', 'event_log', 'adapter_states')
$tableCounts = [ordered]@{}
foreach ($t in $tables) {
  $r = Invoke-Sql -Query "SELECT COUNT(*) as cnt FROM $t"
  $tableCounts[$t] = if ($r -and $r.cnt) { [int]$r.cnt } else { 0 }
}
$checks += [ordered]@{ name = 'table_counts'; status = 'info'; detail = $tableCounts }

# ---------- 2. Orphaned sessions (session with no user) ----------
$orphanedSessions = Invoke-Sql -Query "SELECT COUNT(*) as cnt FROM sessions WHERE user_id NOT IN (SELECT id FROM users)" 
$orphSessCount = if ($orphanedSessions -and $orphanedSessions.cnt) { [int]$orphanedSessions.cnt } else { 0 }
$ok = $orphSessCount -eq 0
$checks += [ordered]@{ name = 'orphaned_sessions'; status = if ($ok) { 'pass' } else { 'fail' }; cnt = $orphSessCount }
if (-not $ok) { $issues += "orphaned_sessions: $orphSessCount sessions have no matching user" }

# ---------- 3. Orphaned messages (message with no session) ----------
$orphanedMsgs = Invoke-Sql -Query "SELECT COUNT(*) as cnt FROM messages WHERE session_id NOT IN (SELECT id FROM sessions)"
$orphMsgCount = if ($orphanedMsgs -and $orphanedMsgs.cnt) { [int]$orphanedMsgs.cnt } else { 0 }
$ok = $orphMsgCount -eq 0
$checks += [ordered]@{ name = 'orphaned_messages'; status = if ($ok) { 'pass' } else { 'fail' }; cnt = $orphMsgCount }
if (-not $ok) { $issues += "orphaned_messages: $orphMsgCount messages have no matching session" }

# ---------- 4. Orphaned events (event with no session) ----------
$orphanedEvts = Invoke-Sql -Query "SELECT COUNT(*) as cnt FROM event_log WHERE session_id NOT IN (SELECT id FROM sessions)"
$orphEvtCount = if ($orphanedEvts -and $orphanedEvts.cnt) { [int]$orphanedEvts.cnt } else { 0 }
$ok = $orphEvtCount -eq 0
$checks += [ordered]@{ name = 'orphaned_events'; status = if ($ok) { 'pass' } else { 'fail' }; cnt = $orphEvtCount }
if (-not $ok) { $issues += "orphaned_events: $orphEvtCount events have no matching session" }

# ---------- 5. Orphaned proposals (proposal with no session) ----------
$orphanedProps = Invoke-Sql -Query "SELECT COUNT(*) as cnt FROM action_proposals WHERE session_id NOT IN (SELECT id FROM sessions)"
$orphPropCount = if ($orphanedProps -and $orphanedProps.cnt) { [int]$orphanedProps.cnt } else { 0 }
$ok = $orphPropCount -eq 0
$checks += [ordered]@{ name = 'orphaned_proposals'; status = if ($ok) { 'pass' } else { 'fail' }; cnt = $orphPropCount }
if (-not $ok) { $issues += "orphaned_proposals: $orphPropCount proposals have no matching session" }

# ---------- 6. Orphaned approvals (approval with no proposal) ----------
$orphanedAppr = Invoke-Sql -Query "SELECT COUNT(*) as cnt FROM approvals WHERE proposal_id NOT IN (SELECT id FROM action_proposals)"
$orphApprCount = if ($orphanedAppr -and $orphanedAppr.cnt) { [int]$orphanedAppr.cnt } else { 0 }
$ok = $orphApprCount -eq 0
$checks += [ordered]@{ name = 'orphaned_approvals'; status = if ($ok) { 'pass' } else { 'fail' }; cnt = $orphApprCount }
if (-not $ok) { $issues += "orphaned_approvals: $orphApprCount approvals have no matching proposal" }

# ---------- 7. Orphaned executions (execution with no proposal) ----------
$orphanedExec = Invoke-Sql -Query "SELECT COUNT(*) as cnt FROM action_executions WHERE proposal_id NOT IN (SELECT id FROM action_proposals)"
$orphExecCount = if ($orphanedExec -and $orphanedExec.cnt) { [int]$orphanedExec.cnt } else { 0 }
$ok = $orphExecCount -eq 0
$checks += [ordered]@{ name = 'orphaned_executions'; status = if ($ok) { 'pass' } else { 'fail' }; cnt = $orphExecCount }
if (-not $ok) { $issues += "orphaned_executions: $orphExecCount executions have no matching proposal" }

# ---------- 8. Null/empty timestamps ----------
$nullTimestamps = @()
foreach ($t in @('users', 'sessions', 'messages', 'action_proposals', 'approvals', 'action_executions', 'event_log', 'adapter_states')) {
  $r = Invoke-Sql -Query "SELECT COUNT(*) as cnt FROM $t WHERE created_at IS NULL OR created_at = ''"
  $cnt = if ($r -and $r.cnt) { [int]$r.cnt } else { 0 }
  if ($cnt -gt 0) { $nullTimestamps += "$t=$cnt" }
}
$nullCount = $nullTimestamps.Count
$ok = $nullCount -eq 0
$checks += [ordered]@{ name = 'null_timestamps'; status = if ($ok) { 'pass' } else { 'fail' }; tables = $nullTimestamps }
if (-not $ok) { $issues += "null_timestamps: $($nullTimestamps -join ', ')" }

# ---------- 9. Duplicate primary keys ----------
$dupTables = @()
foreach ($t in $tables) {
  $r = Invoke-Sql -Query "SELECT id, COUNT(*) as cnt FROM $t GROUP BY id HAVING cnt > 1" -All
  if ($r -and $r.Count -gt 0) { $dupTables += "$t=$($r.Count)" }
}
$ok = $dupTables.Count -eq 0
$checks += [ordered]@{ name = 'duplicate_ids'; status = if ($ok) { 'pass' } else { 'fail' }; tables = $dupTables }
if (-not $ok) { $issues += "duplicate_ids: $($dupTables -join ', ')" }

# ---------- 10. Empty sessions (session with 0 events) ----------
$emptySessions = Invoke-Sql -Query "SELECT COUNT(*) as cnt FROM sessions WHERE id NOT IN (SELECT DISTINCT session_id FROM event_log)"
$emptyCount = if ($emptySessions -and $emptySessions.cnt) { [int]$emptySessions.cnt } else { 0 }
$checks += [ordered]@{ name = 'empty_sessions'; status = 'info'; cnt = $emptyCount }

# ---------- 11. Proposals without approval ----------
$unapprovedProps = Invoke-Sql -Query "SELECT COUNT(*) as cnt FROM action_proposals WHERE id NOT IN (SELECT proposal_id FROM approvals)"
$unappCount = if ($unapprovedProps -and $unapprovedProps.cnt) { [int]$unapprovedProps.cnt } else { 0 }
$checks += [ordered]@{ name = 'unapproved_proposals'; status = 'info'; cnt = $unappCount }

# ---------- 12. DB file size ----------
$dbFile = Get-Item $DbPath
$dbSizeKb = [math]::Round($dbFile.Length / 1024, 1)
$checks += [ordered]@{ name = 'db_size_kb'; status = 'info'; value = $dbSizeKb }

# ---------- overall ----------
$failCount = @($checks | Where-Object { $_.status -eq 'fail' }).Count
$passCount = @($checks | Where-Object { $_.status -eq 'pass' }).Count
$infoCount = @($checks | Where-Object { $_.status -eq 'info' }).Count
$overallStatus = if ($failCount -gt 0) { 'issues_found' } else { 'clean' }

$result = [ordered]@{
  generatedAt   = (Get-Date).ToString("o")
  date          = $today
  overallStatus = $overallStatus
  issueCount    = $failCount
  passCount     = $passCount
  infoCount     = $infoCount
  dbSizeKb      = $dbSizeKb
  tableCounts   = $tableCounts
  issues        = $issues
  checks        = $checks
}

# ---------- file output ----------
if ($OutputPath) {
  $outDir = Split-Path -Path $OutputPath -Parent
  if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
  $result | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputPath -Encoding UTF8
  Write-Host "  [ok] Audit written: $OutputPath" -ForegroundColor Green
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $result | ConvertTo-Json -Depth 8
}
else {
  $statusColor = if ($overallStatus -eq 'clean') { 'Green' } else { 'Red' }
  Write-Host "  Status: $overallStatus" -ForegroundColor $statusColor
  Write-Host "  Issues: $failCount  Passed: $passCount  Info: $infoCount"
  Write-Host "  DB size: $($dbSizeKb) KB"
  Write-Host ''

  Write-Host '  Table Counts:' -ForegroundColor White
  foreach ($t in $tableCounts.Keys) {
    $line = "    {0,-22} {1}" -f $t, $tableCounts[$t]
    Write-Host $line
  }
  Write-Host ''

  Write-Host '  Integrity Checks:' -ForegroundColor White
  foreach ($ch in $checks) {
    if ($ch.status -eq 'info') { continue }
    $icon = if ($ch.status -eq 'pass') { '+' } else { '-' }
    $color = if ($ch.status -eq 'pass') { 'Green' } else { 'Red' }
    $detail = if ($null -ne $ch['cnt']) { "  ($($ch['cnt']))" }
              elseif ($ch['tables'] -and $ch['tables'].Count -gt 0) { "  ($($ch['tables'] -join ', '))" }
              else { '' }
    Write-Host "    $icon $($ch.name)$detail" -ForegroundColor $color
  }
  Write-Host ''

  Write-Host '  Info:' -ForegroundColor White
  Write-Host "    empty_sessions=$emptyCount  unapproved_proposals=$unappCount"
  Write-Host ''

  if ($issues.Count -gt 0) {
    Write-Host '  Issues:' -ForegroundColor Red
    foreach ($iss in $issues) {
      Write-Host "    ! $iss" -ForegroundColor Red
    }
    Write-Host ''
  }
}

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host "  [warn] Triage file not found: $triagePath" -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $totalRows = ($tableCounts.Values | Measure-Object -Sum).Sum
  $issueStr = if ($issues.Count -gt 0) { $issues -join '; ' } else { 'none' }
  $line = "$stamp Data quality audit: status=$overallStatus, issues=$failCount, passed=$passCount, totalRows=$totalRows, dbSize=$($dbSizeKb)KB, details=[$issueStr]"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'Data quality audit:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host "  [ok] Upserted data quality audit in triage: $triagePath" -ForegroundColor Green
}

Write-Host '[done] Data quality audit complete.' -ForegroundColor Green
