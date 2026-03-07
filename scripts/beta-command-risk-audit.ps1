param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [switch]$Json,
  [switch]$AppendToTriage,
  [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"

# Normalize emails
$normalizedEmails = @()
foreach ($entry in $Emails) {
  $normalizedEmails += ($entry -split '[,;\s]+' | Where-Object { $_ -and $_.Trim().Length -gt 0 } | ForEach-Object { $_.Trim() })
}
$Emails = $normalizedEmails | Select-Object -Unique

Write-Host '== tsifulator.ai command risk audit ==' -ForegroundColor Cyan
Write-Host "Date: $today"
Write-Host "Partners: $($Emails.Count)"

# ---------- risk classification ----------
# DB stores risk as: safe, blocked
# We re-classify commands into 4 tiers based on patterns
function Get-CommandRisk {
  param([string]$Command, [string]$DbRisk)
  # Blocked in DB = critical
  if ($DbRisk -eq 'blocked') { return 'critical' }
  # Destructive patterns (rm -rf, Remove-Item -Recurse -Force, del /s /q, format, etc.)
  $destructive = @('rm\s+-rf', 'Remove-Item.*-Recurse.*-Force', 'del\s+/[sS]', 'rmdir\s+/[sS]', 'format\s+', 'Drop-', 'DROP\s+', 'TRUNCATE', 'DELETE\s+FROM')
  foreach ($pat in $destructive) {
    if ($Command -match $pat) { return 'high' }
  }
  # Moderate patterns (write/set/new/stop)
  $moderate = @('Set-', 'New-', 'Stop-', 'Start-', 'Restart-', 'Install-', 'Uninstall-', 'Update-', 'Remove-Item', 'mkdir', 'New-Item')
  foreach ($pat in $moderate) {
    if ($Command -match $pat) { return 'moderate' }
  }
  return 'safe'
}

# ---------- helpers ----------
function Invoke-Sql {
  param([string]$Sql)
  $projectRoot = Split-Path $PSScriptRoot -Parent
  $tmpFile = Join-Path $projectRoot "_risk_audit_tmp.js"
  $jsCode = @"
const db = require('better-sqlite3')('./data/tsifulator.db');
const rows = db.prepare($($Sql | ConvertTo-Json)).all();
console.log(JSON.stringify(rows));
"@
  Set-Content -Path $tmpFile -Value $jsCode -Encoding utf8
  try {
    $raw = & node $tmpFile 2>&1
    if ($LASTEXITCODE -ne 0) { throw "SQL failed: $raw" }
    return ($raw | ConvertFrom-Json)
  } finally {
    Remove-Item -Path $tmpFile -Force -ErrorAction SilentlyContinue
  }
}

# ---------- gather data ----------
$sql = 'SELECT ap.id as proposal_id, ap.command, ap.risk as db_risk, ap.session_id, s.user_id, u.email, a.approved, ae.status as exec_status FROM action_proposals ap JOIN sessions s ON ap.session_id = s.id JOIN users u ON s.user_id = u.id LEFT JOIN approvals a ON a.proposal_id = ap.id LEFT JOIN action_executions ae ON ae.proposal_id = ap.id'
$allRows = Invoke-Sql -Sql $sql

# Filter to target partners only
$partnerRows = @($allRows | Where-Object { $Emails -contains $_.email })

Write-Host ""
Write-Host "Total proposals (all users): $($allRows.Count)"
Write-Host "Partner proposals: $($partnerRows.Count)"

# ---------- per-partner analysis ----------
$partnerResults = [ordered]@{}
$globalCommands = @{}
$globalCritical = @()

foreach ($email in $Emails) {
  $rows = @($partnerRows | Where-Object { $_.email -eq $email })
  $total = $rows.Count

  # Risk classification
  $riskBuckets = [ordered]@{ safe = 0; moderate = 0; high = 0; critical = 0 }
  $commandFreq = @{}
  $blockedCmds = @()
  $approved = 0
  $denied = 0
  $pending = 0
  $executed = 0
  $failed = 0

  foreach ($row in $rows) {
    $risk = Get-CommandRisk -Command $row.command -DbRisk $row.db_risk
    $riskBuckets[$risk] += 1

    # Command frequency
    $cmd = $row.command
    if (-not $commandFreq.Contains($cmd)) { $commandFreq[$cmd] = 0 }
    $commandFreq[$cmd] += 1

    # Global command tracking
    if (-not $globalCommands.Contains($cmd)) { $globalCommands[$cmd] = 0 }
    $globalCommands[$cmd] += 1

    # Approval status
    if ($null -eq $row.approved) { $pending += 1 }
    elseif ($row.approved -eq 1) { $approved += 1 }
    else { $denied += 1 }

    # Execution status
    if ($row.exec_status -eq 'success') { $executed += 1 }
    elseif ($row.exec_status -eq 'failed') { $failed += 1 }

    # Track blocked/critical commands
    if ($risk -eq 'critical') {
      $blockedCmds += $cmd
      $globalCritical += [ordered]@{ email = $email; command = $cmd }
    }
  }

  # Block rate = (critical + high) / total
  $dangerousCount = $riskBuckets['critical'] + $riskBuckets['high']
  $blockRate = if ($total -gt 0) { [math]::Round(($dangerousCount / $total) * 100, 1) } else { 0 }

  # Safety gate effectiveness = blocked / (blocked + high that slipped through)
  $highSlipped = 0
  foreach ($row in $rows) {
    $risk = Get-CommandRisk -Command $row.command -DbRisk $row.db_risk
    if ($risk -eq 'high' -and $row.approved -eq 1) { $highSlipped += 1 }
  }
  $gateEff = if (($riskBuckets['critical'] + $highSlipped) -gt 0) {
    [math]::Round(($riskBuckets['critical'] / ($riskBuckets['critical'] + $highSlipped)) * 100, 1)
  } else { 100 }

  # Top command
  $topCmd = ''
  $topCmdCount = 0
  foreach ($k in $commandFreq.Keys) {
    if ($commandFreq[$k] -gt $topCmdCount) {
      $topCmd = $k
      $topCmdCount = $commandFreq[$k]
    }
  }

  $partnerResults[$email] = [ordered]@{
    email         = $email
    totalProposals = $total
    risk          = $riskBuckets
    approved      = $approved
    denied        = $denied
    pending       = $pending
    executed      = $executed
    failed        = $failed
    blockRate     = $blockRate
    gateEfficiency = $gateEff
    topCommand    = $topCmd
    topCommandCount = $topCmdCount
    blockedCommands = $blockedCmds
  }
}

# ---------- fleet summary ----------
$fleetTotal = ($partnerResults.Values | ForEach-Object { $_['totalProposals'] } | Measure-Object -Sum).Sum
$fleetCritical = ($partnerResults.Values | ForEach-Object { $_['risk']['critical'] } | Measure-Object -Sum).Sum
$fleetHigh = ($partnerResults.Values | ForEach-Object { $_['risk']['high'] } | Measure-Object -Sum).Sum
$fleetSafe = ($partnerResults.Values | ForEach-Object { $_['risk']['safe'] } | Measure-Object -Sum).Sum
$fleetModerate = ($partnerResults.Values | ForEach-Object { $_['risk']['moderate'] } | Measure-Object -Sum).Sum
$fleetBlockRate = if ($fleetTotal -gt 0) { [math]::Round((($fleetCritical + $fleetHigh) / $fleetTotal) * 100, 1) } else { 0 }

# Fleet gate effectiveness
$fleetGateNumerator = $fleetCritical
$fleetHighSlipped = 0
foreach ($email in $Emails) {
  $rows = @($partnerRows | Where-Object { $_.email -eq $email })
  foreach ($row in $rows) {
    $risk = Get-CommandRisk -Command $row.command -DbRisk $row.db_risk
    if ($risk -eq 'high' -and $row.approved -eq 1) { $fleetHighSlipped += 1 }
  }
}
$fleetGateEff = if (($fleetGateNumerator + $fleetHighSlipped) -gt 0) {
  [math]::Round(($fleetGateNumerator / ($fleetGateNumerator + $fleetHighSlipped)) * 100, 1)
} else { 100 }

# Health assessment
$health = if ($fleetBlockRate -gt 30) { 'at-risk' }
          elseif ($fleetBlockRate -gt 15) { 'elevated' }
          elseif ($fleetBlockRate -gt 5) { 'moderate' }
          else { 'healthy' }

# Unique commands across fleet
$uniqueCmds = $globalCommands.Keys.Count

# Top global command
$topGlobal = ''
$topGlobalCount = 0
foreach ($k in $globalCommands.Keys) {
  if ($globalCommands[$k] -gt $topGlobalCount) {
    $topGlobal = $k
    $topGlobalCount = $globalCommands[$k]
  }
}

$summary = [ordered]@{
  date            = $today
  totalProposals  = $fleetTotal
  uniqueCommands  = $uniqueCmds
  riskBreakdown   = [ordered]@{ safe = $fleetSafe; moderate = $fleetModerate; high = $fleetHigh; critical = $fleetCritical }
  blockRate       = $fleetBlockRate
  gateEfficiency  = $fleetGateEff
  health          = $health
  topCommand      = $topGlobal
  topCommandCount = $topGlobalCount
  criticalCount   = $globalCritical.Count
  partners        = [array]$partnerResults.Values
}

# ---------- output ----------
if ($Json) {
  $summary | ConvertTo-Json -Depth 5
  exit 0
}

if ($OutputPath) {
  $summary | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputPath -Encoding utf8
  Write-Host "[saved] $OutputPath"
}

# ---------- display ----------
Write-Host ''
Write-Host '--- Fleet Risk Summary ---' -ForegroundColor Cyan
Write-Host "  Total proposals:  $fleetTotal"
Write-Host "  Unique commands:  $uniqueCmds"
Write-Host "  Top command:      $topGlobal ($topGlobalCount uses)"
Write-Host ''
Write-Host '  Risk breakdown:' -ForegroundColor Yellow
Write-Host "    Safe:     $fleetSafe"
Write-Host "    Moderate: $fleetModerate"
Write-Host "    High:     $fleetHigh"
Write-Host "    Critical: $fleetCritical"
Write-Host ''
$blockStr = "  Block rate: $fleetBlockRate%"
if ($fleetBlockRate -gt 15) {
  Write-Host $blockStr -ForegroundColor Red
} elseif ($fleetBlockRate -gt 5) {
  Write-Host $blockStr -ForegroundColor Yellow
} else {
  Write-Host $blockStr -ForegroundColor Green
}
$gateStr = "  Gate efficiency: $fleetGateEff%"
if ($fleetGateEff -ge 90) {
  Write-Host $gateStr -ForegroundColor Green
} else {
  Write-Host $gateStr -ForegroundColor Red
}
$healthStr = "  Health: $health"
if ($health -eq 'healthy') {
  Write-Host $healthStr -ForegroundColor Green
} elseif ($health -eq 'at-risk') {
  Write-Host $healthStr -ForegroundColor Red
} else {
  Write-Host $healthStr -ForegroundColor Yellow
}

Write-Host ''
Write-Host '--- Per-Partner Detail ---' -ForegroundColor Cyan
foreach ($email in $Emails) {
  $p = $partnerResults[$email]
  Write-Host "  $email" -ForegroundColor White
  Write-Host "    Proposals: $($p['totalProposals'])  Approved: $($p['approved'])  Denied: $($p['denied'])  Pending: $($p['pending'])"
  Write-Host "    Executed: $($p['executed'])  Failed: $($p['failed'])"
  Write-Host "    Risk: safe=$($p['risk']['safe']) moderate=$($p['risk']['moderate']) high=$($p['risk']['high']) critical=$($p['risk']['critical'])"
  Write-Host "    Block rate: $($p['blockRate'])%  Gate efficiency: $($p['gateEfficiency'])%"
  Write-Host "    Top command: $($p['topCommand']) ($($p['topCommandCount'])x)"
  if ($p['blockedCommands'].Count -gt 0) {
    Write-Host "    Blocked: $($p['blockedCommands'] -join ', ')" -ForegroundColor Red
  }
  Write-Host ''
}

if ($globalCritical.Count -gt 0) {
  Write-Host '--- Critical Command Log ---' -ForegroundColor Red
  foreach ($c in $globalCritical) {
    Write-Host "  [$($c['email'])] $($c['command'])" -ForegroundColor Red
  }
  Write-Host ''
}

# ---------- triage ----------
if ($AppendToTriage) {
  if (-not (Test-Path (Split-Path $triagePath -Parent))) {
    New-Item -ItemType Directory -Path (Split-Path $triagePath -Parent) -Force | Out-Null
  }
  $triageLine = "Command risk audit: proposals=$fleetTotal unique=$uniqueCmds safe=$fleetSafe moderate=$fleetModerate high=$fleetHigh critical=$fleetCritical blockRate=$fleetBlockRate% gateEff=$fleetGateEff% health=$health"
  if (Test-Path $triagePath) {
    $existing = Get-Content -Path $triagePath -Raw
    $cleaned = ($existing -split "`n" | Where-Object { $_ -notmatch '^Command risk audit:' }) -join "`n"
    Set-Content -Path $triagePath -Value $cleaned.TrimEnd() -Encoding utf8
  }
  Add-Content -Path $triagePath -Value $triageLine -Encoding utf8
  Write-Host "[triage] appended to $triagePath" -ForegroundColor Green
}

Write-Host 'done.' -ForegroundColor Green
