param(
  [string]$BaseUrl = "http://127.0.0.1:4000",
  [string[]]$Emails = @("partner.a@company.com", "partner.b@company.com"),
  [string]$Date,
  [switch]$Json,
  [switch]$AppendToTriage,
  [int]$ChatLatencyTarget = 5000,
  [int]$ProposalLatencyTarget = 3000,
  [int]$StreamFirstTokenTarget = 2000
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

Write-Host '== tsifulator.ai SLA monitor ==' -ForegroundColor Cyan
Write-Host "Date:     $today"
Write-Host "Partners: $($Emails -join ', ')"
Write-Host "Targets:  chat<$($ChatLatencyTarget)ms  proposal<$($ProposalLatencyTarget)ms  stream<$($StreamFirstTokenTarget)ms"

# API health check
try {
  $health = Invoke-RestMethod -Method Get -Uri "$BaseUrl/health" -TimeoutSec 5
  if ($health.status -ne 'ok') { throw 'bad status' }
  Write-Host '[ok] API healthy: ok' -ForegroundColor Green
} catch {
  Write-Host '[FAIL] API unreachable' -ForegroundColor Red
  exit 1
}

# ---------- helper ----------
function Get-Percentile {
  param([double[]]$Values, [int]$Pct)
  if ($Values.Count -eq 0) { return 0 }
  $sorted = $Values | Sort-Object
  $idx = [math]::Ceiling($Pct / 100.0 * $sorted.Count) - 1
  if ($idx -lt 0) { $idx = 0 }
  return $sorted[$idx]
}

# ---------- collect per-partner SLA metrics ----------
$partnerResults = @()
$allChatLatencies = @()
$allProposalLatencies = @()
$allStreamLatencies = @()
$totalBreaches = 0

foreach ($em in $Emails) {
  $pr = [ordered]@{
    email              = $em
    reachable          = $false
    sessions           = 0
    chatLatencies      = @()
    proposalLatencies  = @()
    streamLatencies    = @()
    chatMedian         = 0
    chatP95            = 0
    chatMax            = 0
    proposalMedian     = 0
    proposalP95        = 0
    proposalMax        = 0
    streamMedian       = 0
    streamP95          = 0
    streamMax          = 0
    chatBreaches       = 0
    proposalBreaches   = 0
    streamBreaches     = 0
    totalBreaches      = 0
    slaStatus          = 'unknown'
  }

  try {
    $lb = @{ email = $em } | ConvertTo-Json
    $lg = Invoke-RestMethod -Method Post -Uri "$BaseUrl/auth/dev-login" -ContentType 'application/json' -Body $lb
    $hd = @{ authorization = "Bearer $($lg.token)" }
    $sess = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions?limit=100" -Headers $hd
    $pr.reachable = $true
    $pr.sessions = if ($sess.sessions) { $sess.sessions.Count } else { 0 }

    foreach ($s in $sess.sessions) {
      try {
        $ev = Invoke-RestMethod -Method Get -Uri "$BaseUrl/sessions/$($s.id)/events" -Headers $hd
        if (-not $ev.events -or $ev.events.Count -eq 0) { continue }

        # Sort events chronologically
        $events = @($ev.events | Sort-Object { [DateTime]::Parse($_.createdAt) })

        for ($i = 0; $i -lt $events.Count; $i++) {
          $e = $events[$i]
          $eTime = [DateTime]::Parse($e.createdAt)

          # Chat latency: user_message -> next assistant_message
          if ($e.type -eq 'chat_user_message') {
            for ($j = $i + 1; $j -lt $events.Count; $j++) {
              if ($events[$j].type -eq 'chat_assistant_message') {
                $respTime = [DateTime]::Parse($events[$j].createdAt)
                $latencyMs = [math]::Round(($respTime - $eTime).TotalMilliseconds)
                $pr.chatLatencies += $latencyMs
                if ($latencyMs -gt $ChatLatencyTarget) { $pr.chatBreaches++ }
                break
              }
            }
          }

          # Proposal latency: user_message -> next action_proposed
          if ($e.type -eq 'chat_user_message') {
            for ($j = $i + 1; $j -lt $events.Count; $j++) {
              if ($events[$j].type -eq 'action_proposed') {
                $propTime = [DateTime]::Parse($events[$j].createdAt)
                $latencyMs = [math]::Round(($propTime - $eTime).TotalMilliseconds)
                $pr.proposalLatencies += $latencyMs
                if ($latencyMs -gt $ProposalLatencyTarget) { $pr.proposalBreaches++ }
                break
              }
              # Stop if we hit another user message first
              if ($events[$j].type -eq 'chat_user_message') { break }
            }
          }

          # Stream first-token latency: stream_started -> stream_completed
          if ($e.type -eq 'chat_stream_started') {
            for ($j = $i + 1; $j -lt $events.Count; $j++) {
              if ($events[$j].type -eq 'chat_stream_completed') {
                $endTime = [DateTime]::Parse($events[$j].createdAt)
                $latencyMs = [math]::Round(($endTime - $eTime).TotalMilliseconds)
                $pr.streamLatencies += $latencyMs
                if ($latencyMs -gt $StreamFirstTokenTarget) { $pr.streamBreaches++ }
                break
              }
            }
          }
        }
      } catch {}
    }

    # Compute stats
    if ($pr.chatLatencies.Count -gt 0) {
      $pr.chatMedian = Get-Percentile -Values $pr.chatLatencies -Pct 50
      $pr.chatP95 = Get-Percentile -Values $pr.chatLatencies -Pct 95
      $pr.chatMax = ($pr.chatLatencies | Measure-Object -Maximum).Maximum
      $allChatLatencies += $pr.chatLatencies
    }
    if ($pr.proposalLatencies.Count -gt 0) {
      $pr.proposalMedian = Get-Percentile -Values $pr.proposalLatencies -Pct 50
      $pr.proposalP95 = Get-Percentile -Values $pr.proposalLatencies -Pct 95
      $pr.proposalMax = ($pr.proposalLatencies | Measure-Object -Maximum).Maximum
      $allProposalLatencies += $pr.proposalLatencies
    }
    if ($pr.streamLatencies.Count -gt 0) {
      $pr.streamMedian = Get-Percentile -Values $pr.streamLatencies -Pct 50
      $pr.streamP95 = Get-Percentile -Values $pr.streamLatencies -Pct 95
      $pr.streamMax = ($pr.streamLatencies | Measure-Object -Maximum).Maximum
      $allStreamLatencies += $pr.streamLatencies
    }

    $pr.totalBreaches = $pr.chatBreaches + $pr.proposalBreaches + $pr.streamBreaches
    $totalBreaches += $pr.totalBreaches

    # SLA status per partner
    $pr.slaStatus = if ($pr.totalBreaches -eq 0) { 'met' }
                    elseif ($pr.totalBreaches -le 2) { 'warning' }
                    else { 'breached' }

  } catch {
    Write-Host "  [warn] Could not reach $em" -ForegroundColor DarkYellow
  }

  $partnerResults += $pr
}

# ---------- global stats ----------
$globalChatMedian = if ($allChatLatencies.Count -gt 0) { Get-Percentile -Values $allChatLatencies -Pct 50 } else { 0 }
$globalChatP95 = if ($allChatLatencies.Count -gt 0) { Get-Percentile -Values $allChatLatencies -Pct 95 } else { 0 }
$globalProposalMedian = if ($allProposalLatencies.Count -gt 0) { Get-Percentile -Values $allProposalLatencies -Pct 50 } else { 0 }
$globalProposalP95 = if ($allProposalLatencies.Count -gt 0) { Get-Percentile -Values $allProposalLatencies -Pct 95 } else { 0 }
$globalStreamMedian = if ($allStreamLatencies.Count -gt 0) { Get-Percentile -Values $allStreamLatencies -Pct 50 } else { 0 }
$globalStreamP95 = if ($allStreamLatencies.Count -gt 0) { Get-Percentile -Values $allStreamLatencies -Pct 95 } else { 0 }

$reachable = @($partnerResults | Where-Object { $_.reachable })
$globalStatus = if ($totalBreaches -eq 0) { 'met' }
                elseif ($totalBreaches -le ($reachable.Count * 2)) { 'warning' }
                else { 'breached' }

$totalMeasurements = $allChatLatencies.Count + $allProposalLatencies.Count + $allStreamLatencies.Count
$compliancePct = if ($totalMeasurements -gt 0) {
  [math]::Round(($totalMeasurements - $totalBreaches) / $totalMeasurements * 100, 1)
} else { 100 }

# ---------- build output ----------
$output = [ordered]@{
  generatedAt       = (Get-Date).ToString('o')
  date              = $today
  slaStatus         = $globalStatus
  compliancePct     = $compliancePct
  totalMeasurements = $totalMeasurements
  totalBreaches     = $totalBreaches
  targets           = [ordered]@{
    chatLatencyMs        = $ChatLatencyTarget
    proposalLatencyMs    = $ProposalLatencyTarget
    streamFirstTokenMs   = $StreamFirstTokenTarget
  }
  global            = [ordered]@{
    chat     = [ordered]@{ samples = $allChatLatencies.Count; medianMs = $globalChatMedian; p95Ms = $globalChatP95 }
    proposal = [ordered]@{ samples = $allProposalLatencies.Count; medianMs = $globalProposalMedian; p95Ms = $globalProposalP95 }
    stream   = [ordered]@{ samples = $allStreamLatencies.Count; medianMs = $globalStreamMedian; p95Ms = $globalStreamP95 }
  }
  partners          = @()
}

foreach ($pr in $partnerResults) {
  $p = [ordered]@{
    email             = $pr.email
    reachable         = $pr.reachable
    sessions          = $pr.sessions
    slaStatus         = $pr.slaStatus
    totalBreaches     = $pr.totalBreaches
    chat              = [ordered]@{
      samples   = $pr.chatLatencies.Count
      medianMs  = $pr.chatMedian
      p95Ms     = $pr.chatP95
      maxMs     = $pr.chatMax
      breaches  = $pr.chatBreaches
    }
    proposal          = [ordered]@{
      samples   = $pr.proposalLatencies.Count
      medianMs  = $pr.proposalMedian
      p95Ms     = $pr.proposalP95
      maxMs     = $pr.proposalMax
      breaches  = $pr.proposalBreaches
    }
    stream            = [ordered]@{
      samples   = $pr.streamLatencies.Count
      medianMs  = $pr.streamMedian
      p95Ms     = $pr.streamP95
      maxMs     = $pr.streamMax
      breaches  = $pr.streamBreaches
    }
  }
  $output.partners += $p
}

# ---------- display ----------
Write-Host ''
if ($Json) {
  $output | ConvertTo-Json -Depth 8
} else {
  $statusColor = switch ($globalStatus) {
    'met'      { 'Green' }
    'warning'  { 'Yellow' }
    'breached' { 'Red' }
    default    { 'DarkGray' }
  }
  $statusIcon = switch ($globalStatus) {
    'met'      { '+' }
    'warning'  { '~' }
    'breached' { '!' }
    default    { '?' }
  }
  Write-Host "  SLA: $globalStatus ($statusIcon) - compliance $compliancePct% ($($totalMeasurements - $totalBreaches)/$totalMeasurements measurements pass)" -ForegroundColor $statusColor
  Write-Host ''

  # Global latency table
  Write-Host '  Global latencies:' -ForegroundColor White
  $ghdr = '  {0,-12} {1,8} {2,8} {3,8} {4,10}' -f 'METRIC','MEDIAN','P95','TARGET','STATUS'
  Write-Host $ghdr -ForegroundColor DarkGray

  $chatStatus = if ($globalChatP95 -le $ChatLatencyTarget) { 'OK' } else { 'BREACH' }
  $chatColor = if ($chatStatus -eq 'OK') { 'Green' } else { 'Red' }
  $chatRow = '  {0,-12} {1,6}ms {2,6}ms {3,6}ms {4,10}' -f 'chat', $globalChatMedian, $globalChatP95, $ChatLatencyTarget, $chatStatus
  Write-Host $chatRow -ForegroundColor $chatColor

  $propStatus = if ($globalProposalP95 -le $ProposalLatencyTarget) { 'OK' } else { 'BREACH' }
  $propColor = if ($propStatus -eq 'OK') { 'Green' } else { 'Red' }
  $propRow = '  {0,-12} {1,6}ms {2,6}ms {3,6}ms {4,10}' -f 'proposal', $globalProposalMedian, $globalProposalP95, $ProposalLatencyTarget, $propStatus
  Write-Host $propRow -ForegroundColor $propColor

  $streamStatus = if ($globalStreamP95 -le $StreamFirstTokenTarget) { 'OK' } else { 'BREACH' }
  $streamColor = if ($streamStatus -eq 'OK') { 'Green' } else { 'Red' }
  $streamRow = '  {0,-12} {1,6}ms {2,6}ms {3,6}ms {4,10}' -f 'stream', $globalStreamMedian, $globalStreamP95, $StreamFirstTokenTarget, $streamStatus
  Write-Host $streamRow -ForegroundColor $streamColor

  Write-Host ''

  # Per-partner breakdown
  Write-Host '  Per partner:' -ForegroundColor White
  $phdr = '  {0,-25} {1,4} {2,8} {3,8} {4,8} {5,8} {6,5}' -f 'EMAIL','SESS','CHAT-M','CHAT-95','PROP-M','PROP-95','BRCHS'
  Write-Host $phdr -ForegroundColor DarkGray

  foreach ($pr in $partnerResults) {
    if (-not $pr.reachable) {
      Write-Host "  $($pr.email): [unreachable]" -ForegroundColor Red
      continue
    }

    $pColor = switch ($pr.slaStatus) {
      'met'      { 'Green' }
      'warning'  { 'Yellow' }
      'breached' { 'Red' }
      default    { 'DarkGray' }
    }

    $shortEmail = $pr.email
    if ($shortEmail.Length -gt 25) { $shortEmail = $shortEmail.Substring(0, 22) + '...' }
    $pRow = '  {0,-25} {1,4} {2,6}ms {3,6}ms {4,6}ms {5,6}ms {6,5}' -f $shortEmail, $pr.sessions, $pr.chatMedian, $pr.chatP95, $pr.proposalMedian, $pr.proposalP95, $pr.totalBreaches
    Write-Host $pRow -ForegroundColor $pColor

    # Breach details
    if ($pr.chatBreaches -gt 0) {
      Write-Host "      ! chat: $($pr.chatBreaches) breach(es), max=$($pr.chatMax)ms" -ForegroundColor Red
    }
    if ($pr.proposalBreaches -gt 0) {
      Write-Host "      ! proposal: $($pr.proposalBreaches) breach(es), max=$($pr.proposalMax)ms" -ForegroundColor Red
    }
    if ($pr.streamBreaches -gt 0) {
      Write-Host "      ! stream: $($pr.streamBreaches) breach(es), max=$($pr.streamMax)ms" -ForegroundColor Red
    }
  }

  # Insights
  Write-Host ''
  Write-Host '  Insights:' -ForegroundColor White
  if ($totalBreaches -eq 0) {
    Write-Host '    + All SLA targets met across all partners' -ForegroundColor Green
  }
  if ($compliancePct -ge 99 -and $totalBreaches -gt 0) {
    Write-Host '    ~ Minor SLA breaches detected but compliance is above 99%' -ForegroundColor Yellow
  }
  if ($compliancePct -lt 95) {
    Write-Host '    ! SLA compliance below 95% - investigate response times' -ForegroundColor Red
  }
  $highLatency = @($partnerResults | Where-Object { $_.chatP95 -gt $ChatLatencyTarget })
  if ($highLatency.Count -gt 0) {
    $names = ($highLatency | ForEach-Object { $_.email }) -join ', '
    Write-Host "    ! Partners with high chat latency (p95 > target): $names" -ForegroundColor Red
  }
  if ($allChatLatencies.Count -eq 0 -and $allProposalLatencies.Count -eq 0) {
    Write-Host '    ~ No latency samples found - events may lack timing resolution' -ForegroundColor DarkYellow
  }
  if ($globalChatMedian -le 100 -and $allChatLatencies.Count -gt 0) {
    Write-Host "    + Sub-100ms median chat response - excellent performance" -ForegroundColor Green
  }
}

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host ('  [warn] Triage file not found: ' + $triagePath) -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString('o')
  $partnerSummaries = ($partnerResults | Where-Object { $_.reachable } | ForEach-Object {
    "$($_.email)=$($_.slaStatus)(chat:$($_.chatMedian)ms,prop:$($_.proposalMedian)ms,breaches:$($_.totalBreaches))"
  }) -join ', '
  $line = "$stamp SLA monitor: status=$globalStatus, compliance=$compliancePct%, measurements=$totalMeasurements, breaches=$totalBreaches, chatP95=$($globalChatP95)ms, propP95=$($globalProposalP95)ms, partners=($partnerSummaries)"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'SLA monitor:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ''
  Write-Host ('  [ok] Upserted SLA monitor in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host ''
Write-Host '[done] SLA monitor complete.' -ForegroundColor Green
