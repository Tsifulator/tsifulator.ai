param(
  [string]$HistoryPath = "docs/reports/partner-compare-history.csv",
  [int]$RetainDays = 30,
  [switch]$Archive,
  [string]$ArchivePath = "",
  [switch]$DryRun,
  [switch]$AppendToTriage,
  [string]$Date
)

$ErrorActionPreference = "Stop"

$today = if ($Date) { $Date } else { Get-Date -Format "yyyy-MM-dd" }
$triagePath = "docs/daily-triage/$today.md"
$cutoffDate = (Get-Date $today).AddDays(-$RetainDays).ToString("yyyy-MM-dd")

Write-Host '== tsifulator.ai history rotation ==' -ForegroundColor Cyan
Write-Host "Date: $today"
Write-Host "History: $HistoryPath"
Write-Host "Retain: $RetainDays days  (cutoff: $cutoffDate)"
if ($DryRun) { Write-Host '  [mode] DRY RUN - no files will be modified' -ForegroundColor DarkYellow }

# ---------- guard ----------
if (-not (Test-Path $HistoryPath)) {
  Write-Host ('  [warn] History file not found: ' + $HistoryPath) -ForegroundColor DarkYellow
  Write-Host '[done] Nothing to rotate.' -ForegroundColor Green
  exit 0
}

$allRows = Import-Csv -Path $HistoryPath
$totalBefore = if ($allRows) { $allRows.Count } else { 0 }

if ($totalBefore -eq 0) {
  Write-Host '  [info] History file is empty.' -ForegroundColor Gray
  Write-Host '[done] Nothing to rotate.' -ForegroundColor Green
  exit 0
}

# ---------- partition ----------
$keepRows = @()
$pruneRows = @()

foreach ($row in $allRows) {
  if ($row.snapshotDate -ge $cutoffDate) {
    $keepRows += $row
  }
  else {
    $pruneRows += $row
  }
}

$prunedCount = $pruneRows.Count
$keptCount = $keepRows.Count
$prunedDates = @($pruneRows | Select-Object -ExpandProperty snapshotDate -Unique | Sort-Object)
$keptDates = @($keepRows | Select-Object -ExpandProperty snapshotDate -Unique | Sort-Object)

Write-Host ''
Write-Host '  Summary:' -ForegroundColor White
Write-Host "    total rows before:  $totalBefore"
Write-Host "    rows to prune:      $prunedCount"
Write-Host "    rows to keep:       $keptCount"
if ($prunedDates.Count -gt 0) {
  Write-Host "    pruned date range:  $($prunedDates[0]) .. $($prunedDates[-1])  ($($prunedDates.Count) days)"
}
if ($keptDates.Count -gt 0) {
  Write-Host "    kept date range:    $($keptDates[0]) .. $($keptDates[-1])  ($($keptDates.Count) days)"
}

if ($prunedCount -eq 0) {
  Write-Host ''
  Write-Host '  [ok] No rows older than cutoff. Nothing to rotate.' -ForegroundColor Green
  Write-Host '[done] History rotation complete.' -ForegroundColor Green
  exit 0
}

# ---------- dry run exit ----------
if ($DryRun) {
  Write-Host ''
  Write-Host '  [dry-run] Would prune ' -NoNewline
  Write-Host "$prunedCount rows" -ForegroundColor Yellow -NoNewline
  Write-Host ' and keep ' -NoNewline
  Write-Host "$keptCount rows" -ForegroundColor Green
  Write-Host '[done] Dry run complete.' -ForegroundColor Green
  exit 0
}

# ---------- archive ----------
if ($Archive) {
  $effectiveArchivePath = if ($ArchivePath) { $ArchivePath } else {
    $dir = Split-Path -Path $HistoryPath -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($HistoryPath)
    Join-Path $dir "$baseName-archive.csv"
  }

  $archiveDir = Split-Path -Path $effectiveArchivePath -Parent
  if ($archiveDir -and -not (Test-Path $archiveDir)) {
    New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null
  }

  # Append to existing archive (merge, no duplicates by snapshotAt + email)
  $existingArchive = @()
  if (Test-Path $effectiveArchivePath) {
    $existingArchive = @(Import-Csv -Path $effectiveArchivePath)
  }

  $archiveKeys = @{}
  foreach ($ea in $existingArchive) {
    $key = "$($ea.snapshotAt)|$($ea.email)"
    $archiveKeys[$key] = $true
  }

  $newArchiveRows = @()
  foreach ($pr in $pruneRows) {
    $key = "$($pr.snapshotAt)|$($pr.email)"
    if (-not $archiveKeys.ContainsKey($key)) {
      $newArchiveRows += $pr
    }
  }

  $mergedArchive = $existingArchive + $newArchiveRows
  if ($mergedArchive.Count -gt 0) {
    $mergedArchive | Export-Csv -Path $effectiveArchivePath -NoTypeInformation -Encoding UTF8
  }

  Write-Host ('  [ok] Archived ' + $newArchiveRows.Count + ' rows to ' + $effectiveArchivePath) -ForegroundColor Green
}

# ---------- write rotated history ----------
if ($keptCount -gt 0) {
  $keepRows | Export-Csv -Path $HistoryPath -NoTypeInformation -Encoding UTF8
}
else {
  # Keep header only
  $header = (Get-Content -Path $HistoryPath -TotalCount 1)
  Set-Content -Path $HistoryPath -Value $header -Encoding UTF8
}

Write-Host ('  [ok] Rotated history: ' + $totalBefore + ' -> ' + $keptCount + ' rows') -ForegroundColor Green

# ---------- triage upsert ----------
if ($AppendToTriage) {
  if (-not (Test-Path $triagePath)) {
    Write-Host ('  [warn] Triage file not found: ' + $triagePath) -ForegroundColor DarkYellow
    exit 1
  }

  $stamp = (Get-Date).ToString("o")
  $archiveNote = if ($Archive) { ', archived' } else { '' }
  $line = "$stamp History rotation: pruned=$prunedCount, kept=$keptCount, retainDays=$RetainDays, cutoff=$cutoffDate$archiveNote"

  $triageRaw = Get-Content -Path $triagePath -Raw
  $triageLines = $triageRaw -split "`r?`n"
  $filteredLines = $triageLines | Where-Object { $_ -notmatch 'History rotation:' }
  $updatedTriage = ($filteredLines -join "`r`n").TrimEnd() + "`r`n- $line`r`n"
  Set-Content -Path $triagePath -Value $updatedTriage -Encoding UTF8

  Write-Host ('  [ok] Upserted history rotation in triage: ' + $triagePath) -ForegroundColor Green
}

Write-Host '[done] History rotation complete.' -ForegroundColor Green
