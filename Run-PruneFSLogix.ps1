[CmdletBinding()]
param(
    [ValidateSet('South','West','Both')]
    [string]$Region = 'Both',

    [switch]$DryRun,
    [switch]$Live,
    [int]$DaysOld = 60,
    [switch]$SkipIfLocked,
    [string[]]$UserFilter,

    [int]$RecurseDepth = 5,
    [string[]]$IncludePatterns = @('*.vhdx'),
    [switch]$VerboseOutput,

    # NEW: only print the helper’s own roll-up if you ask for it
    [switch]$ExtraSummary
)

# ---------------- CONFIG ----------------
$ScriptPath = '\\saavduksg2v.file.core.windows.net\fslogix\_SCRIPTS\Prune-FSLogixVhdx.ps1'
$SouthPath = '\\saavduksg2v.file.core.windows.net\fslogix\Profiles'
$WestPath  = '\\SAAVDUKWG2V.File.core.windows.net\fslogix\Profiles_W11'

switch ($Region) {
    'South' { $Paths = @($SouthPath) }
    'West'  { $Paths = @($WestPath)  }
    'Both'  { $Paths = @($SouthPath,$WestPath) }
}

if (-not (Test-Path $ScriptPath)) { Write-Error "Main script not found at $ScriptPath"; exit 1 }

if (-not $Live) { $DryRun = $true }

# ---------------- Run Main Script ----------------
$paramHash = @{
    DaysOld         = $DaysOld
    Paths           = $Paths
    IncludePatterns = $IncludePatterns
    RecurseDepth    = $RecurseDepth
}
if ($SkipIfLocked) { $paramHash['SkipIfLocked'] = $true }
if ($VerboseOutput){ $paramHash['Verbose']      = $true }
if ($UserFilter)   { $paramHash['UserFilter']   = $UserFilter }

if ($Live) {
    $paramHash['Confirm'] = $false
    Write-Host "🔴 LIVE DELETE (files will be removed)" -ForegroundColor Red
    Start-Sleep -Seconds 1
} else {
    $paramHash['WhatIf'] = $true
    Write-Host "🟡 DRY RUN (no deletions will occur)" -ForegroundColor Yellow
}

Write-Host "`nTarget region: $Region" -ForegroundColor Cyan
Write-Host "Scanning paths:`n  $($Paths -join "`n  `n")" -ForegroundColor Gray
Write-Host "---------------------------------------------------------------`n"

# The main script prints its own summary and returns raw objects
$results = & $ScriptPath @paramHash

# ----- By default, stop here so you only see ONE summary -----
if (-not $ExtraSummary) { return }

# ----- Optional extra roll-up from helper (only if -ExtraSummary) -----
if (-not $results) { return }

function Get-Bytes($r) {
    if ($r.PSObject.Properties.Name -contains 'Bytes' -and $r.Bytes) { [double]$r.Bytes }
    elseif ($r.PSObject.Properties.Name -contains 'SizeGB' -and $r.SizeGB) { [double]$r.SizeGB * 1GB }
    else { 0 }
}
function Format-Bytes([double]$b) {
    if ($b -ge 1TB) { "{0:N2} TB" -f ($b / 1TB) }
    elseif ($b -ge 1GB) { "{0:N2} GB" -f ($b / 1GB) }
    elseif ($b -ge 1MB) { "{0:N2} MB" -f ($b / 1MB) }
    else { "{0:N0} B" -f $b }
}

$would = @($results | Where-Object { $_.Action -eq 'WouldDelete' })
$del   = @($results | Where-Object { $_.Action -eq 'Deleted' })
$skip  = @($results | Where-Object { $_.Action -eq 'SkippedLocked' })
$fail  = @($results | Where-Object { $_.Action -eq 'DeleteFailed' })
$total = $results.Count

$wouldBytes = ($would | ForEach-Object { Get-Bytes $_ } | Measure-Object -Sum).Sum
$delBytes   = ($del   | ForEach-Object { Get-Bytes $_ } | Measure-Object -Sum).Sum

Write-Host "`n=== Helper extra summary ===" -ForegroundColor DarkCyan
Write-Host ("  Total profiles not used for {0} days: {1}" -f $DaysOld, $total)
Write-Host ("  Would delete (dry run): {0}" -f $would.Count)
Write-Host ("  Deleted:                {0}" -f $del.Count)
Write-Host ("  Skipped (locked):       {0}" -f $skip.Count)
Write-Host ("  Delete failures:        {0}" -f $fail.Count)
if ($wouldBytes -gt 0) { Write-Host ("  Estimated space saved (dry run): {0}" -f (Format-Bytes $wouldBytes)) }
if ($delBytes -gt 0)   { Write-Host ("  Actual space saved:               {0}" -f (Format-Bytes $delBytes)) }

Write-Host "`nPer-share totals (helper):" -ForegroundColor DarkCyan
$results | Group-Object Share | ForEach-Object {
    $wBytes=0;$dBytes=0;$wCount=0;$dCount=0
    foreach ($r in $_.Group) {
        $b = Get-Bytes $r
        if ($r.Action -eq 'WouldDelete') { $wBytes += $b; $wCount++ }
        if ($r.Action -eq 'Deleted')     { $dBytes += $b; $dCount++ }
    }
    Write-Host ("  Share: {0}" -f $_.Name)
    Write-Host ("    Would delete: {0} files, {1}" -f $wCount, (Format-Bytes $wBytes))
    Write-Host ("    Deleted:      {0} files, {1}" -f $dCount, (Format-Bytes $dBytes))
}
