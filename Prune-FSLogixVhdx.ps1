[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string[]] $Paths = @(
        '\\saavduksg2v.file.core.windows.net\fslogix\Profiles',
        '\\SAAVDUKWG2V.File.core.windows.net\fslogix\Profiles_W11'
    ),
    [int] $DaysOld = 60,
    [int] $RecurseDepth = 5,
    [string[]] $IncludePatterns = @('*.vhdx'),
    [string] $LogPath,
    [switch] $SkipIfLocked,
    [string[]] $UserFilter
)

begin {
    $ErrorActionPreference = 'Stop'
    $cutoff = (Get-Date).AddDays(-[double]$DaysOld)

    function Test-FileLocked {
        param([Parameter(Mandatory)][string]$Path)
        try {
            $f = [System.IO.File]::Open($Path,'Open','ReadWrite','None')
            $f.Close(); return $false
        } catch { return $true }
    }

    function Format-Bytes {
        param([double]$Bytes)
        if ($Bytes -ge 1TB) { "{0:N2} TB" -f ($Bytes / 1TB) }
        elseif ($Bytes -ge 1GB) { "{0:N2} GB" -f ($Bytes / 1GB) }
        elseif ($Bytes -ge 1MB) { "{0:N2} MB" -f ($Bytes / 1MB) }
        else { "{0:N0} B" -f $Bytes }
    }

    $results    = New-Object System.Collections.Generic.List[object]
    $shareStats = @{}   # key = share path, value = stats object (user folder count + tallies)
}

process {
    foreach ($root in $Paths) {
        if (-not (Test-Path $root)) {
            Write-Warning ("Path not found or inaccessible: {0}" -f $root)
            continue
        }

        # Count immediate user folders (optionally filtered)
        try {
            $dirs = Get-ChildItem -Path $root -Directory -ErrorAction Stop
            if ($UserFilter) {
                $tmp = @()
                foreach ($p in $UserFilter) { $tmp += $dirs | Where-Object { $_.Name -like $p } }
                $dirs = $tmp | Select-Object -Unique
            }
            $userFolderCount = ($dirs | Measure-Object).Count
        } catch {
            Write-Warning ("Failed to enumerate immediate folders under {0}: {1}" -f $root, $_.Exception.Message)
            $userFolderCount = $null
        }

        # Init per-share stats
        $shareStats[$root] = [pscustomobject]@{
            UserFolderCount = $userFolderCount
            WouldDeleteCount = 0
            WouldDeleteBytes = 0
            DeletedCount     = 0
            DeletedBytes     = 0
        }

        # Build search scopes (optionally per-user filtered)
        $userFolders = @()
        if ($UserFilter) {
            try {
                $immediate = Get-ChildItem -Path $root -Directory -ErrorAction Stop
                foreach ($p in $UserFilter) { $userFolders += $immediate | Where-Object { $_.Name -like $p } }
                $userFolders = $userFolders | Select-Object -Unique
            } catch {
                Write-Warning ("Failed to enumerate user folders under {0}: {1}" -f $root, $_.Exception.Message)
                continue
            }
        }
        $searchScopes = if ($userFolders) { $userFolders.FullName } else { @($root) }

        foreach ($scope in $searchScopes) {
            foreach ($pattern in $IncludePatterns) {
                try {
                    # Try -Depth (PS 7+), fall back for PS 5.1
                    $files = Get-ChildItem -Path $scope -Filter $pattern -File -ErrorAction Stop -Recurse -Depth $RecurseDepth
                } catch {
                    try { $files = Get-ChildItem -Path $scope -Filter $pattern -File -ErrorAction Stop -Recurse }
                    catch {
                        Write-Warning ("Failed to enumerate files in {0}: {1}" -f $scope, $_.Exception.Message)
                        continue
                    }
                }

                foreach ($f in $files) {
                    if ($f.LastWriteTime -ge $cutoff) { continue }

                    $locked = $false
                    if ($SkipIfLocked) { $locked = Test-FileLocked -Path $f.FullName }

                    $bytes = [double]$f.Length
                    $rec = [pscustomobject]@{
                        Share         = $root
                        Path          = $f.FullName
                        Name          = $f.Name
                        Directory     = $f.DirectoryName
                        LastWriteTime = $f.LastWriteTime
                        AgeDays       = [math]::Round(((Get-Date) - $f.LastWriteTime).TotalDays, 1)
                        SizeGB        = [math]::Round($bytes / 1GB, 3)
                        Bytes         = $bytes
                        Locked        = $locked
                        Action        = 'Candidate'
                    }

                    if ($locked) {
                        $rec.Action = 'SkippedLocked'
                    } else {
                        if ($PSCmdlet.ShouldProcess($f.FullName,'Delete')) {
                            try {
                                Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                                $rec.Action = 'Deleted'
                            } catch {
                                $rec.Action = 'DeleteFailed'
                                $rec | Add-Member NoteProperty Error $_.Exception.Message
                            }
                        } else {
                            $rec.Action = 'WouldDelete'
                        }
                    }

                    $results.Add($rec) | Out-Null
                }
            }
        }
    }
}

end {
    if ($results.Count -eq 0) {
        Write-Host ("No stale VHDX files found older than {0} days." -f $DaysOld)
        return
    }

    # ---- DISPLAY TABLE (to screen only) ----
    $results |
      Sort-Object Share, Directory, Name |
      Select-Object Share, @{n='Size';e={ "{0:N2} GB" -f $_.SizeGB }}, AgeDays, LastWriteTime, Action, Path |
      Format-Table -AutoSize | Out-Host

    # ---- Global tallies ----
    $deleted      = $results | Where-Object { $_.Action -eq 'Deleted' }
    $wouldDelete  = $results | Where-Object { $_.Action -eq 'WouldDelete' }
    $skipped      = $results | Where-Object { $_.Action -eq 'SkippedLocked' }
    $failed       = $results | Where-Object { $_.Action -eq 'DeleteFailed' }

    $total        = $results.Count
    $delBytes     = ($deleted     | Measure-Object -Property Bytes -Sum).Sum
    $wouldBytes   = ($wouldDelete | Measure-Object -Property Bytes -Sum).Sum

    # Overall user-folder denominator & percentage
    $overallFolders = 0
    foreach ($root in $shareStats.Keys) {
        $uc = $shareStats[$root].UserFolderCount
        if ($null -ne $uc) { $overallFolders += [int]$uc }
    }
    $overallPerc = $null
    if ($overallFolders -gt 0) {
        $overallPerc = [math]::Round(($total / $overallFolders) * 100, 1)
    }

    # ---- Per-share tallies ----
    foreach ($root in $shareStats.Keys) {
        $shareWould = $wouldDelete | Where-Object { $_.Share -eq $root }
        $shareDel   = $deleted     | Where-Object { $_.Share -eq $root }
        $shareStats[$root].WouldDeleteCount = ($shareWould | Measure-Object).Count
        $shareStats[$root].WouldDeleteBytes = ($shareWould | Measure-Object -Property Bytes -Sum).Sum
        $shareStats[$root].DeletedCount     = ($shareDel   | Measure-Object).Count
        $shareStats[$root].DeletedBytes     = ($shareDel   | Measure-Object -Property Bytes -Sum).Sum
    }

    # ---- Console summary ----
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Cyan
    $overallSuffix = ""
    if ($overallPerc -ne $null) { $overallSuffix = (" (of {0} = {1}%)" -f $overallFolders, $overallPerc) }
    Write-Host ("  Total profiles not used for {0} days: {1}{2}" -f $DaysOld, $total, $overallSuffix)
    Write-Host ("  Deleted:                {0}" -f ($deleted.Count))
    Write-Host ("  Would delete (dry run): {0}" -f ($wouldDelete.Count))
    Write-Host ("  Skipped (locked):       {0}" -f ($skipped.Count))
    Write-Host ("  Delete failures:        {0}" -f ($failed.Count))
    if ($wouldDelete.Count -gt 0) { Write-Host ("  Estimated space saved (dry run): {0}" -f (Format-Bytes $wouldBytes)) }
    if ($deleted.Count   -gt 0)   { Write-Host ("  Actual space saved:               {0}" -f (Format-Bytes $delBytes)) }

    Write-Host "`nPer-share totals:" -ForegroundColor Cyan
    foreach ($root in $shareStats.Keys) {
        $s = $shareStats[$root]
        $userCountTxt = if ($s.UserFolderCount -ne $null) { $s.UserFolderCount } else { 'N/A' }

        # Percentages per share (guard for null/zero)
        $wdPct = $null; $delPct = $null
        if ($s.UserFolderCount -ne $null -and $s.UserFolderCount -gt 0) {
            $wdPct  = [math]::Round(($s.WouldDeleteCount / $s.UserFolderCount) * 100, 1)
            $delPct = [math]::Round(($s.DeletedCount     / $s.UserFolderCount) * 100, 1)
        }
        $wdPctText  = if ($wdPct  -ne $null) { ("{0}%" -f $wdPct) }  else { "N/A" }
        $delPctText = if ($delPct -ne $null) { ("{0}%" -f $delPct) } else { "N/A" }

        Write-Host ("  Share: {0}" -f $root)
        Write-Host ("    Total user folders : {0}" -f $userCountTxt)
        Write-Host ("    Would delete        : {0} files ({1} of user folders), {2}" -f $s.WouldDeleteCount, $wdPctText,  (Format-Bytes $s.WouldDeleteBytes))
        Write-Host ("    Actually deleted    : {0} files ({1} of user folders), {2}" -f $s.DeletedCount,     $delPctText, (Format-Bytes $s.DeletedBytes))
    }

    Write-Host "`nTop 10 largest stale profiles:" -ForegroundColor Cyan
    $results |
      Where-Object { $_.Action -in 'WouldDelete','Deleted' } |
      Sort-Object Bytes -Descending |
      Select-Object -First 10 @{n='Size';e={ Format-Bytes $_.Bytes }},
                            Name, Directory, AgeDays, Action |
      Format-Table -AutoSize | Out-Host

    # Optional CSV logging
    if ($LogPath) {
        try {
            $dir = Split-Path -Path $LogPath -Parent
            if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
            $results | ConvertTo-Csv -NoTypeInformation | Set-Content -LiteralPath $LogPath -Encoding UTF8
            Write-Host ("`nDetailed CSV written to: {0}" -f $LogPath) -ForegroundColor Gray
        } catch {
            Write-Warning ("Failed to write log to {0}: {1}" -f $LogPath, $_.Exception.Message)
        }
    }

    # Return raw objects so the helper can compute/augment summaries
    $results
}
