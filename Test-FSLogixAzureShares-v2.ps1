# Read-only monitoring of two Azure Files FSLogix stores.
# Windows PowerShell 5.1+, run on the Citrix VDA as a user/service identity
# with permission to list the profile folder (ideally an administrator).
[CmdletBinding()]
param(
    [ValidateRange(1,120)][int]$DurationMinutes = 10,
    [ValidateRange(2,60)][int]$IntervalSeconds = 5,
    [ValidateRange(2,30)][int]$ProbeTimeoutSeconds = 8,
    [ValidateRange(5,300)][int]$ProbeEverySeconds = 30,
    [ValidateRange(100,60000)][int]$SlowProbeMs = 2000,
    [string]$OutputFolder = 'C:\Temp\FSLogixDiagnostics'
)

$ErrorActionPreference = 'Stop'
$Shares = @(
    [pscustomobject]@{ Name='UK-G2V';  Account='saavdukwg2v'; HostName='saavdukwg2v.file.core.windows.net'; ShareName='fslogix'; Folder='Profiles_W11' },
    [pscustomobject]@{ Name='UK-SG2V'; Account='saavduksg2v'; HostName='saavduksg2v.file.core.windows.net'; ShareName='fslogix'; Folder='Profiles_W11' }
)

New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
$Machine = $env:COMPUTERNAME
$Stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$SampleCsv  = Join-Path $OutputFolder "${Machine}_SMB_samples_${Stamp}.csv"
$AccessCsv  = Join-Path $OutputFolder "${Machine}_Share_access_${Stamp}.csv"
$SummaryCsv = Join-Path $OutputFolder "${Machine}_Share_summary_${Stamp}.csv"

$Samples = New-Object 'System.Collections.Generic.List[object]'
$AccessChecks = New-Object 'System.Collections.Generic.List[object]'

# Directory enumeration checks an actual folder, not just whether TCP 445 is open.
# Each probe runs in a separate process with a timeout, so a hung UNC operation
# cannot stall the monitoring session indefinitely.
function Invoke-ShareProbe {
    param([object]$Share, [int]$TimeoutSeconds, [int]$SlowMs)
    $Path = "\\$($Share.HostName)\$($Share.ShareName)\$($Share.Folder)"
    $Start = Get-Date
    $Job = $null
    $State = 'UNKNOWN'; $Message = ''; $Elapsed = $null
    try {
        $Job = Start-Job -ArgumentList $Path -ScriptBlock {
            param($ProbePath)
            $Timer = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                # Enumerate only the first entry; no reads of profile content.
                $Enumerator = [System.IO.Directory]::EnumerateFileSystemEntries($ProbePath).GetEnumerator()
                try { $null = $Enumerator.MoveNext() }
                finally { if ($Enumerator -is [System.IDisposable]) { $Enumerator.Dispose() } }
                [pscustomobject]@{ Status='OK'; Ms=[math]::Round($Timer.Elapsed.TotalMilliseconds,0); Error='' }
            } catch {
                [pscustomobject]@{ Status='ACCESS_ERROR'; Ms=[math]::Round($Timer.Elapsed.TotalMilliseconds,0); Error=$_.Exception.Message }
            }
        }
        if (Wait-Job -Job $Job -Timeout $TimeoutSeconds) {
            $Result = Receive-Job -Job $Job -ErrorAction Stop | Select-Object -First 1
            if ($null -eq $Result) { $State = 'PROBE_ERROR'; $Message = 'No probe result returned' }
            else {
                $Elapsed = $Result.Ms; $State = $Result.Status; $Message = $Result.Error
                if ($State -eq 'OK' -and $Elapsed -ge $SlowMs) { $State = 'SLOW' }
            }
        } else {
            $State = 'TIMEOUT'; $Elapsed = $TimeoutSeconds * 1000
            $Message = "No folder response within $TimeoutSeconds seconds"
            Stop-Job -Job $Job -ErrorAction SilentlyContinue
        }
    } catch {
        $State = 'PROBE_ERROR'; $Message = $_.Exception.Message
    } finally {
        if ($null -ne $Job) { Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue }
    }
    [pscustomobject]@{
        Timestamp=$Start.ToString('s'); Server=$Machine; Storage=$Share.Name
        UNC=$Path; Status=$State; ResponseMs=$Elapsed; Detail=$Message
    }
}

function Test-SmbPort {
    param([string]$HostName,[int]$TimeoutMs=2000)
    $Client = New-Object System.Net.Sockets.TcpClient
    try {
        $Async = $Client.BeginConnect($HostName,445,$null,$null)
        if (-not $Async.AsyncWaitHandle.WaitOne($TimeoutMs)) { return 'TIMEOUT' }
        $Client.EndConnect($Async)
        return 'OK'
    } catch { return 'FAILED' }
    finally { $Client.Close() }
}

# Discover available SMB client counters. Some Windows versions lack counters.
$Wanted = @('Avg. sec/Read','Avg. sec/Write','Read Bytes/sec','Write Bytes/sec',
            'Read Requests/sec','Write Requests/sec','Credit Stalls/sec')
$CounterPaths = @()
try {
    $Available = (Get-Counter -ListSet 'SMB Client Shares' -ErrorAction Stop).Counter
    foreach ($Name in $Wanted) {
        if (@($Available | Where-Object { $_ -like "*\$Name" }).Count -gt 0) {
            $CounterPaths += "\SMB Client Shares(*)\$Name"
        }
    }
} catch {
    Write-Warning "SMB counters unavailable; access probes will still run: $($_.Exception.Message)"
}

Write-Host "`nFSLogix Azure Files monitor on $Machine" -ForegroundColor Cyan
Write-Host "Running $DurationMinutes minutes; sample every $IntervalSeconds sec; access probe every $ProbeEverySeconds sec"
$Shares | ForEach-Object { Write-Host "  $($_.Name): \\$($_.HostName)\$($_.ShareName)\$($_.Folder)" }
$End = (Get-Date).AddMinutes($DurationMinutes)
$NextProbe = Get-Date

while ((Get-Date) -lt $End) {
    $CycleStart = Get-Date
    if ($CycleStart -ge $NextProbe) {
        foreach ($Share in $Shares) {
            $Port = Test-SmbPort -HostName $Share.HostName
            $Probe = Invoke-ShareProbe -Share $Share -TimeoutSeconds $ProbeTimeoutSeconds -SlowMs $SlowProbeMs
            $AccessChecks.Add([pscustomobject]@{
                Timestamp=$Probe.Timestamp; Server=$Machine; Storage=$Share.Name; UNC=$Probe.UNC
                TCP445=$Port; FolderStatus=$Probe.Status; FolderResponseMs=$Probe.ResponseMs
                Detail=$Probe.Detail
            })
            $Colour = if ($Probe.Status -eq 'OK' -and $Port -eq 'OK') { 'Green' } else { 'Yellow' }
            Write-Host "[$($Probe.Timestamp)] $($Share.Name) TCP=$Port Folder=$($Probe.Status) Response=$($Probe.ResponseMs)ms" -ForegroundColor $Colour
        }
        $NextProbe = (Get-Date).AddSeconds($ProbeEverySeconds)
    }

    if ($CounterPaths.Count -gt 0) {
        try {
            # Two samples ensure a warm-up for rate counters. Keep the last sample.
            $Snapshot = @(Get-Counter -Counter $CounterPaths -SampleInterval 1 -MaxSamples 2 -ErrorAction Stop)[-1]
            foreach ($Share in $Shares) {
                $Match = @($Snapshot.CounterSamples | Where-Object {
                    $_.InstanceName -like "*$($Share.Account)*" -and
                    $_.InstanceName -like "*$($Share.ShareName)*"
                })
                foreach ($Group in ($Match | Group-Object InstanceName)) {
                    $Values = @{}
                    foreach ($Item in $Group.Group) {
                        $Name = ($Item.Path -split '\\')[-1]
                        $Values[$Name] = [double]$Item.CookedValue
                    }
                    $Samples.Add([pscustomobject]@{
                        Timestamp=$Snapshot.Timestamp.ToString('s'); Server=$Machine
                        Storage=$Share.Name; SMBInstance=$Group.Name
                        ReadMs=if($Values.ContainsKey('Avg. sec/Read')){[math]::Round($Values['Avg. sec/Read']*1000,2)}else{$null}
                        WriteMs=if($Values.ContainsKey('Avg. sec/Write')){[math]::Round($Values['Avg. sec/Write']*1000,2)}else{$null}
                        ReadMBps=if($Values.ContainsKey('Read Bytes/sec')){[math]::Round($Values['Read Bytes/sec']/1MB,3)}else{$null}
                        WriteMBps=if($Values.ContainsKey('Write Bytes/sec')){[math]::Round($Values['Write Bytes/sec']/1MB,3)}else{$null}
                        ReadIOPS=if($Values.ContainsKey('Read Requests/sec')){[math]::Round($Values['Read Requests/sec'],2)}else{$null}
                        WriteIOPS=if($Values.ContainsKey('Write Requests/sec')){[math]::Round($Values['Write Requests/sec'],2)}else{$null}
                        CreditStallsPerSec=if($Values.ContainsKey('Credit Stalls/sec')){[math]::Round($Values['Credit Stalls/sec'],2)}else{$null}
                    })
                }
            }
        } catch { Write-Warning "SMB sample failed: $($_.Exception.Message)" }
    }
    $Sleep = $IntervalSeconds - ((Get-Date)-$CycleStart).TotalSeconds
    if ($Sleep -gt 0) { Start-Sleep -Milliseconds ([int]($Sleep*1000)) }
}

# Always create files, including when there is no active SMB counter instance.
if ($Samples.Count) { $Samples | Export-Csv $SampleCsv -NoTypeInformation }
else { 'Timestamp,Server,Storage,SMBInstance,ReadMs,WriteMs,ReadMBps,WriteMBps,ReadIOPS,WriteIOPS,CreditStallsPerSec' | Set-Content $SampleCsv }
if ($AccessChecks.Count) { $AccessChecks | Export-Csv $AccessCsv -NoTypeInformation }
else { 'Timestamp,Server,Storage,UNC,TCP445,FolderStatus,FolderResponseMs,Detail' | Set-Content $AccessCsv }

$Summary = foreach ($Share in $Shares) {
    $Probes = @($AccessChecks | Where-Object { $_.Storage -eq $Share.Name })
    $Metrics = @($Samples | Where-Object { $_.Storage -eq $Share.Name })
    $Reads = @($Metrics | Where-Object { $null -ne $_.ReadMs })
    $Writes = @($Metrics | Where-Object { $null -ne $_.WriteMs })
    [pscustomobject]@{
        Server=$Machine; Storage=$Share.Name; Checks=$Probes.Count
        FolderOK=@($Probes | Where-Object FolderStatus -eq 'OK').Count
        FolderSlow=@($Probes | Where-Object FolderStatus -eq 'SLOW').Count
        FolderTimeouts=@($Probes | Where-Object FolderStatus -eq 'TIMEOUT').Count
        FolderErrors=@($Probes | Where-Object { $_.FolderStatus -in @('ACCESS_ERROR','PROBE_ERROR') }).Count
        TCP445Issues=@($Probes | Where-Object TCP445 -ne 'OK').Count
        MaxFolderResponseMs=if($Probes.Count){($Probes | Measure-Object FolderResponseMs -Maximum).Maximum}else{$null}
        SMBSamples=$Metrics.Count
        AvgReadMs=if($Reads.Count){[math]::Round(($Reads | Measure-Object ReadMs -Average).Average,2)}else{$null}
        MaxReadMs=if($Reads.Count){($Reads | Measure-Object ReadMs -Maximum).Maximum}else{$null}
        AvgWriteMs=if($Writes.Count){[math]::Round(($Writes | Measure-Object WriteMs -Average).Average,2)}else{$null}
        MaxWriteMs=if($Writes.Count){($Writes | Measure-Object WriteMs -Maximum).Maximum}else{$null}
    }
}
$Summary | Export-Csv $SummaryCsv -NoTypeInformation
Write-Host "`nSUMMARY" -ForegroundColor Cyan
$Summary | Format-Table -AutoSize
Write-Host "Access checks: $AccessCsv"
Write-Host "SMB samples:   $SampleCsv"
Write-Host "Summary:       $SummaryCsv"
Write-Host 'A TIMEOUT means this directory probe exceeded the configured limit; it does not prove Azure-side throttling.'
Write-Host 'An ACCESS_ERROR may reflect the identity running the script lacking folder permissions.'
Write-Host 'Counters are share-level, and this script does not inspect FSLogix profile logs or capture every SMB request timeout.'
