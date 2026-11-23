# Read-only SMB performance collection for two Azure Files FSLogix shares.
# Run on a Citrix session host using Windows PowerShell 5.1 or PowerShell 7.
[CmdletBinding()]
param(
    [ValidateRange(1,120)][int]$DurationMinutes = 10,
    [ValidateRange(1,60)][int]$IntervalSeconds = 5,
    [string]$OutputFolder = 'C:\Temp\FSLogixDiagnostics'
)

$ErrorActionPreference = 'Stop'
$Shares = @(
    [pscustomobject]@{ Name='UK-G2V'; Account='saavdukwg2v'; HostName='saavdukwg2v.file.core.windows.net'; ShareName='fslogix'; ProfileFolder='Profiles_W11' },
    [pscustomobject]@{ Name='UK-SG2V'; Account='saavduksg2v'; HostName='saavduksg2v.file.core.windows.net'; ShareName='fslogix'; ProfileFolder='Profiles_W11' }
)

New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
$Machine = $env:COMPUTERNAME
$Stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$RawCsv = Join-Path $OutputFolder "${Machine}_SMB_samples_${Stamp}.csv"
$SummaryCsv = Join-Path $OutputFolder "${Machine}_SMB_summary_${Stamp}.csv"
$ConnectivityCsv = Join-Path $OutputFolder "${Machine}_SMB_connectivity_${Stamp}.csv"

Write-Host "Monitoring existing SMB connections from $Machine" -ForegroundColor Cyan
Write-Host "Duration: $DurationMinutes min; interval: $IntervalSeconds sec" 
$Connections = @(Get-SmbConnection -ErrorAction SilentlyContinue)
$Checks = foreach ($Share in $Shares) {
    $Related = @($Connections | Where-Object { $_.ServerName -ieq $Share.HostName -and $_.ShareName -ieq $Share.ShareName })
    try {
        $Dns = Resolve-DnsName -Name $Share.HostName -ErrorAction Stop | Where-Object { $_.Type -eq 'A' } | Select-Object -ExpandProperty IPAddress -First 1
    } catch { $Dns = $null }
    try { $Tcp445 = [bool](Test-NetConnection -ComputerName $Share.HostName -Port 445 -InformationLevel Quiet -WarningAction SilentlyContinue) }
    catch { $Tcp445 = $false }
    [pscustomobject]@{
        Server=$Machine; Storage=$Share.Name; UNC="\\$($Share.HostName)\$($Share.ShareName)\$($Share.ProfileFolder)"
        DNSAddress=$Dns; TCP445=$Tcp445; ActiveSMBConnections=$Related.Count
        SMBOpenHandles=if ($Related.Count) { ($Related | Measure-Object NumOpens -Sum).Sum } else { 0 }
    }
}
$Checks | Export-Csv -Path $ConnectivityCsv -NoTypeInformation
$Checks | Format-Table Storage,TCP445,ActiveSMBConnections,SMBOpenHandles,DNSAddress -AutoSize

# These counters observe existing SMB traffic only; do not create load on the file shares.
$Wanted = @(
    'Avg. sec/Read', 'Avg. sec/Write', 'Read Bytes/sec', 'Write Bytes/sec',
    'Read Requests/sec', 'Write Requests/sec', 'Credit Stalls/sec'
)
try {
    $Available = (Get-Counter -ListSet 'SMB Client Shares' -ErrorAction Stop).Counter
} catch {
    throw 'SMB Client Shares performance counters are unavailable on this machine. Run: Get-Counter -ListSet "SMB Client Shares"'
}
$Paths = @()
foreach ($CounterName in $Wanted) {
    $Pattern = "\SMB Client Shares(*)\$CounterName"
    if (@($Available | Where-Object { $_ -like "*\$CounterName" }).Count) { $Paths += $Pattern }
}
if (-not $Paths.Count) { throw 'No requested SMB client counters were found.' }
Write-Host "Using counters: $($Paths -join ', ')"

$SampleCount = [math]::Max(2, [int][math]::Ceiling(($DurationMinutes * 60.0) / $IntervalSeconds))
Write-Host "Collecting $SampleCount samples. Leave this PowerShell window open..." -ForegroundColor Cyan
try {
    $Capture = Get-Counter -Counter $Paths -SampleInterval $IntervalSeconds -MaxSamples $SampleCount -ErrorAction Stop
} catch {
    throw "Failed to collect SMB counters: $($_.Exception.Message)"
}

$Rows = [System.Collections.Generic.List[object]]::new()
foreach ($Snapshot in @($Capture)) {
    foreach ($Share in $Shares) {
        $InstanceSamples = @($Snapshot.CounterSamples | Where-Object {
            $_.InstanceName -like "*$($Share.Account)*" -and
            $_.InstanceName -like "*$($Share.ShareName)*"
        })
        if (-not $InstanceSamples.Count) { continue }
        foreach ($Group in ($InstanceSamples | Group-Object InstanceName)) {
            $Values = @{}
            foreach ($Sample in $Group.Group) {
                $Suffix = ($Sample.Path -split '\\')[-1]
                if ($null -ne $Sample.CookedValue) { $Values[$Suffix] = [double]$Sample.CookedValue }
            }
            $ReadMs  = if ($Values.ContainsKey('Avg. sec/Read')) { [math]::Round($Values['Avg. sec/Read']*1000,2) } else { $null }
            $WriteMs = if ($Values.ContainsKey('Avg. sec/Write')) { [math]::Round($Values['Avg. sec/Write']*1000,2) } else { $null }
            $ReadMBps = if ($Values.ContainsKey('Read Bytes/sec')) { [math]::Round($Values['Read Bytes/sec']/1MB,3) } else { $null }
            $WriteMBps = if ($Values.ContainsKey('Write Bytes/sec')) { [math]::Round($Values['Write Bytes/sec']/1MB,3) } else { $null }
            $ReadIOPS = if ($Values.ContainsKey('Read Requests/sec')) { [math]::Round($Values['Read Requests/sec'],2) } else { $null }
            $WriteIOPS = if ($Values.ContainsKey('Write Requests/sec')) { [math]::Round($Values['Write Requests/sec'],2) } else { $null }
            $Stalls = if ($Values.ContainsKey('Credit Stalls/sec')) { [math]::Round($Values['Credit Stalls/sec'],2) } else { $null }
            $Rows.Add([pscustomobject]@{
                Timestamp=$Snapshot.Timestamp; Server=$Machine; Storage=$Share.Name
                Account=$Share.Account; Share=$Share.ShareName; SMBInstance=$Group.Name
                ReadMs=$ReadMs; WriteMs=$WriteMs; ReadMBps=$ReadMBps; WriteMBps=$WriteMBps
                ReadIOPS=$ReadIOPS; WriteIOPS=$WriteIOPS; CreditStallsPerSec=$Stalls
            })
        }
    }
}
$Rows | Export-Csv -Path $RawCsv -NoTypeInformation
if ($Rows.Count -eq 0) {
    Write-Warning 'No SMB counter instances matched either Azure Files share. These accounts may not have active SMB traffic in this session/host; check Get-Counter -Counter "\SMB Client Shares(*)\Read Bytes/sec" -MaxSamples 2.'
    Write-Host "Connectivity output: $ConnectivityCsv"
    Write-Host "Empty sample CSV: $RawCsv"
    return
}

function Get-FieldStats {
    param([object[]]$Objects,[string]$Field)
    $Valid = @($Objects | Where-Object { $null -ne $_.$Field })
    if (-not $Valid.Count) { return @{ Average=$null; Maximum=$null } }
    $Measure = $Valid | Measure-Object -Property $Field -Average -Maximum
    return @{ Average=[math]::Round($Measure.Average,2); Maximum=[math]::Round($Measure.Maximum,2) }
}
$Summary = foreach ($Group in ($Rows | Group-Object Storage)) {
    $Read = Get-FieldStats $Group.Group 'ReadMs'
    $Write = Get-FieldStats $Group.Group 'WriteMs'
    $ReadRate = Get-FieldStats $Group.Group 'ReadMBps'
    $WriteRate = Get-FieldStats $Group.Group 'WriteMBps'
    $Stall = Get-FieldStats $Group.Group 'CreditStallsPerSec'
    [pscustomobject]@{
        Server=$Machine; Storage=$Group.Name; Samples=$Group.Count
        AvgReadMs=$Read.Average; MaxReadMs=$Read.Maximum
        AvgWriteMs=$Write.Average; MaxWriteMs=$Write.Maximum
        AvgReadMBps=$ReadRate.Average; AvgWriteMBps=$WriteRate.Average
        MaxCreditStallsPerSec=$Stall.Maximum
    }
}
$Summary | Export-Csv -Path $SummaryCsv -NoTypeInformation
Write-Host "`nResults (share-level SMB client measurements):" -ForegroundColor Green
$Summary | Format-Table -AutoSize
Write-Host "`nSamples:      $RawCsv"
Write-Host "Summary:      $SummaryCsv"
Write-Host "Connectivity: $ConnectivityCsv"
Write-Host 'Note: Client SMB counters do not prove Azure-side throttling or isolate Profiles_W11 folder I/O.'
