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
$HtmlPath = Join-Path $OutputFolder "${Machine}_FSLogix_Report_${Stamp}.html"

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


# Build a self-contained HTML report. All values are captured locally; no CDN or external assets.
function HtmlEncode([object]$Value) {
    if ($null -eq $Value) { return '—' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}
function Format-Metric([object]$Value, [string]$Unit = '') {
    if ($null -eq $Value -or [string]$Value -eq '') { return 'N/A' }
    return ('{0:N2}{1}' -f [double]$Value, $Unit)
}
function Get-ShareStatus([object]$Row) {
    if ($Row.FolderTimeouts -gt 0 -or $Row.FolderErrors -gt 0 -or $Row.TCP445Issues -gt 0) { return 'Issue detected' }
    if ($Row.Checks -eq 0) { return 'No access checks' }
    if ($Row.FolderSlow -gt 0) { return 'Slow access detected' }
    if ($Row.SMBSamples -eq 0) { return 'Access OK / no SMB data' }
    return 'Access checks passed'
}
function Get-StatusClass([string]$Status) {
    if ($Status -eq 'Issue detected') { return 'bad' }
    if ($Status -eq 'Slow access detected' -or $Status -eq 'Access OK / no SMB data' -or $Status -eq 'No access checks') { return 'warn' }
    return 'good'
}
function Get-LineChartSvg([object[]]$Rows, [string]$Field, [string]$Caption) {
    $W = 760; $H = 150; $L = 38; $R = 15; $T = 12; $B = 25
    $Rows = @($Rows | Where-Object { $null -ne $_.$Field } | Sort-Object Timestamp)
    if ($Rows.Count -eq 0) { return '<div class="empty">No active SMB data for this measurement.</div>' }
    $Max = [double](($Rows | Measure-Object -Property $Field -Maximum).Maximum)
    if ($Max -le 0) { $Max = 1 }
    $PlotW = $W - $L - $R; $PlotH = $H - $T - $B
    $Points = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $X = $L + ($i * $PlotW / [math]::Max(1,($Rows.Count - 1)))
        $Y = $T + $PlotH - ([double]$Rows[$i].$Field / $Max * $PlotH)
        $Points.Add(('{0:N1},{1:N1}' -f $X,$Y))
    }
    $Coords = $Points -join ' '
    $Peak = Format-Metric $Max ' ms'
    return ('<div class="chart-caption">{0} <span>Peak: {1}</span></div><svg class="graph" viewBox="0 0 {2} {3}" role="img" aria-label="{0}"><line x1="{4}" y1="{5}" x2="{6}" y2="{5}" stroke="#dde6f0"/><line x1="{4}" y1="{7}" x2="{6}" y2="{7}" stroke="#dde6f0"/><text x="2" y="{8}" fill="#66778d" font-size="11">{9:N1}</text><text x="2" y="{10}" fill="#66778d" font-size="11">0</text><polyline fill="none" stroke="#3877f6" stroke-width="2.5" stroke-linejoin="round" points="{11}"/></svg>' -f (HtmlEncode $Caption),$Peak,$W,$H,$L,$T,$($W-$R),$($T+$PlotH),$($T+9),$Max,$($T+$PlotH+4),$Coords)
}

$Cards = New-Object 'System.Collections.Generic.List[string]'
$ChartSections = New-Object 'System.Collections.Generic.List[string]'
foreach ($Row in $Summary) {
    $Share = @($Shares | Where-Object { $_.Name -eq $Row.Storage })[0]
    $Status = Get-ShareStatus $Row
    $Class = Get-StatusClass $Status
    $Rows = @($Samples | Where-Object { $_.Storage -eq $Row.Storage })
    $Card = @"
<article class="panel share">
  <div class="share-heading"><div><div class="eyebrow">$(HtmlEncode $Row.Storage)</div><h2>$(HtmlEncode $Share.Account)</h2></div><span class="pill $Class">$(HtmlEncode $Status)</span></div>
  <div class="path">$(HtmlEncode ('\\' + $Share.HostName + '\' + $Share.ShareName + '\' + $Share.Folder))</div>
  <div class="metrics">
    <div><span>Folder checks passed</span><strong>$($Row.FolderOK) / $($Row.Checks)</strong></div>
    <div><span>Folder timeouts</span><strong class="$(if($Row.FolderTimeouts){'red'}else{'green'})">$($Row.FolderTimeouts)</strong></div>
    <div><span>Folder errors</span><strong class="$(if($Row.FolderErrors){'red'}else{'green'})">$($Row.FolderErrors)</strong></div>
    <div><span>TCP 445 issues</span><strong class="$(if($Row.TCP445Issues){'red'}else{'green'})">$($Row.TCP445Issues)</strong></div>
    <div><span>Slow directory checks</span><strong>$($Row.FolderSlow)</strong></div>
    <div><span>Peak directory response</span><strong>$(if($null -ne $Row.MaxFolderResponseMs){$Row.MaxFolderResponseMs.ToString()+' ms'}else{'N/A'})</strong></div>
    <div><span>Average SMB read</span><strong>$(Format-Metric $Row.AvgReadMs ' ms')</strong></div>
    <div><span>Peak SMB read</span><strong>$(Format-Metric $Row.MaxReadMs ' ms')</strong></div>
    <div><span>Average SMB write</span><strong>$(Format-Metric $Row.AvgWriteMs ' ms')</strong></div>
    <div><span>Peak SMB write</span><strong>$(Format-Metric $Row.MaxWriteMs ' ms')</strong></div>
    <div><span>SMB samples</span><strong>$($Row.SMBSamples)</strong></div>
  </div>
</article>
"@
    $Cards.Add($Card)
    $ReadChart = Get-LineChartSvg $Rows 'ReadMs' 'Read latency (ms)'
    $WriteChart = Get-LineChartSvg $Rows 'WriteMs' 'Write latency (ms)'
    $ChartSections.Add("<article class='panel'><h2>$(HtmlEncode $Row.Storage) — SMB latency</h2>$ReadChart$WriteChart</article>")
}

$ProblemRows = @($AccessChecks | Where-Object { $_.TCP445 -ne 'OK' -or $_.FolderStatus -ne 'OK' } | Sort-Object Timestamp -Descending)
$Events = New-Object 'System.Collections.Generic.List[string]'
foreach ($Entry in ($ProblemRows | Select-Object -First 100)) {
    $Events.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td></tr>' -f (HtmlEncode $Entry.Timestamp),(HtmlEncode $Entry.Storage),(HtmlEncode $Entry.TCP445),(HtmlEncode $Entry.FolderStatus),(HtmlEncode $Entry.FolderResponseMs),(HtmlEncode $Entry.Detail)))
}
$EventHtml = if ($Events.Count) { $Events -join "`n" } else { '<tr><td colspan="6" class="empty">No failures or slow checks captured.</td></tr>' }
$CompletedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$StartedAt = $Stamp
$OverallProblems = @($Summary | Where-Object { (Get-ShareStatus $_) -eq 'Issue detected' }).Count
$OverallLabel = if ($OverallProblems) { "$OverallProblems share(s) with access issues" } else { 'No access failures observed' }
$OverallClass = if ($OverallProblems) { 'bad' } else { 'good' }
$TotalChecks = ($Summary | Measure-Object Checks -Sum).Sum
$TotalSamples = ($Summary | Measure-Object SMBSamples -Sum).Sum
$Html = @"
<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>FSLogix Azure Files diagnostics — $(HtmlEncode $Machine)</title>
<style>
:root{--bg:#f3f6fb;--ink:#1b2b43;--muted:#64758d;--line:#dce5f1;--brand:#275bd5;--surface:#fff}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 'Segoe UI',Arial,sans-serif}
main{max-width:1180px;margin:auto;padding:32px 22px 55px}.hero{background:linear-gradient(115deg,#102b5e,#255bc8);color:#fff;padding:28px 32px;border-radius:18px;box-shadow:0 8px 32px #102b5e18}
.hero h1{margin:4px 0 8px;font-size:28px}.hero p{margin:0;color:#dce9ff}.eyebrow{text-transform:uppercase;letter-spacing:1.5px;font-size:11px;font-weight:700;color:#7390bd}.hero .eyebrow{color:#bed5ff}
.topstats{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:18px 0 26px}.topstats>div,.panel{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:22px;box-shadow:0 3px 12px #1c376008}
.topstats strong{display:block;font-size:27px;letter-spacing:-.6px}.topstats span{color:var(--muted);font-size:12px}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}
h2{font-size:19px;margin:0 0 15px}.share-heading{display:flex;align-items:flex-start;justify-content:space-between;gap:12px}.share-heading h2{margin:0}.pill{white-space:nowrap;padding:5px 10px;border-radius:20px;font-size:12px;font-weight:700}.pill.good{background:#dff8e8;color:#176a3c}.pill.warn{background:#fff0c9;color:#906100}.pill.bad{background:#ffe1e1;color:#ad2020}
.path{font-family:Consolas,monospace;overflow-wrap:anywhere;color:#526c90;font-size:12px;margin:14px 0 20px;background:#f4f7fc;padding:10px;border-radius:7px}
.metrics{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.metrics>div{border-top:1px solid #ecf0f6;padding-top:10px}.metrics span{display:block;font-size:12px;color:var(--muted)}.metrics strong{font-size:20px}.red{color:#bd2929}.green{color:#17804a}
.section-title{margin:30px 0 14px;font-size:20px}.chart-caption{font-size:13px;font-weight:700;display:flex;justify-content:space-between;color:#435570;margin-top:22px}.chart-caption span{font-weight:400;color:var(--muted)}.graph{display:block;width:100%;height:auto;max-height:170px}
.table-scroll{overflow-x:auto}table{width:100%;border-collapse:collapse;text-align:left}th,td{padding:10px 12px;border-bottom:1px solid #e6edf6;font-size:13px}th{background:#f5f8fd;color:#49607b}td{overflow-wrap:anywhere}.empty{padding:18px;color:var(--muted);text-align:center}
.note{margin-top:24px;background:#edf3fc;border-left:4px solid var(--brand);padding:17px 20px;border-radius:8px;color:#425777}
.footer{margin-top:22px;color:var(--muted);font-size:12px}code{font-family:Consolas,monospace}
@media(max-width:850px){.grid,.topstats{grid-template-columns:1fr}.metrics{grid-template-columns:1fr}.hero{padding:22px}.share-heading{flex-wrap:wrap}}
@media print{body{background:white}main{max-width:100%;padding:0}.panel,.topstats>div,.hero{box-shadow:none;break-inside:avoid}}
</style></head><body><main>
<header class="hero"><div class="eyebrow">Infrastructure diagnostics / Azure Files</div><h1>FSLogix storage health report</h1><p>Server: $(HtmlEncode $Machine) &nbsp; · &nbsp; Completed: $(HtmlEncode $CompletedAt)</p></header>
<section class="topstats"><div><span>Access check result</span><strong><span class="pill $OverallClass">$(HtmlEncode $OverallLabel)</span></strong></div><div><span>Total folder checks</span><strong>$TotalChecks</strong></div><div><span>SMB performance samples</span><strong>$TotalSamples</strong></div></section>
<h2 class="section-title">Storage account comparison</h2><section class="grid">$($Cards -join "`n")</section>
<h2 class="section-title">SMB performance trends</h2><section class="grid">$($ChartSections -join "`n")</section>
<h2 class="section-title">Access problems and slow checks</h2><section class="panel table-scroll"><table><thead><tr><th>Time</th><th>Share</th><th>TCP 445</th><th>Directory status</th><th>Response ms</th><th>Details</th></tr></thead><tbody>$EventHtml</tbody></table></section>
<section class="note"><strong>How to interpret this:</strong> Directory probes and TCP 445 checks show sampled accessibility from this server; they do not test every FSLogix profile operation. SMB latency reflects local share-level counters, and a lack of observed errors does not rule out Azure-side IOPS throttling, intermittent timeouts or a problem on other VDAs. The latency graph uses sample order, not a uniform time axis. Compare with Azure Monitor and FSLogix profile logs before drawing a root-cause conclusion.</section>
<p class="footer">Read-only monitoring · Generated on $(HtmlEncode $Machine) · Source files: $(HtmlEncode $AccessCsv), $(HtmlEncode $SampleCsv), $(HtmlEncode $SummaryCsv)</p>
</main></body></html>
"@
[System.IO.File]::WriteAllText($HtmlPath,$Html,[System.Text.UTF8Encoding]::new($false))
Write-Host "HTML report:   $HtmlPath" -ForegroundColor Green
