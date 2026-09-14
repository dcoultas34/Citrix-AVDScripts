[CmdletBinding()]
param(
    [string]$ShareRoot = "\\stprodukstcitrix01.file.core.windows.net\profiles\Citrix Reporting\Image Updates",
    [string]$Period    = (Get-Date).ToString("MMMM yyyy")
)

# ============================================================
# COMBINE-ONLY CITRIX MONTHLY PATCHING REPORT
#
# This script DOES NOT scan the local machine.
# It only reads existing CSV reports from the repository
# and creates a combined HTML report.
# ============================================================


# ---------------------- Master -> Citrix mapping ----------------------

$MasterMap = @(
    @{ Master='UKST1MICTXMI3U'; Environment='Non prod'  ; Citrix='Windows-10-MS-NonProd-Standard' }
    @{ Master='UKST1MICTXMI5' ; Environment='Production'; Citrix='Windows-10-MS-Prod-Standard' }
    @{ Master='UKST1MICTXMI4C'; Environment='Admin tools'; Citrix='Windows-10-MS-Prod-Access' }
    @{ Master='UKST1MICTXMI1D'; Environment='Dev'        ; Citrix='Windows-10-MS-DEVL' }
    @{ Master='UKST1MICTXMI1U'; Environment='CTE'        ; Citrix='Windows-10-MS-Test' }
) | ForEach-Object { [pscustomobject]$_ }


# ---------------------- Citrix Machine Catalogues -> Master mapping ----------------------

$CatalogMap = @(

    @{ Name='Windows 10 Multi Session UK South 3E Azure DR Prod' ; Master='UKST1MICTXMI4C'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South AuditTrack Prod'  ; Master='UKST1MICTXMI4C'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South DEVQA Prod'       ; Master='UKST1MICTXMI4C'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South Ent Apps Prod'    ; Master='UKST1MICTXMI4C'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South HR Systems Prod'  ; Master='UKST1MICTXMI4C'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South iManage Prod'     ; Master='UKST1MICTXMI4C'; Type='Multi Session' }

    @{ Name='Windows 10 Multi Session UK South Standard Prod'    ; Master='UKST1MICTXMI5' ; Type='Multi Session' }

    @{ Name='Windows 10 Multi Session UK South CTE'              ; Master='UKST1MICTXMI1U'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South CTE - UAT'        ; Master='UKST1MICTXMI1U'; Type='Multi Session' }

    @{ Name='Windows 10 Multi Session UK South 3E Azure DR UAT'  ; Master='UKST1MICTXMI4C'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South AuditTrack UAT'   ; Master='UKST1MICTXMI4C'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South DEVQA UAT'        ; Master='UKST1MICTXMI4C'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South Ent Apps UAT'     ; Master='UKST1MICTXMI4C'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South HR Systems UAT'   ; Master='UKST1MICTXMI4C'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South iManage UAT'      ; Master='UKST1MICTXMI4C'; Type='Multi Session' }

    @{ Name='Windows 10 Multi Session UK South DEV'              ; Master='UKST1MICTXMI1D'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South DEV - UAT'        ; Master='UKST1MICTXMI1D'; Type='Multi Session' }

    @{ Name='Windows 10 Multi Session UK South Standard Non-Prod'; Master='UKST1MICTXMI3U'; Type='Multi Session' }
    @{ Name='Windows 10 Multi Session UK South Standard Non-Prod IT UAT'; Master='UKST1MICTXMI3U'; Type='Multi Session' }

) | ForEach-Object { [pscustomobject]$_ }


# ---------------------- Repository ----------------------

$monthFolder = Join-Path $ShareRoot $Period

if (-not (Test-Path $monthFolder)) {
    throw "Repository folder does not exist: $monthFolder"
}

Write-Host ""
Write-Host "Citrix Monthly Report - Combine Only" -ForegroundColor Cyan
Write-Host "Repository : $monthFolder" -ForegroundColor DarkCyan
Write-Host "Period     : $Period" -ForegroundColor DarkCyan
Write-Host ""


# ---------------------- Robust CSV Import ----------------------

function Import-CsvRobust {

    param(
        [string]$Path,
        [int]$Retries = 5,
        [int]$DelayMs = 500
    )

    for ($i = 1; $i -le $Retries; $i++) {

        try {
            return @(Import-Csv -Path $Path -ErrorAction Stop)
        }
        catch {

            if ($i -eq $Retries) {
                Write-Warning "Failed to read '$Path': $($_.Exception.Message)"
                return @()
            }

            Start-Sleep -Milliseconds $DelayMs
        }
    }
}


# ---------------------- Date Conversion ----------------------

function Convert-ToDate {

    param($Value)

    if (-not $Value) {
        return $null
    }

    if ($Value -is [datetime]) {
        return $Value
    }

    foreach ($format in @(
        'dd/MM/yyyy HH:mm:ss',
        'dd/MM/yyyy',
        'MM/dd/yyyy HH:mm:ss',
        'MM/dd/yyyy',
        'yyyy-MM-ddTHH:mm:ss',
        'yyyy-MM-dd HH:mm:ss'
    )) {

        try {
            return [datetime]::ParseExact(
                "$Value",
                $format,
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
        catch {}
    }

    try {
        return [datetime]::Parse(
            "$Value",
            [System.Globalization.CultureInfo]::GetCultureInfo('en-GB')
        )
    }
    catch {}

    return $null
}


# ---------------------- HTML Generator ----------------------

function New-CombinedHtml {

    param(
        [array]$AppData,
        [array]$OsData,
        [hashtable]$OsInfoMap,
        [string]$Title,
        [string]$OutPath,
        [array]$Map,
        [array]$CatalogMap,
        [string]$Period
    )

    $css = @"
<style>

body {
    font-family: Segoe UI, Arial, sans-serif;
    margin: 24px;
}

h1 {
    margin: 0 0 16px 0;
}

h3 {
    margin: 16px 0 12px 0;
    color: #374151;
}

.panel {
    border: 1px solid #e0e0e0;
    border-radius: 10px;
    padding: 16px;
    margin: 18px 0;
}

.tag {
    display: inline-block;
    background: #e5e7eb;
    color: #111827;
    padding: 2px 10px;
    border-radius: 999px;
    font-size: 12px;
    margin-left: 8px;
}

.small {
    color: #6b7280;
    font-size: 12px;
    margin-top: 16px;
}

table {
    border-collapse: collapse;
    width: 100%;
    margin-bottom: 18px;
}

th, td {
    border: 1px solid #ddd;
    padding: 8px;
    text-align: left;
    background: #ffffff;
}

th {
    background: #f3f3f3;
}

.ok {
    color: #2e7d32;
    font-weight: 600;
}

.bad {
    color: #c62828;
    font-weight: 600;
}

.info {
    color: #1d4ed8;
    font-weight: 600;
}

.os-in-month {
    color: #2e7d32;
    font-weight: 600;
}

.headerline {
    font-size: 18px;
    font-weight: 700;
}

ul {
    margin: 6px 0 12px 18px;
}

.catlabel {
    font-weight: 600;
}

.warn {
    color: #c62828;
    font-weight: 600;
}

</style>
"@

    # Work out reporting month
    try {
        $repMonth = [datetime]::ParseExact(
            $Period,
            'MMMM yyyy',
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    }
    catch {
        throw "Period must be in the format 'MMMM yyyy', for example 'September 2026'."
    }

    $monthStart = Get-Date `
        -Year $repMonth.Year `
        -Month $repMonth.Month `
        -Day 1 `
        -Hour 0 `
        -Minute 0 `
        -Second 0

    $monthEnd = $monthStart.AddMonths(1)


    # Get all computers represented by reports
    $computers = @(
        $AppData.Computer
        $OsData.Computer
        $OsInfoMap.Keys
    ) |
    Where-Object { $_ } |
    Select-Object -Unique |
    Sort-Object


    # Expected / missing master images
    $allMasters = $Map |
        Select-Object -ExpandProperty Master -Unique |
        Sort-Object

    $presentMasters = @(
        $computers |
        Where-Object { $allMasters -contains $_ } |
        Sort-Object
    )

    $missingMasters = @(
        $allMasters |
        Where-Object { $presentMasters -notcontains $_ } |
        Sort-Object
    )

    $countLine = "Master image count: {0} of {1}" -f `
        $presentMasters.Count,
        $allMasters.Count

    $missingLine = ""

    if ($missingMasters.Count -gt 0) {
        $missingLine = "No repository report found for: " + ($missingMasters -join ', ')
    }


    $headerBlock = @"
<div class='panel'>
    <div class='headerline'>$countLine</div>
    $(if ($missingLine) { "<div class='warn'>$missingLine</div>" })
</div>
"@


    # ---------------------- Per Master Sections ----------------------

    $sections = foreach ($comp in $computers) {

        # Applications
        $apps = @(
            $AppData |
            Where-Object {
                $_.Computer -eq $comp -and
                $_.Status -ne 'Application not installed'
            } |
            Sort-Object Application
        )


        # OS updates from report month
        $osMonth = @(
            $OsData |
            Where-Object {
                $_.Computer -eq $comp -and
                $_.InstalledOn -and
                $_.InstalledOn -ge $monthStart -and
                $_.InstalledOn -lt $monthEnd
            } |
            Sort-Object InstalledOn -Descending
        )


        # Three most recent prior updates
        $osPrev3 = @(
            $OsData |
            Where-Object {
                $_.Computer -eq $comp -and
                $_.InstalledOn -and
                $_.InstalledOn -lt $monthStart
            } |
            Sort-Object InstalledOn -Descending |
            Select-Object -First 3
        )


        # Environment mapping
        $mapRows = @(
            $Map |
            Where-Object { $_.Master -ieq $comp }
        )

        $envs = (
            $mapRows.Environment |
            ForEach-Object { "$_".Trim() } |
            Select-Object -Unique
        ) -join ', '

        $ctxs = (
            $mapRows.Citrix |
            ForEach-Object { "$_".Trim() } |
            Select-Object -Unique
        ) -join ', '

        if (-not $envs) { $envs = '-' }
        if (-not $ctxs) { $ctxs = '-' }


        # OS information
        $osTag = ''

        if ($OsInfoMap.ContainsKey($comp)) {

            $oi = $OsInfoMap[$comp]

            $osTag = "OS: $($oi.OSName) $($oi.OSVersion) (Build $($oi.OSBuild))"
        }


        # Machine catalogues
        $cats = @(
            foreach ($c in ($CatalogMap | Where-Object { $_.Master -ieq $comp })) {

                $type = ($c.Type -as [string]).Trim()

                if ($type -match '^(multi\s*session)$') {
                    $type = 'Multi Session'
                }
                elseif ($type -match '^(persistent|persistant|single.*)$') {
                    $type = 'Persistent'
                }

                [pscustomobject]@{
                    Name = ($c.Name -as [string]).Trim()
                    Type = $type
                }
            }
        )

        $catTotal = $cats.Count

        $multiSession = @(
            $cats |
            Where-Object Type -eq 'Multi Session' |
            Select-Object -ExpandProperty Name
        )

        $persistent = @(
            $cats |
            Where-Object Type -eq 'Persistent' |
            Select-Object -ExpandProperty Name
        )


        $catHtml = @()

        if ($multiSession.Count -gt 0) {

            $catHtml += @"
<div class='catlabel'>Multi Session</div>
<ul>
$(
    (
        $multiSession |
        ForEach-Object { "<li>$_</li>" }
    ) -join "`n"
)
</ul>
"@
        }


        if ($persistent.Count -gt 0) {

            $catHtml += @"
<div class='catlabel'>Persistent</div>
<ul>
$(
    (
        $persistent |
        ForEach-Object { "<li>$_</li>" }
    ) -join "`n"
)
</ul>
"@
        }


        if ($catHtml.Count -eq 0) {
            $catHtml = @('<em>No catalogues mapped.</em>')
        }


        # ---------------------- Application Rows ----------------------

        $appRows = @(
            foreach ($a in $apps) {

                $rowCss = $a.Css

                $latestCss = if ($a.LatestCss) {
                    $a.LatestCss
                }
                else {
                    $rowCss
                }

                @"
<tr>
    <td class='$rowCss'>$($a.Application)</td>
    <td class='$rowCss'>$($a.Installed)</td>
    <td class='$latestCss'>$($a.Latest)</td>
    <td class='$rowCss'>$($a.Status)</td>
</tr>
"@
            }
        )

        if ($appRows.Count -eq 0) {
            $appRows = @(
                "<tr><td colspan='4'><em>No application report data found.</em></td></tr>"
            )
        }


        # ---------------------- OS Rows ----------------------

        $dateFormat = 'dd/MM/yyyy'

        $osRowsMonth = @()

        foreach ($o in $osMonth) {

            $date = '-'

            if ($o.InstalledOn) {
                $date = $o.InstalledOn.ToString($dateFormat)
            }

            $osRowsMonth += @"
<tr>
    <td class='os-in-month'>$($o.KB)</td>
    <td class='os-in-month'>$($o.Description)</td>
    <td class='os-in-month'>$date</td>
</tr>
"@
        }

        if ($osRowsMonth.Count -eq 0) {
            $osRowsMonth = @(
                "<tr><td colspan='3'><em>No OS updates were installed in $Period.</em></td></tr>"
            )
        }


        $osRowsPrev3 = @()

        foreach ($o in $osPrev3) {

            $date = '-'

            if ($o.InstalledOn) {
                $date = $o.InstalledOn.ToString($dateFormat)
            }

            $osRowsPrev3 += @"
<tr>
    <td>$($o.KB)</td>
    <td>$($o.Description)</td>
    <td>$date</td>
</tr>
"@
        }

        if ($osRowsPrev3.Count -eq 0) {
            $osRowsPrev3 = @(
                "<tr><td colspan='3'><em>No previous OS update data found.</em></td></tr>"
            )
        }


        # ---------------------- Master HTML Panel ----------------------

@"
<div class='panel'>

    <div class='headerline'>
        $comp
        <span class='tag'>Env: $envs</span>
        <span class='tag'>Citrix: $ctxs</span>
        <span class='tag'>Catalogues: $catTotal</span>
        $(if ($osTag) { "<span class='tag'>$osTag</span>" })
    </div>

    <br/>

    <div>
        <span class='catlabel'>Citrix Machine catalogues:</span>
        <br/>
        $(($catHtml -join "`n"))
    </div>

    <h3>Applications</h3>

    <table>
        <thead>
            <tr>
                <th>Application</th>
                <th>Installed</th>
                <th>Latest</th>
                <th>Status</th>
            </tr>
        </thead>

        <tbody>
            $(($appRows -join "`n"))
        </tbody>
    </table>


    <h3>OS updates applied in $Period</h3>

    <table>
        <thead>
            <tr>
                <th>KB</th>
                <th>Description</th>
                <th>Installed On</th>
            </tr>
        </thead>

        <tbody>
            $(($osRowsMonth -join "`n"))
        </tbody>
    </table>


    <h3>OS updates - recent 3 excluding this month</h3>

    <table>
        <thead>
            <tr>
                <th>KB</th>
                <th>Description</th>
                <th>Installed On</th>
            </tr>
        </thead>

        <tbody>
            $(($osRowsPrev3 -join "`n"))
        </tbody>
    </table>

</div>
"@
    }


    # ---------------------- Final HTML ----------------------

    $html = @"
<!DOCTYPE html>
<html>

<head>
    <meta charset='utf-8'>
    <title>$Title</title>
    $css
</head>

<body>

<h1>$Title</h1>

$headerBlock

$(($sections -join "`n"))

<p class='small'>
Application data is taken directly from the existing AppStatus CSV reports
stored in the repository.
</p>

<p class='small'>
OS update and operating system information is taken directly from the existing
OSUpdates and OSInfo CSV reports.
</p>

<p class='small'>
Generated from repository data: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
</p>

</body>

</html>
"@

    $html | Set-Content -Path $OutPath -Encoding UTF8

    return $OutPath
}


# ============================================================
# READ EXISTING REPOSITORY REPORTS
# ============================================================

Write-Host "Reading existing repository reports..." -ForegroundColor Cyan


# ---------------------- Application Reports ----------------------

$appFiles = @(
    Get-ChildItem `
        -Path $monthFolder `
        -Filter "* AppStatus *.csv" `
        -File `
        -ErrorAction SilentlyContinue
)

Write-Host "AppStatus files found : $($appFiles.Count)"


$allApps = @()

foreach ($file in $appFiles) {

    Write-Host "  Reading $($file.Name)" -ForegroundColor DarkGray

    $allApps += @(
        Import-CsvRobust -Path $file.FullName
    )
}


# ---------------------- OS Update Reports ----------------------

$osFiles = @(
    Get-ChildItem `
        -Path $monthFolder `
        -Filter "* OSUpdates *.csv" `
        -File `
        -ErrorAction SilentlyContinue
)

Write-Host "OSUpdates files found : $($osFiles.Count)"


$allOs = @()

foreach ($file in $osFiles) {

    Write-Host "  Reading $($file.Name)" -ForegroundColor DarkGray

    $allOs += @(
        Import-CsvRobust -Path $file.FullName
    )
}


# Convert imported InstalledOn strings back to DateTime
$allOs = @(
    $allOs |
    ForEach-Object {

        if ($_.InstalledOn) {
            $_.InstalledOn = Convert-ToDate $_.InstalledOn
        }

        $_
    }
)


# ---------------------- OS Information Reports ----------------------

$osInfoFiles = @(
    Get-ChildItem `
        -Path $monthFolder `
        -Filter "* OSInfo *.csv" `
        -File `
        -ErrorAction SilentlyContinue
)

Write-Host "OSInfo files found    : $($osInfoFiles.Count)"


$osInfoMap = @{}


foreach ($file in $osInfoFiles) {

    Write-Host "  Reading $($file.Name)" -ForegroundColor DarkGray

    $info = @(
        Import-CsvRobust -Path $file.FullName
    )

    foreach ($row in $info) {

        if ($row.Computer) {

            # Later file wins if duplicate reports exist
            $osInfoMap[$row.Computer] = $row
        }
    }
}


# ---------------------- Validation ----------------------

if ($allApps.Count -eq 0) {

    Write-Warning "No AppStatus CSV report data was found in:"
    Write-Warning $monthFolder

    return
}


# ============================================================
# CREATE COMBINED REPORT
# ============================================================

$todayIso = Get-Date -Format 'yyyy-MM-dd'

$combinedHtml = Join-Path `
    $monthFolder `
    "Citrix Monthly patching report $todayIso.html"


Write-Host ""
Write-Host "Building combined HTML..." -ForegroundColor Cyan


New-CombinedHtml `
    -AppData $allApps `
    -OsData $allOs `
    -OsInfoMap $osInfoMap `
    -Title "Citrix Monthly patching report ($Period)" `
    -OutPath $combinedHtml `
    -Map $MasterMap `
    -CatalogMap $CatalogMap `
    -Period $Period |
    Out-Null


Write-Host ""
Write-Host "==============================================" -ForegroundColor Green
Write-Host "Combined HTML report created successfully." -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Green
Write-Host ""
Write-Host $combinedHtml -ForegroundColor Yellow
Write-Host ""
