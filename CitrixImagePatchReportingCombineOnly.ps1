[CmdletBinding()]
param(
    [string]$ShareRoot = "\\stprodukstcitrix01.file.core.windows.net\profiles\Citrix Reporting\Image Updates",
    [string]$Period    = (Get-Date).ToString("MMMM yyyy")
)

# ============================================================
# CITRIX MONTHLY PATCHING REPORT - COMBINE ONLY
#
# DOES NOT SCAN THE LOCAL MACHINE
#
# Reads existing repository CSVs:
#
#   SERVER-AppStatus-YYYY-MM-DD.csv
#   SERVER-OSInfo-YYYY-MM-DD.csv
#   SERVER-OSUpdates-YYYY-MM-DD.csv
#
# Creates one combined HTML report.
#
# If multiple reports exist for the same server, the newest
# dated report is used.
# ============================================================


# ============================================================
# MASTER IMAGE MAPPING
# ============================================================

$MasterMap = @(

    [pscustomobject]@{
        Master      = 'UKST1MICTXMI1D'
        Environment = 'Dev'
        Citrix      = 'Windows-10-MS-DEVL'
    }

    [pscustomobject]@{
        Master      = 'UKST1MICTXMI1U'
        Environment = 'CTE'
        Citrix      = 'Windows-10-MS-Test'
    }

    [pscustomobject]@{
        Master      = 'UKST1MICTXMI2D'
        Environment = 'Dev'
        Citrix      = 'Persistent DEV'
    }

    [pscustomobject]@{
        Master      = 'UKST1MICTXMI2U'
        Environment = 'CTE'
        Citrix      = 'Persistent CTE'
    }

    [pscustomobject]@{
        Master      = 'UKST1MICTXMI3U'
        Environment = 'Non prod'
        Citrix      = 'Windows-10-MS-NonProd-Standard'
    }

    [pscustomobject]@{
        Master      = 'UKST1MICTXMI4C'
        Environment = 'Admin tools'
        Citrix      = 'Windows-10-MS-Prod-Access'
    }

    [pscustomobject]@{
        Master      = 'UKST1MICTXMI5'
        Environment = 'Production'
        Citrix      = 'Windows-10-MS-Prod-Standard'
    }
)


# ============================================================
# CITRIX MACHINE CATALOGUE MAPPING
# ============================================================

$CatalogMap = @(

    # ========================================================
    # MI4C - Admin / Specialist applications
    # ========================================================

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South 3E Azure DR Prod'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South AuditTrack Prod'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South DEVQA Prod'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South Ent Apps Prod'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South HR Systems Prod'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South iManage Prod'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South 3E Azure DR UAT'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South AuditTrack UAT'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South DEVQA UAT'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South Ent Apps UAT'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South HR Systems UAT'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South iManage UAT'
        Master = 'UKST1MICTXMI4C'
        Type   = 'Multi Session'
    }


    # ========================================================
    # MI5 - Production Standard
    # ========================================================

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South Standard Prod'
        Master = 'UKST1MICTXMI5'
        Type   = 'Multi Session'
    }


    # ========================================================
    # MI1U - CTE Multi Session
    # ========================================================

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South CTE'
        Master = 'UKST1MICTXMI1U'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South CTE - UAT'
        Master = 'UKST1MICTXMI1U'
        Type   = 'Multi Session'
    }


    # ========================================================
    # MI1D - DEV Multi Session
    # ========================================================

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South DEV'
        Master = 'UKST1MICTXMI1D'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South DEV - UAT'
        Master = 'UKST1MICTXMI1D'
        Type   = 'Multi Session'
    }


    # ========================================================
    # MI3U - Non-Prod Standard
    # ========================================================

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South Standard Non-Prod'
        Master = 'UKST1MICTXMI3U'
        Type   = 'Multi Session'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 Multi Session UK South Standard Non-Prod IT UAT'
        Master = 'UKST1MICTXMI3U'
        Type   = 'Multi Session'
    }


    # ========================================================
    # MI2D - DEV PERSISTENT
    # ========================================================

    [pscustomobject]@{
        Name   = 'Windows 10 SU UK South Persistent DEV'
        Master = 'UKST1MICTXMI2D'
        Type   = 'Persistent'
    }

    [pscustomobject]@{
        Name   = 'Windows 10 SU UK South Persistent DEV 3rd Party Access'
        Master = 'UKST1MICTXMI2D'
        Type   = 'Persistent'
    }


    # ========================================================
    # MI2U - CTE PERSISTENT
    # ========================================================

    [pscustomobject]@{
        Name   = 'Windows 10 SU UK South Persistent CTE'
        Master = 'UKST1MICTXMI2U'
        Type   = 'Persistent'
    }
)


# ============================================================
# PATHS
# ============================================================

$monthFolder = Join-Path $ShareRoot $Period

if (-not (Test-Path $monthFolder)) {

    throw "Repository folder does not exist: $monthFolder"
}


Write-Host ""
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host " Citrix Monthly Patching Report - Combine Only" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Repository : $monthFolder"
Write-Host "Period     : $Period"
Write-Host ""


# ============================================================
# ROBUST CSV IMPORT
# ============================================================

function Import-CsvRobust {

    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [int]$Retries = 5,

        [int]$DelayMs = 500
    )


    for ($i = 1; $i -le $Retries; $i++) {

        try {

            return @(
                Import-Csv -Path $Path -ErrorAction Stop
            )

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


# ============================================================
# DATE CONVERSION
# ============================================================

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
        'yyyy-MM-dd HH:mm:ss',
        'yyyy-MM-dd'

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


# ============================================================
# FIND LATEST REPORT PER SERVER
# ============================================================

function Get-LatestReportFiles {

    param(

        [Parameter(Mandatory)]
        [string]$Folder,

        [Parameter(Mandatory)]
        [string]$ReportType
    )


    $files = @(

        Get-ChildItem `
            -Path $Folder `
            -Filter "*-$ReportType-*.csv" `
            -File `
            -ErrorAction SilentlyContinue
    )


    if (-not $files) {

        return @()
    }


    $parsed = foreach ($file in $files) {


        # Expected filename:
        #
        # UKST1MICTXMI5-AppStatus-2026-09-14.csv
        #
        # UKST1MICTXMI5-OSInfo-2026-09-14.csv
        #
        # UKST1MICTXMI5-OSUpdates-2026-09-14.csv


        if (
            $file.Name -match
            "^(?<Computer>.+)-$ReportType-(?<Date>\d{4}-\d{2}-\d{2})\.csv$"
        ) {

            $reportDate = $null


            try {

                $reportDate = [datetime]::ParseExact(
                    $Matches.Date,
                    'yyyy-MM-dd',
                    [System.Globalization.CultureInfo]::InvariantCulture
                )

            }
            catch {}


            [pscustomobject]@{

                Computer   = $Matches.Computer

                ReportDate = $reportDate

                File       = $file
            }
        }
    }


    # Select newest dated report per server

    $latest = foreach (
        $group in (
            $parsed |
            Group-Object Computer
        )
    ) {

        $group.Group |

            Sort-Object `
                @{
                    Expression = 'ReportDate'
                    Descending = $true
                },
                @{
                    Expression = { $_.File.LastWriteTime }
                    Descending = $true
                } |

            Select-Object -First 1
    }


    return @($latest)
}


# ============================================================
# HTML GENERATOR
# ============================================================

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


# ============================================================
# CSS
# ============================================================

$css = @"
<style>

body {
    font-family: Segoe UI, Arial, sans-serif;
    margin: 24px;
    background: #f8fafc;
    color: #111827;
}

h1 {
    margin: 0 0 18px 0;
}

h3 {
    margin: 18px 0 10px 0;
    color: #374151;
}

.panel {
    background: #ffffff;
    border: 1px solid #dcdfe4;
    border-radius: 10px;
    padding: 18px;
    margin: 18px 0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}

.summary {
    background: #ffffff;
    border: 1px solid #dcdfe4;
    border-radius: 10px;
    padding: 16px;
    margin-bottom: 20px;
}

.tag {
    display: inline-block;
    background: #e5e7eb;
    color: #111827;
    padding: 3px 10px;
    border-radius: 999px;
    font-size: 12px;
    margin-left: 7px;
    margin-top: 4px;
}

.small {
    color: #6b7280;
    font-size: 12px;
    margin-top: 16px;
}

table {
    border-collapse: collapse;
    width: 100%;
    margin-bottom: 20px;
}

th,
td {
    border: 1px solid #d1d5db;
    padding: 8px;
    text-align: left;
    background: #ffffff;
}

th {
    background: #f3f4f6;
    font-weight: 600;
}

.ok {
    color: #15803d;
    font-weight: 600;
}

.bad {
    color: #dc2626;
    font-weight: 600;
}

.info {
    color: #2563eb;
    font-weight: 600;
}

.os-in-month {
    color: #15803d;
    font-weight: 600;
}

.headerline {
    font-size: 18px;
    font-weight: 700;
}

.warn {
    color: #dc2626;
    font-weight: 600;
    margin-top: 6px;
}

.catlabel {
    font-weight: 600;
}

.catalogue-type {
    margin-top: 10px;
    font-weight: 600;
}

ul {
    margin: 5px 0 15px 22px;
}

</style>
"@


# ============================================================
# REPORT MONTH
# ============================================================

    try {

        $repMonth = [datetime]::ParseExact(
            $Period,
            'MMMM yyyy',
            [System.Globalization.CultureInfo]::InvariantCulture
        )

    }
    catch {

        throw "Period must be in the format 'MMMM yyyy'. Example: September 2026"
    }


    $monthStart = Get-Date `
        -Year $repMonth.Year `
        -Month $repMonth.Month `
        -Day 1 `
        -Hour 0 `
        -Minute 0 `
        -Second 0


    $monthEnd = $monthStart.AddMonths(1)


# ============================================================
# DISCOVER SERVERS FROM REPOSITORY DATA
# ============================================================

    $computers = @(

        $AppData.Computer

        $OsData.Computer

        $OsInfoMap.Keys

    ) |

    Where-Object { $_ } |

    Select-Object -Unique |

    Sort-Object


# ============================================================
# EXPECTED MASTER COUNT
# ============================================================

    $allMasters = @(

        $Map |

        Select-Object -ExpandProperty Master -Unique |

        Sort-Object
    )


    $presentMasters = @(

        $computers |

        Where-Object {

            $allMasters -contains $_

        } |

        Sort-Object
    )


    $missingMasters = @(

        $allMasters |

        Where-Object {

            $presentMasters -notcontains $_

        } |

        Sort-Object
    )


    $countLine =
        "Master image count: $($presentMasters.Count) of $($allMasters.Count)"


    $missingLine = ""


    if ($missingMasters.Count -gt 0) {

        $missingLine =

            "No current repository report found for: " +

            ($missingMasters -join ', ')
    }


    $headerBlock = @"
<div class='summary'>

    <div class='headerline'>

        $countLine

    </div>

    $(if ($missingLine) {

        "<div class='warn'>$missingLine</div>"

    })

</div>
"@


# ============================================================
# CREATE SERVER SECTIONS
# ============================================================

    $sections = foreach ($comp in $computers) {


# ------------------------------------------------------------
# APPLICATION DATA
# ------------------------------------------------------------

        $apps = @(

            $AppData |

            Where-Object {

                $_.Computer -ieq $comp -and

                $_.Status -ne 'Application not installed'

            } |

            Sort-Object Application
        )


# ------------------------------------------------------------
# OS UPDATES THIS MONTH
# ------------------------------------------------------------

        $osMonth = @(

            $OsData |

            Where-Object {

                $_.Computer -ieq $comp -and

                $_.InstalledOn -and

                $_.InstalledOn -ge $monthStart -and

                $_.InstalledOn -lt $monthEnd

            } |

            Sort-Object InstalledOn -Descending
        )


# ------------------------------------------------------------
# THREE PREVIOUS OS UPDATES
# ------------------------------------------------------------

        $osPrev3 = @(

            $OsData |

            Where-Object {

                $_.Computer -ieq $comp -and

                $_.InstalledOn -and

                $_.InstalledOn -lt $monthStart

            } |

            Sort-Object InstalledOn -Descending |

            Select-Object -First 3
        )


# ------------------------------------------------------------
# MASTER IMAGE INFORMATION
# ------------------------------------------------------------

        $mapRows = @(

            $Map |

            Where-Object {

                $_.Master -ieq $comp
            }
        )


        $envs = (

            $mapRows.Environment |

            Where-Object { $_ } |

            Select-Object -Unique

        ) -join ', '


        $ctxs = (

            $mapRows.Citrix |

            Where-Object { $_ } |

            Select-Object -Unique

        ) -join ', '


        if (-not $envs) {

            $envs = '-'
        }


        if (-not $ctxs) {

            $ctxs = '-'
        }


# ------------------------------------------------------------
# OS INFORMATION
# ------------------------------------------------------------

        $osTag = ''


        if ($OsInfoMap.ContainsKey($comp)) {

            $oi = $OsInfoMap[$comp]

            $osParts = @()


            if (
                $oi.OSName -and
                $oi.OSName -ne '-'
            ) {

                $osParts += $oi.OSName
            }


            if (
                $oi.OSVersion -and
                $oi.OSVersion -ne '-'
            ) {

                $osParts += $oi.OSVersion
            }


            if (
                $oi.OSBuild -and
                $oi.OSBuild -ne '-'
            ) {

                $osParts += "(Build $($oi.OSBuild))"
            }


            if ($osParts.Count -gt 0) {

                $osTag = $osParts -join ' '
            }
        }


# ------------------------------------------------------------
# MACHINE CATALOGUES
# ------------------------------------------------------------

        $cats = @(

            foreach (
                $c in (
                    $CatalogMap |
                    Where-Object {
                        $_.Master -ieq $comp
                    }
                )
            ) {

                [pscustomobject]@{

                    Name = $c.Name

                    Type = $c.Type
                }
            }
        )


        $catTotal = $cats.Count


        $multiSession = @(

            $cats |

            Where-Object {

                $_.Type -eq 'Multi Session'

            } |

            Select-Object -ExpandProperty Name
        )


        $persistent = @(

            $cats |

            Where-Object {

                $_.Type -eq 'Persistent'

            } |

            Select-Object -ExpandProperty Name
        )


        $catHtml = @()


# ------------------------------------------------------------
# MULTI SESSION CATALOGUES
# ------------------------------------------------------------

        if ($multiSession.Count -gt 0) {

            $items = (

                $multiSession |

                ForEach-Object {

                    "<li>$_</li>"
                }

            ) -join "`n"


            $catHtml += @"
<div class='catalogue-type'>
Multi Session
</div>

<ul>
$items
</ul>
"@
        }


# ------------------------------------------------------------
# PERSISTENT CATALOGUES
# ------------------------------------------------------------

        if ($persistent.Count -gt 0) {

            $items = (

                $persistent |

                ForEach-Object {

                    "<li>$_</li>"
                }

            ) -join "`n"


            $catHtml += @"
<div class='catalogue-type'>
Persistent
</div>

<ul>
$items
</ul>
"@
        }


        if ($catHtml.Count -eq 0) {

            $catHtml = @(

                '<em>No catalogues mapped.</em>'
            )
        }


# ------------------------------------------------------------
# APPLICATION HTML ROWS
# ------------------------------------------------------------

        $appRows = @(

            foreach ($a in $apps) {


                $rowCss = $a.Css


                $latestCss =

                    if ($a.LatestCss) {

                        $a.LatestCss

                    }
                    else {

                        $rowCss
                    }


@"
<tr>

    <td class='$rowCss'>
        $($a.Application)
    </td>

    <td class='$rowCss'>
        $($a.Installed)
    </td>

    <td class='$latestCss'>
        $($a.Latest)
    </td>

    <td class='$rowCss'>
        $($a.Status)
    </td>

</tr>
"@

            }
        )


        if ($appRows.Count -eq 0) {

            $appRows = @(

                "<tr><td colspan='4'><em>No application report data found.</em></td></tr>"
            )
        }


# ------------------------------------------------------------
# OS UPDATES THIS MONTH
# ------------------------------------------------------------

        $osRowsMonth = @()


        foreach ($o in $osMonth) {


            $date = '-'


            if ($o.InstalledOn) {

                $date =

                    $o.InstalledOn.ToString(
                        'dd/MM/yyyy'
                    )
            }


            $osRowsMonth += @"
<tr>

    <td class='os-in-month'>
        $($o.KB)
    </td>

    <td class='os-in-month'>
        $($o.Description)
    </td>

    <td class='os-in-month'>
        $date
    </td>

</tr>
"@

        }


        if ($osRowsMonth.Count -eq 0) {

            $osRowsMonth = @(

                "<tr><td colspan='3'><em>No OS updates were installed in $Period.</em></td></tr>"
            )
        }


# ------------------------------------------------------------
# PREVIOUS THREE OS UPDATES
# ------------------------------------------------------------

        $osRowsPrev3 = @()


        foreach ($o in $osPrev3) {


            $date = '-'


            if ($o.InstalledOn) {

                $date =

                    $o.InstalledOn.ToString(
                        'dd/MM/yyyy'
                    )
            }


            $osRowsPrev3 += @"
<tr>

    <td>
        $($o.KB)
    </td>

    <td>
        $($o.Description)
    </td>

    <td>
        $date
    </td>

</tr>
"@

        }


        if ($osRowsPrev3.Count -eq 0) {

            $osRowsPrev3 = @(

                "<tr><td colspan='3'><em>No previous OS update data found.</em></td></tr>"
            )
        }


# ------------------------------------------------------------
# SERVER PANEL
# ------------------------------------------------------------

@"
<div class='panel'>


    <div class='headerline'>

        $comp


        <span class='tag'>

            Env: $envs

        </span>


        <span class='tag'>

            Citrix: $ctxs

        </span>


        <span class='tag'>

            Catalogues: $catTotal

        </span>


        $(if ($osTag) {

            "<span class='tag'>OS: $osTag</span>"

        })

    </div>


    <br/>


    <div>

        <span class='catlabel'>

            Citrix Machine catalogues:

        </span>


        <br/>


        $(($catHtml -join "`n"))

    </div>


    <h3>
        Applications
    </h3>


    <table>


        <thead>

            <tr>

                <th>
                    Application
                </th>

                <th>
                    Installed
                </th>

                <th>
                    Latest
                </th>

                <th>
                    Status
                </th>

            </tr>

        </thead>


        <tbody>

            $(($appRows -join "`n"))

        </tbody>


    </table>


    <h3>

        OS updates applied in $Period

    </h3>


    <table>


        <thead>

            <tr>

                <th>
                    KB
                </th>

                <th>
                    Description
                </th>

                <th>
                    Installed On
                </th>

            </tr>

        </thead>


        <tbody>

            $(($osRowsMonth -join "`n"))

        </tbody>


    </table>


    <h3>

        OS updates - recent 3 excluding this month

    </h3>


    <table>


        <thead>

            <tr>

                <th>
                    KB
                </th>

                <th>
                    Description
                </th>

                <th>
                    Installed On
                </th>

            </tr>

        </thead>


        <tbody>

            $(($osRowsPrev3 -join "`n"))

        </tbody>


    </table>


</div>
"@

    }


# ============================================================
# FINAL HTML
# ============================================================

$html = @"
<!DOCTYPE html>

<html>

<head>

<meta charset='utf-8'>

<title>
$Title
</title>

$css

</head>


<body>


<h1>
$Title
</h1>


$headerBlock


$(($sections -join "`n"))


<p class='small'>

Application information is taken from the latest
AppStatus CSV report found for each master image.

</p>


<p class='small'>

Operating system information is taken from the latest
OSInfo and OSUpdates CSV reports found for each master image.

</p>


<p class='small'>

Generated from repository data:

$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')

</p>


</body>

</html>
"@


    $html |

        Set-Content `
            -Path $OutPath `
            -Encoding UTF8


    return $OutPath
}


# ============================================================
# LOAD APPSTATUS REPORTS
# ============================================================

Write-Host ""
Write-Host "Finding latest AppStatus report for each master..." -ForegroundColor Cyan


$appReportFiles =

    Get-LatestReportFiles `
        -Folder $monthFolder `
        -ReportType 'AppStatus'


Write-Host "AppStatus reports found: $($appReportFiles.Count)"


$allApps = @()


foreach ($report in $appReportFiles) {


    Write-Host `

        "  $($report.Computer) -> $($report.File.Name)" `

        -ForegroundColor DarkGray


    $rows =

        Import-CsvRobust `
            -Path $report.File.FullName


    foreach ($row in $rows) {


        if (-not $row.Computer) {


            $row |

                Add-Member `
                    -NotePropertyName Computer `
                    -NotePropertyValue $report.Computer `
                    -Force
        }
    }


    $allApps += $rows
}


# ============================================================
# LOAD OS UPDATE REPORTS
# ============================================================

Write-Host ""
Write-Host "Finding latest OSUpdates report for each master..." -ForegroundColor Cyan


$osReportFiles =

    Get-LatestReportFiles `
        -Folder $monthFolder `
        -ReportType 'OSUpdates'


Write-Host "OSUpdates reports found: $($osReportFiles.Count)"


$allOs = @()


foreach ($report in $osReportFiles) {


    Write-Host `

        "  $($report.Computer) -> $($report.File.Name)" `

        -ForegroundColor DarkGray


    $rows =

        Import-CsvRobust `
            -Path $report.File.FullName


    foreach ($row in $rows) {


        if (-not $row.Computer) {


            $row |

                Add-Member `
                    -NotePropertyName Computer `
                    -NotePropertyValue $report.Computer `
                    -Force
        }


        if ($row.InstalledOn) {


            $row.InstalledOn =

                Convert-ToDate `
                    $row.InstalledOn
        }
    }


    $allOs += $rows
}


# ============================================================
# LOAD OSINFO REPORTS
# ============================================================

Write-Host ""
Write-Host "Finding latest OSInfo report for each master..." -ForegroundColor Cyan


$osInfoReportFiles =

    Get-LatestReportFiles `
        -Folder $monthFolder `
        -ReportType 'OSInfo'


Write-Host "OSInfo reports found: $($osInfoReportFiles.Count)"


$osInfoMap = @{}


foreach ($report in $osInfoReportFiles) {


    Write-Host `

        "  $($report.Computer) -> $($report.File.Name)" `

        -ForegroundColor DarkGray


    $rows =

        Import-CsvRobust `
            -Path $report.File.FullName


    foreach ($row in $rows) {


        if (-not $row.Computer) {


            $row |

                Add-Member `
                    -NotePropertyName Computer `
                    -NotePropertyValue $report.Computer `
                    -Force
        }


        $osInfoMap[$report.Computer] = $row
    }
}


# ============================================================
# SUMMARY
# ============================================================

Write-Host ""
Write-Host "Repository summary:" -ForegroundColor Cyan
Write-Host ""

Write-Host "  AppStatus : $($appReportFiles.Count)"
Write-Host "  OSUpdates : $($osReportFiles.Count)"
Write-Host "  OSInfo    : $($osInfoReportFiles.Count)"

Write-Host ""


# ============================================================
# VALIDATE DATA
# ============================================================

if (
    -not $allApps -or
    $allApps.Count -eq 0
) {

    Write-Warning "No AppStatus report data was found."

    Write-Warning "Expected filename format:"

    Write-Warning "UKST1MICTXMI5-AppStatus-2026-09-14.csv"

    return
}


# ============================================================
# CREATE COMBINED HTML
# ============================================================

$todayIso =

    Get-Date -Format 'yyyy-MM-dd'


$combinedHtml =

    Join-Path `
        $monthFolder `
        "Citrix-Monthly-Patching-Report-$todayIso.html"


Write-Host ""
Write-Host "Building combined HTML report..." -ForegroundColor Cyan


New-CombinedHtml `
    -AppData $allApps `
    -OsData $allOs `
    -OsInfoMap $osInfoMap `
    -Title "Citrix Monthly Patching Report ($Period)" `
    -OutPath $combinedHtml `
    -Map $MasterMap `
    -CatalogMap $CatalogMap `
    -Period $Period |

    Out-Null


# ============================================================
# COMPLETE
# ============================================================

Write-Host ""
Write-Host "====================================================" -ForegroundColor Green
Write-Host " REPORT COMPLETE" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green
Write-Host ""

Write-Host "Combined HTML:" -ForegroundColor Green

Write-Host $combinedHtml -ForegroundColor Yellow

Write-Host ""
