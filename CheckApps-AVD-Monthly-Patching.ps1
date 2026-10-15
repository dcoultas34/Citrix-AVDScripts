<#
.SYNOPSIS
    AVD Master Image Application / Patching Report

.DESCRIPTION
    Creates an HTML report containing:

      1. Operating System Information
      2. Recent Windows Updates
      3. Core Application Update Status
      4. Other Applications Installed or Updated This Month
      5. Application Inventory

    Designed for AVD master image reporting.

    Output:
        C:\Ultima\AVDPatchReports\<Month Year>
        \\transfer\transfer\Citrixapps\AV\<Month Year>

.NOTES
    Application data is read from both 64-bit and 32-bit
    Windows uninstall registry locations.

    Where duplicate registry entries exist for the same
    application/version, the entry containing a valid
    InstallDate is preferred.
#>

[CmdletBinding()]
param()

# ============================================================
# CONFIGURATION
# ============================================================

$ComputerName = $env:COMPUTERNAME
$Now          = Get-Date
$MonthName    = $Now.ToString("MMMM yyyy")
$DateStamp    = $Now.ToString("yyyy-MM-dd")
$RunTime      = $Now.ToString("dd/MM/yyyy HH:mm")

# ============================================================
# OUTPUT LOCATIONS
# ============================================================

$LocalRoot = "C:\Ultima\AVDPatchReports"
$ShareRoot = "\\transfer\transfer\Citrixapps\AV"

$LocalMonthFolder = Join-Path $LocalRoot $MonthName
$ShareMonthFolder = Join-Path $ShareRoot $MonthName

# Create local reporting folder
if (-not (Test-Path $LocalMonthFolder)) {

    New-Item `
        -Path $LocalMonthFolder `
        -ItemType Directory `
        -Force | Out-Null
}

# Test/create network reporting folder
$ShareAvailable = $false

try {

    if (-not (Test-Path $ShareMonthFolder)) {

        New-Item `
            -Path $ShareMonthFolder `
            -ItemType Directory `
            -Force `
            -ErrorAction Stop | Out-Null
    }

    $ShareAvailable = $true
}
catch {

    Write-Warning "Unable to access reporting share:"
    Write-Warning $ShareMonthFolder
    Write-Warning $_.Exception.Message
}

# ============================================================
# CORE APPLICATION UPDATE STATUS
# ============================================================

$CoreApplications = @(

    @{
        Name = "Microsoft 365 Apps"

        Matches = @(
            "Microsoft 365 Apps for business - en-gb",
            "Microsoft 365 Apps for business*",
            "Microsoft 365 Apps*"
        )
    },

    @{
        Name = "Adobe Acrobat"

        Matches = @(
            "Adobe Acrobat (64-bit)",
            "Adobe Acrobat*"
        )
    },

    @{
        Name = "Google Chrome"

        Matches = @(
            "Google Chrome"
        )
    },

    @{
        Name = "Microsoft Edge"

        Matches = @(
            "Microsoft Edge"
        )
    },

    @{
        Name = "Microsoft Power BI Desktop"

        Matches = @(
            "Microsoft Power BI Desktop (x64)",
            "Microsoft PowerBI Desktop (x64)",
            "Microsoft Power BI Desktop*",
            "Microsoft PowerBI Desktop*"
        )
    },

    @{
        Name = "Microsoft OneDrive"

        Matches = @(
            "Microsoft OneDrive"
        )
    },

    @{
        Name = "Mozilla Firefox"

        Matches = @(
            "Mozilla Firefox*"
        )
    }
)

# ============================================================
# APPLICATION INVENTORY
# ============================================================

$InventoryApps = @(

    @{
        Name = "8x8 Work"

        Matches = @(
            "8x8 Work*"
        )
    },

    @{
        Name = "Adobe Acrobat (64-bit)"

        Matches = @(
            "Adobe Acrobat (64-bit)",
            "Adobe Acrobat*"
        )
    },

    @{
        Name = "Foxit PDF Editor"

        Matches = @(
            "Foxit PDF Editor*"
        )
    },

    @{
        Name = "Infor Query & Analysis 11"

        Matches = @(
            "Infor Query & Analysis 11"
        )
    },

    @{
        Name = "Infor Query & Analysis 11 DataLink Server"

        Matches = @(
            "Infor Query & Analysis 11 DataLink Server*"
        )
    },

    @{
        Name = "Microsoft 365 Apps for business - en-gb"

        Matches = @(
            "Microsoft 365 Apps for business - en-gb",
            "Microsoft 365 Apps for business*"
        )
    },

    @{
        Name = "Microsoft Edge"

        Matches = @(
            "Microsoft Edge"
        )
    },

    @{
        Name = "Microsoft OneDrive"

        Matches = @(
            "Microsoft OneDrive"
        )
    },

    @{
        Name = "Microsoft PowerBI Desktop (x64)"

        Matches = @(
            "Microsoft Power BI Desktop (x64)",
            "Microsoft PowerBI Desktop (x64)",
            "Microsoft Power BI Desktop*",
            "Microsoft PowerBI Desktop*"
        )
    },

    @{
        Name = "Microsoft Project - en-us"

        Matches = @(
            "Microsoft Project - en-us*",
            "Microsoft Project*"
        )
    },

    @{
        Name = "Microsoft Report Builder"

        Matches = @(
            "Microsoft Report Builder*"
        )
    },

    @{
        Name = "Microsoft SQL Server 2012 Native Client"

        Matches = @(
            "Microsoft SQL Server 2012 Native Client*"
        )
    },

    @{
        Name = "Microsoft SQL Server Management Studio"

        Matches = @(
            "Microsoft SQL Server Management Studio*"
        )
    },

    @{
        Name = "Microsoft Visio - en-us"

        Matches = @(
            "Microsoft Visio - en-us*",
            "Microsoft Visio*"
        )
    },

    @{
        Name = "Microsoft Visual Studio 2010 Tools for Office Runtime"

        Matches = @(
            "Microsoft Visual Studio 2010 Tools for Office Runtime*"
        )
    },

    @{
        Name = "SunSystems 6.4.0"

        Matches = @(
            "SunSystems 6.4.0*",
            "SunSystems*"
        )
    },

    @{
        Name = "Git"

        Matches = @(
            "Git",
            "Git version*"
        )
    }
)

# ============================================================
# FUNCTIONS
# ============================================================

function Convert-InstallDate {

    param(
        [string]$InstallDate
    )

    if ([string]::IsNullOrWhiteSpace($InstallDate)) {

        return "Unknown"
    }

    # Standard uninstall registry format = yyyyMMdd
    if ($InstallDate -match '^\d{8}$') {

        try {

            return [datetime]::ParseExact(
                $InstallDate,
                "yyyyMMdd",
                [System.Globalization.CultureInfo]::InvariantCulture
            ).ToString("dd/MM/yyyy")
        }
        catch {

            return $InstallDate
        }
    }

    try {

        return ([datetime]$InstallDate).ToString("dd/MM/yyyy")
    }
    catch {

        return $InstallDate
    }
}

# ============================================================
# HTML ENCODING
# ============================================================

function HtmlEncode {

    param(
        [object]$Text
    )

    if ($null -eq $Text) {

        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode(
        [string]$Text
    )
}

# ============================================================
# GET INSTALLED APPLICATIONS
# ============================================================

function Get-InstalledApplications {

    $RegistryLocations = @(

        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",

        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $Results = @()

    foreach ($RegistryPath in $RegistryLocations) {

        $Apps = Get-ItemProperty `
            -Path $RegistryPath `
            -ErrorAction SilentlyContinue |
            Where-Object {

                -not [string]::IsNullOrWhiteSpace(
                    $_.DisplayName
                )
            }

        foreach ($App in $Apps) {

            $ConvertedDate = Convert-InstallDate `
                $App.InstallDate

            $Results += [PSCustomObject]@{

                DisplayName = $App.DisplayName

                DisplayVersion = $App.DisplayVersion

                Publisher = $App.Publisher

                InstallDate = $ConvertedDate

                RawInstallDate = $App.InstallDate

                RegistryPath = $App.PSPath
            }
        }
    }

    return $Results
}

# ============================================================
# FIND APPLICATION
#
# IMPORTANT:
# Duplicate entries for the same version are grouped together.
#
# Preference is:
#
#   1. Entry with a proper install date
#   2. First matching name/pattern
#
# This fixes the duplicate Power BI issue and also allows us
# to favour the dated Microsoft 365 registry entry.
# ============================================================

function Find-Application {

    param(

        [array]$InstalledApps,

        [array]$Patterns
    )

    $Matches = @()

    # Patterns are processed in preference order
    foreach ($Pattern in $Patterns) {

        $Found = @(
            $InstalledApps |
                Where-Object {

                    $_.DisplayName -like $Pattern
                }
        )

        if ($Found.Count -gt 0) {

            foreach ($Item in $Found) {

                # Record which pattern preference located it
                $PatternIndex = [array]::IndexOf(
                    $Patterns,
                    $Pattern
                )

                $Matches += [PSCustomObject]@{

                    DisplayName    = $Item.DisplayName
                    DisplayVersion = $Item.DisplayVersion
                    Publisher      = $Item.Publisher
                    InstallDate    = $Item.InstallDate
                    RawInstallDate = $Item.RawInstallDate
                    RegistryPath   = $Item.RegistryPath
                    PatternIndex   = $PatternIndex
                }
            }
        }
    }

    if ($Matches.Count -eq 0) {

        return @()
    }

    # Remove exact registry duplicates first
    $Matches = @(
        $Matches |
            Sort-Object `
                DisplayName,
                DisplayVersion,
                InstallDate `
                -Unique
    )

    # --------------------------------------------------------
    # Group entries by VERSION.
    #
    # Power BI example:
    #
    # Microsoft Power BI Desktop (x64)
    # Microsoft PowerBI Desktop (x64)
    #
    # Same version = treat as same application.
    #
    # Prefer the entry that has a real install date.
    # --------------------------------------------------------

    $Deduplicated = @()

    $VersionGroups = $Matches |
        Group-Object {

            if (
                [string]::IsNullOrWhiteSpace(
                    $_.DisplayVersion
                )
            ) {

                "__NOVERSION__$($_.DisplayName)"
            }
            else {

                $_.DisplayVersion
            }
        }

    foreach ($Group in $VersionGroups) {

        $Preferred = $Group.Group |
            Sort-Object `

                @{
                    Expression = {

                        if (
                            $_.InstallDate -and
                            $_.InstallDate -ne "Unknown"
                        ) {

                            0
                        }
                        else {

                            1
                        }
                    }
                },

                @{
                    Expression = {

                        $_.PatternIndex
                    }
                },

                @{
                    Expression = {

                        $_.DisplayName
                    }
                } |

            Select-Object -First 1

        $Deduplicated += $Preferred
    }

    return $Deduplicated
}

# ============================================================
# TEST CURRENT MONTH
# ============================================================

function Test-CurrentMonth {

    param(
        [string]$Date
    )

    if (
        [string]::IsNullOrWhiteSpace($Date) -or
        $Date -eq "Unknown"
    ) {

        return $false
    }

    try {

        $ParsedDate = [datetime]::ParseExact(
            $Date,
            "dd/MM/yyyy",
            [System.Globalization.CultureInfo]::InvariantCulture
        )

        if (
            $ParsedDate.Month -eq $Now.Month -and
            $ParsedDate.Year -eq $Now.Year
        ) {

            return $true
        }
    }
    catch {

        return $false
    }

    return $false
}

# ============================================================
# CONSOLE HEADER
# ============================================================

Write-Host ""
Write-Host "===============================================" `
    -ForegroundColor Cyan

Write-Host " AVD APPLICATION REPORT" `
    -ForegroundColor Cyan

Write-Host "===============================================" `
    -ForegroundColor Cyan

Write-Host ""

Write-Host "Computer : $ComputerName"
Write-Host "Period   : $MonthName"

Write-Host ""

# ============================================================
# COLLECT APPLICATION DATA
# ============================================================

Write-Host "Collecting installed applications..." `
    -ForegroundColor Yellow

$InstalledApps = @(
    Get-InstalledApplications
)

Write-Host `
    "Applications discovered: $($InstalledApps.Count)" `
    -ForegroundColor Green

# ============================================================
# OS INFORMATION
# ============================================================

Write-Host "Collecting Windows information..." `
    -ForegroundColor Yellow

$OS = Get-CimInstance `
    Win32_OperatingSystem

$OSInfo = [PSCustomObject]@{

    ComputerName = $ComputerName

    OS = $OS.Caption

    Version = $OS.Version

    Build = $OS.BuildNumber

    Architecture = $OS.OSArchitecture

    LastBootTime = $OS.LastBootUpTime.ToString(
        "dd/MM/yyyy HH:mm"
    )
}

# ============================================================
# WINDOWS UPDATE HISTORY
# ============================================================

$RecentUpdates = @()

try {

    $HotFixes = Get-HotFix |
        Sort-Object InstalledOn -Descending

    foreach ($HotFix in $HotFixes) {

        if ($HotFix.InstalledOn) {

            $RecentUpdates += [PSCustomObject]@{

                HotFixID = $HotFix.HotFixID

                Description = $HotFix.Description

                InstalledOn = $HotFix.InstalledOn
            }
        }
    }
}
catch {

    Write-Warning `
        "Unable to retrieve Windows Update history."
}

# ============================================================
# CORE APPLICATION UPDATE STATUS
# ============================================================

Write-Host `
    "Building core application update status..." `
    -ForegroundColor Yellow

$CoreResults = @()

foreach ($Definition in $CoreApplications) {

    $Matches = @(
        Find-Application `
            -InstalledApps $InstalledApps `
            -Patterns $Definition.Matches
    )

    if ($Matches.Count -gt 0) {

        foreach ($Match in $Matches) {

            $UpdatedThisMonth = Test-CurrentMonth `
                $Match.InstallDate

            if ($UpdatedThisMonth) {

                $Status = "Updated this month"
            }
            else {

                # User-requested wording.
                #
                # This indicates the required application is
                # already present even where Windows does not
                # expose a usable update date.
                $Status = "Latest version already installed"
            }

            $CoreResults += [PSCustomObject]@{

                Application = $Definition.Name

                InstalledName = $Match.DisplayName

                Version = if (
                    [string]::IsNullOrWhiteSpace(
                        $Match.DisplayVersion
                    )
                ) {

                    "Unknown"
                }
                else {

                    $Match.DisplayVersion
                }

                InstallDate = $Match.InstallDate

                Status = $Status
            }
        }
    }
    else {

        $CoreResults += [PSCustomObject]@{

            Application = $Definition.Name

            InstalledName = "Not Installed"

            Version = "-"

            InstallDate = "-"

            Status = "Not Installed"
        }
    }
}

# ============================================================
# APPLICATION INVENTORY
# ============================================================

Write-Host `
    "Building application inventory..." `
    -ForegroundColor Yellow

$InventoryResults = @()

foreach ($Definition in $InventoryApps) {

    $Matches = @(
        Find-Application `
            -InstalledApps $InstalledApps `
            -Patterns $Definition.Matches
    )

    if ($Matches.Count -gt 0) {

        foreach ($Match in $Matches) {

            $InventoryResults += [PSCustomObject]@{

                ApplicationName = $Match.DisplayName

                Version = if (
                    [string]::IsNullOrWhiteSpace(
                        $Match.DisplayVersion
                    )
                ) {

                    "Unknown"
                }
                else {

                    $Match.DisplayVersion
                }

                InstallDate = $Match.InstallDate
            }
        }
    }
    else {

        $InventoryResults += [PSCustomObject]@{

            ApplicationName = $Definition.Name

            Version = "Not Installed"

            InstallDate = "-"
        }
    }
}

# ============================================================
# OTHER APPLICATIONS INSTALLED / UPDATED THIS MONTH
# ============================================================

Write-Host `
    "Looking for other applications installed or updated this month..." `
    -ForegroundColor Yellow

$OtherApplications = @()

foreach ($App in $InstalledApps) {

    if (-not (Test-CurrentMonth $App.InstallDate)) {

        continue
    }

    $IsCoreApplication = $false

    foreach ($Definition in $CoreApplications) {

        foreach ($Pattern in $Definition.Matches) {

            if ($App.DisplayName -like $Pattern) {

                $IsCoreApplication = $true

                break
            }
        }

        if ($IsCoreApplication) {

            break
        }
    }

    # Do not repeat applications already shown in Core table
    if (-not $IsCoreApplication) {

        $OtherApplications += [PSCustomObject]@{

            Application = $App.DisplayName

            Version = if (
                [string]::IsNullOrWhiteSpace(
                    $App.DisplayVersion
                )
            ) {

                "Unknown"
            }
            else {

                $App.DisplayVersion
            }

            Publisher = $App.Publisher

            InstallDate = $App.InstallDate
        }
    }
}

# ------------------------------------------------------------
# Remove duplicate entries from Other Applications
# ------------------------------------------------------------

$OtherApplications = @(
    $OtherApplications |
        Sort-Object `
            Application,
            Version,
            InstallDate `
            -Unique
)

# ============================================================
# CSV OUTPUT
# ============================================================

$Prefix = "$ComputerName-$DateStamp"

$CoreStatusCSV = Join-Path `
    $LocalMonthFolder `
    "$Prefix-CoreApplicationUpdateStatus.csv"

$InventoryCSV = Join-Path `
    $LocalMonthFolder `
    "$Prefix-ApplicationInventory.csv"

$OSInfoCSV = Join-Path `
    $LocalMonthFolder `
    "$Prefix-OSInfo.csv"

$OSUpdatesCSV = Join-Path `
    $LocalMonthFolder `
    "$Prefix-OSUpdates.csv"

$OtherAppsCSV = Join-Path `
    $LocalMonthFolder `
    "$Prefix-OtherApplications.csv"

$CoreResults |
    Export-Csv `
        -Path $CoreStatusCSV `
        -NoTypeInformation `
        -Encoding UTF8

$InventoryResults |
    Export-Csv `
        -Path $InventoryCSV `
        -NoTypeInformation `
        -Encoding UTF8

$OSInfo |
    Export-Csv `
        -Path $OSInfoCSV `
        -NoTypeInformation `
        -Encoding UTF8

$RecentUpdates |
    Export-Csv `
        -Path $OSUpdatesCSV `
        -NoTypeInformation `
        -Encoding UTF8

$OtherApplications |
    Export-Csv `
        -Path $OtherAppsCSV `
        -NoTypeInformation `
        -Encoding UTF8

# ============================================================
# OPERATING SYSTEM HTML
# ============================================================

$OSHtml = @"
<table>

<tr>
    <th>Computer Name</th>
    <th>Operating System</th>
    <th>Version</th>
    <th>Build</th>
    <th>Architecture</th>
    <th>Last Boot</th>
</tr>

<tr>
    <td>$(HtmlEncode $OSInfo.ComputerName)</td>
    <td>$(HtmlEncode $OSInfo.OS)</td>
    <td>$(HtmlEncode $OSInfo.Version)</td>
    <td>$(HtmlEncode $OSInfo.Build)</td>
    <td>$(HtmlEncode $OSInfo.Architecture)</td>
    <td>$(HtmlEncode $OSInfo.LastBootTime)</td>
</tr>

</table>
"@

# ============================================================
# WINDOWS UPDATE HTML
# ============================================================

$UpdateHtml = @"
<table>

<tr>
    <th>KB</th>
    <th>Description</th>
    <th>Installed On</th>
</tr>
"@

$RecentUpdates |
    Select-Object -First 10 |
    ForEach-Object {

        $UpdateClass = ""

        if (
            $_.InstalledOn.Month -eq $Now.Month -and
            $_.InstalledOn.Year -eq $Now.Year
        ) {

            $UpdateClass = "good"
        }

        $UpdateHtml += @"
<tr>

    <td>
        $(HtmlEncode $_.HotFixID)
    </td>

    <td>
        $(HtmlEncode $_.Description)
    </td>

    <td class="$UpdateClass">
        $($_.InstalledOn.ToString("dd/MM/yyyy"))
    </td>

</tr>
"@
    }

$UpdateHtml += "</table>"

# ============================================================
# CORE APPLICATION UPDATE STATUS HTML
# ============================================================

$CoreHtml = @"
<table>

<tr>
    <th>Application</th>
    <th>Installed Name</th>
    <th>Version</th>
    <th>Install / Update Date</th>
    <th>Status</th>
</tr>
"@

foreach ($App in $CoreResults) {

    switch ($App.Status) {

        "Updated this month" {

            $StatusClass = "good"
        }

        "Latest version already installed" {

            $StatusClass = "good"
        }

        "Not Installed" {

            $StatusClass = "bad"
        }

        default {

            $StatusClass = ""
        }
    }

    $CoreHtml += @"
<tr>

    <td>
        $(HtmlEncode $App.Application)
    </td>

    <td>
        $(HtmlEncode $App.InstalledName)
    </td>

    <td>
        $(HtmlEncode $App.Version)
    </td>

    <td>
        $(HtmlEncode $App.InstallDate)
    </td>

    <td class="$StatusClass">
        $(HtmlEncode $App.Status)
    </td>

</tr>
"@
}

$CoreHtml += "</table>"

# ============================================================
# OTHER APPLICATIONS HTML
# ============================================================

if ($OtherApplications.Count -gt 0) {

    $OtherAppsHtml = @"
<table>

<tr>
    <th>Application</th>
    <th>Version</th>
    <th>Publisher</th>
    <th>Install / Update Date</th>
</tr>
"@

    foreach ($App in $OtherApplications) {

        $OtherAppsHtml += @"
<tr>

    <td>
        $(HtmlEncode $App.Application)
    </td>

    <td>
        $(HtmlEncode $App.Version)
    </td>

    <td>
        $(HtmlEncode $App.Publisher)
    </td>

    <td class="good">
        $(HtmlEncode $App.InstallDate)
    </td>

</tr>
"@
    }

    $OtherAppsHtml += "</table>"
}
else {

    $OtherAppsHtml = @"
<div class="info">

No additional applications were detected as installed or
updated during $MonthName.

</div>
"@
}

# ============================================================
# APPLICATION INVENTORY HTML
# ============================================================

$InventoryHtml = @"
<table>

<tr>
    <th>Application Name</th>
    <th>Version</th>
    <th>Install Date</th>
</tr>
"@

foreach ($App in $InventoryResults) {

    if ($App.Version -eq "Not Installed") {

        $RowClass = "notinstalled"
    }
    else {

        $RowClass = ""
    }

    $InventoryHtml += @"
<tr class="$RowClass">

    <td>
        $(HtmlEncode $App.ApplicationName)
    </td>

    <td>
        $(HtmlEncode $App.Version)
    </td>

    <td>
        $(HtmlEncode $App.InstallDate)
    </td>

</tr>
"@
}

$InventoryHtml += "</table>"

# ============================================================
# COMPLETE HTML REPORT
# ============================================================

$HtmlReport = @"
<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<title>
AVD Application Report - $ComputerName
</title>

<style>

body {

    font-family:
        "Segoe UI",
        Arial,
        sans-serif;

    background-color: #f4f6f8;

    margin: 0;

    padding: 25px;

    color: #2f3542;
}

.container {

    max-width: 1500px;

    margin: auto;

    background: white;

    padding: 30px;

    box-shadow:
        0 2px 12px
        rgba(0,0,0,0.12);

    border-radius: 8px;
}

h1 {

    margin-top: 0;

    color: #1f4e78;

    border-bottom:
        4px solid #1f4e78;

    padding-bottom: 12px;
}

h2 {

    margin-top: 35px;

    color: #1f4e78;

    border-bottom:
        2px solid #d9e2f3;

    padding-bottom: 8px;
}

.subtitle {

    color: #666;

    margin-bottom: 25px;
}

.summary {

    background: #eef4fa;

    border-left:
        5px solid #1f4e78;

    padding: 15px 20px;

    margin-bottom: 25px;
}

.summary strong {

    color: #1f4e78;
}

table {

    border-collapse: collapse;

    width: 100%;

    margin-top: 12px;

    margin-bottom: 25px;

    font-size: 14px;
}

th {

    background-color: #1f4e78;

    color: white;

    text-align: left;

    padding: 10px;

    border:
        1px solid #d0d7de;
}

td {

    padding: 9px 10px;

    border:
        1px solid #d0d7de;
}

tr:nth-child(even) {

    background-color: #f7f9fb;
}

tr:hover {

    background-color: #eef4fa;
}

.good {

    background-color:
        #dff0d8 !important;

    color: #2e6b2e;

    font-weight: 600;
}

.warning {

    background-color:
        #fff3cd !important;

    color: #856404;

    font-weight: 600;
}

.bad {

    background-color:
        #f8d7da !important;

    color: #842029;

    font-weight: 600;
}

.notinstalled td {

    background-color:
        #f3f3f3 !important;

    color: #777;
}

.info {

    background-color: #eef4fa;

    padding: 15px;

    border-left:
        5px solid #1f4e78;
}

.footer {

    margin-top: 40px;

    font-size: 12px;

    color: #777;

    text-align: center;
}

</style>

</head>

<body>

<div class="container">

<h1>
AVD Master Image Application Report
</h1>

<div class="subtitle">

Generated: $RunTime

</div>

<div class="summary">

<strong>Server:</strong>
$ComputerName

<br>

<strong>Reporting Period:</strong>
$MonthName

<br>

<strong>Installed Applications Detected:</strong>
$($InstalledApps.Count)

<br>

<strong>Core Applications:</strong>
$($CoreApplications.Count)

<br>

<strong>Inventory Applications:</strong>
$($InventoryApps.Count)

</div>


<h2>
1. Operating System Information
</h2>

$OSHtml


<h2>
2. Recent Windows Updates
</h2>

$UpdateHtml


<h2>
3. Core Application Update Status
</h2>

$CoreHtml


<h2>
4. Other Applications Installed or Updated in $MonthName
</h2>

$OtherAppsHtml


<h2>
5. Application Inventory
</h2>

<p>

The table below shows the currently installed
version and installation/update date for the
specifically tracked business applications.

</p>

$InventoryHtml


<div class="footer">

AVD Application Report |
$ComputerName |
Generated $RunTime

</div>

</div>

</body>

</html>
"@

# ============================================================
# WRITE HTML REPORT
# ============================================================

$HtmlFileName = `
    "$ComputerName-$DateStamp-AVD-Application-Report.html"

$LocalHtml = Join-Path `
    $LocalMonthFolder `
    $HtmlFileName

$HtmlReport |
    Out-File `
        -FilePath $LocalHtml `
        -Encoding UTF8 `
        -Force

# ============================================================
# COPY OUTPUT TO NETWORK SHARE
# ============================================================

if ($ShareAvailable) {

    Write-Host `
        "Copying reports to network share..." `
        -ForegroundColor Yellow

    $FilesToCopy = @(

        $LocalHtml,

        $CoreStatusCSV,

        $InventoryCSV,

        $OSInfoCSV,

        $OSUpdatesCSV,

        $OtherAppsCSV
    )

    foreach ($File in $FilesToCopy) {

        try {

            Copy-Item `
                -Path $File `
                -Destination $ShareMonthFolder `
                -Force `
                -ErrorAction Stop
        }
        catch {

            Write-Warning `
                "Unable to copy: $File"
        }
    }
}

# ============================================================
# CONSOLE SUMMARY
# ============================================================

Write-Host ""

Write-Host `
    "===============================================" `
    -ForegroundColor Green

Write-Host `
    " REPORT COMPLETE" `
    -ForegroundColor Green

Write-Host `
    "===============================================" `
    -ForegroundColor Green

Write-Host ""

Write-Host `
    "Server:" `
    -NoNewline

Write-Host `
    " $ComputerName" `
    -ForegroundColor Cyan

Write-Host `
    "Applications discovered:" `
    -NoNewline

Write-Host `
    " $($InstalledApps.Count)" `
    -ForegroundColor Cyan

Write-Host `
    "Core applications tracked:" `
    -NoNewline

Write-Host `
    " $($CoreApplications.Count)" `
    -ForegroundColor Cyan

Write-Host `
    "Inventory applications tracked:" `
    -NoNewline

Write-Host `
    " $($InventoryApps.Count)" `
    -ForegroundColor Cyan

Write-Host ""

Write-Host "Local report:"

Write-Host `
    $LocalHtml `
    -ForegroundColor Cyan

if ($ShareAvailable) {

    Write-Host ""

    Write-Host `
        "Network report folder:"

    Write-Host `
        $ShareMonthFolder `
        -ForegroundColor Cyan
}

Write-Host ""

# ============================================================
# OPEN REPORT
# ============================================================

Start-Process $LocalHtml
