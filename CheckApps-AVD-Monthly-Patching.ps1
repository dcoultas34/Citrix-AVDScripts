<#
.SYNOPSIS
    AVD Master Image Application / Patching Report

.DESCRIPTION
    Creates an HTML report containing:

      1. Server / OS Information
      2. Application Patching Status
      3. Application Inventory
      4. Other Applications Installed / Updated This Month

    Designed for:
        CUK-AVDGI-P-001
        CUK-AVDGI-P-002

    Output:
        C:\Ultima\AVDPatchReports\<Month Year>
        \\transfer\transfer\Citrixapps\AV\<Month Year>

.NOTES
    Application install dates are obtained from the Windows
    uninstall registry where available.

    Some applications do not populate InstallDate.
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

# ------------------------------------------------------------
# Master Images
# ------------------------------------------------------------

$MasterImages = @{
    "CUK-AVDGI-P-001" = "AVD Master Image 001"
    "CUK-AVDGI-P-002" = "AVD Master Image 002"
}

# ------------------------------------------------------------
# Output locations
# ------------------------------------------------------------

$LocalRoot = "C:\Ultima\AVDPatchReports"
$ShareRoot = "\\transfer\transfer\Citrixapps\AV"

$LocalMonthFolder = Join-Path $LocalRoot $MonthName
$ShareMonthFolder = Join-Path $ShareRoot $MonthName

# Create local folder
if (-not (Test-Path $LocalMonthFolder)) {
    New-Item `
        -Path $LocalMonthFolder `
        -ItemType Directory `
        -Force | Out-Null
}

# Try to create share folder
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

    Write-Warning "Unable to access reporting share: $ShareMonthFolder"
    Write-Warning $_.Exception.Message
}

# ============================================================
# APPLICATIONS IN MAIN PATCHING TABLE
# ============================================================

$AppsToCheck = @(

    @{
        Name    = "Microsoft 365 Apps"
        Matches = @(
            "Microsoft 365 Apps*"
        )
    },

    @{
        Name    = "Adobe Acrobat"
        Matches = @(
            "Adobe Acrobat*"
        )
    },

    @{
        Name    = "Google Chrome"
        Matches = @(
            "Google Chrome*"
        )
    },

    @{
        Name    = "Microsoft Edge"
        Matches = @(
            "Microsoft Edge"
        )
    },

    @{
        Name    = "Microsoft Visual Studio Code"
        Matches = @(
            "Microsoft Visual Studio Code*"
        )
    },

    @{
        Name    = "Microsoft Power BI Desktop"
        Matches = @(
            "Microsoft PowerBI Desktop*",
            "Microsoft Power BI Desktop*"
        )
    },

    @{
        Name    = "Microsoft Teams"
        Matches = @(
            "Microsoft Teams*",
            "Microsoft Teams classic*"
        )
    },

    @{
        Name    = "Microsoft OneDrive"
        Matches = @(
            "Microsoft OneDrive*"
        )
    },

    @{
        Name    = "Mozilla Firefox"
        Matches = @(
            "Mozilla Firefox*"
        )
    }
)

# ============================================================
# NEW APPLICATION INVENTORY
# ============================================================

$InventoryApps = @(

    @{
        Name    = "8x8 Work"
        Matches = @(
            "8x8 Work*"
        )
    },

    @{
        Name    = "Adobe Acrobat (64-bit)"
        Matches = @(
            "Adobe Acrobat*"
        )
    },

    @{
        Name    = "Foxit PDF Editor"
        Matches = @(
            "Foxit PDF Editor*"
        )
    },

    @{
        Name    = "Infor Query & Analysis 11"
        Matches = @(
            "Infor Query & Analysis 11"
        )
    },

    @{
        Name    = "Infor Query & Analysis 11 DataLink Server"
        Matches = @(
            "Infor Query & Analysis 11 DataLink Server*"
        )
    },

    @{
        Name    = "Microsoft 365 Apps for business - en-gb"
        Matches = @(
            "Microsoft 365 Apps for business*"
        )
    },

    @{
        Name    = "Microsoft Edge"
        Matches = @(
            "Microsoft Edge"
        )
    },

    @{
        Name    = "Microsoft OneDrive"
        Matches = @(
            "Microsoft OneDrive*"
        )
    },

    @{
        Name    = "Microsoft PowerBI Desktop (x64)"
        Matches = @(
            "Microsoft PowerBI Desktop*",
            "Microsoft Power BI Desktop*"
        )
    },

    @{
        Name    = "Microsoft Project - en-us"
        Matches = @(
            "Microsoft Project - en-us*",
            "Microsoft Project*"
        )
    },

    @{
        Name    = "Microsoft Report Builder"
        Matches = @(
            "Microsoft Report Builder*"
        )
    },

    @{
        Name    = "Microsoft SQL Server 2012 Native Client"
        Matches = @(
            "Microsoft SQL Server 2012 Native Client*"
        )
    },

    @{
        Name    = "Microsoft SQL Server Management Studio"
        Matches = @(
            "Microsoft SQL Server Management Studio*"
        )
    },

    @{
        Name    = "Microsoft Visio - en-us"
        Matches = @(
            "Microsoft Visio - en-us*",
            "Microsoft Visio*"
        )
    },

    @{
        Name    = "Microsoft Visual Studio 2010 Tools for Office Runtime"
        Matches = @(
            "Microsoft Visual Studio 2010 Tools for Office Runtime*"
        )
    },

    @{
        Name    = "SunSystems 6.4.0"
        Matches = @(
            "SunSystems 6.4.0*",
            "SunSystems*"
        )
    },

    @{
        Name    = "Git"
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

    # Most MSI dates are yyyyMMdd
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

# ------------------------------------------------------------
# Get all installed applications
# ------------------------------------------------------------

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
                -not [string]::IsNullOrWhiteSpace($_.DisplayName)
            }

        foreach ($App in $Apps) {

            $InstallDate = Convert-InstallDate $App.InstallDate

            $Results += [PSCustomObject]@{

                DisplayName    = $App.DisplayName
                DisplayVersion = $App.DisplayVersion
                Publisher      = $App.Publisher
                InstallDate    = $InstallDate
                RawInstallDate = $App.InstallDate
                RegistryPath   = $App.PSPath
            }
        }
    }

    # Remove duplicate entries
    $Results |
        Sort-Object DisplayName, DisplayVersion -Unique
}

# ------------------------------------------------------------
# Match application
# ------------------------------------------------------------

function Find-Application {

    param(

        [array]$InstalledApps,

        [array]$Patterns
    )

    $Matches = @()

    foreach ($Pattern in $Patterns) {

        $Found = $InstalledApps |
            Where-Object {
                $_.DisplayName -like $Pattern
            }

        if ($Found) {
            $Matches += $Found
        }
    }

    $Matches |
        Sort-Object DisplayName, DisplayVersion -Unique
}

# ------------------------------------------------------------
# Determine whether date belongs to current month
# ------------------------------------------------------------

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
            $ParsedDate.Month -eq (Get-Date).Month -and
            $ParsedDate.Year -eq (Get-Date).Year
        ) {
            return $true
        }

    }
    catch {
        return $false
    }

    return $false
}

# ------------------------------------------------------------
# HTML encode
# ------------------------------------------------------------

function HtmlEncode {

    param(
        [string]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode($Text)
}

# ============================================================
# COLLECT APPLICATION DATA
# ============================================================

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " AVD APPLICATION REPORT" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Computer : $ComputerName"
Write-Host "Period   : $MonthName"
Write-Host ""

Write-Host "Collecting installed applications..." -ForegroundColor Yellow

$InstalledApps = @(Get-InstalledApplications)

Write-Host "Applications discovered: $($InstalledApps.Count)" -ForegroundColor Green

# ============================================================
# OS INFORMATION
# ============================================================

Write-Host "Collecting Windows information..." -ForegroundColor Yellow

$OS = Get-CimInstance Win32_OperatingSystem

$OSInfo = [PSCustomObject]@{

    ComputerName = $ComputerName
    OS           = $OS.Caption
    Version      = $OS.Version
    Build        = $OS.BuildNumber
    Architecture = $OS.OSArchitecture
    LastBootTime = $OS.LastBootUpTime.ToString("dd/MM/yyyy HH:mm")
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

                HotFixID    = $HotFix.HotFixID
                Description = $HotFix.Description
                InstalledOn = $HotFix.InstalledOn
            }
        }
    }

}
catch {

    Write-Warning "Unable to retrieve Windows Update history."
}

# ============================================================
# MAIN PATCHING APPLICATION TABLE
# ============================================================

Write-Host "Building patching application table..." -ForegroundColor Yellow

$AppStatusResults = @()

foreach ($AppDefinition in $AppsToCheck) {

    $Matches = @(
        Find-Application `
            -InstalledApps $InstalledApps `
            -Patterns $AppDefinition.Matches
    )

    if ($Matches.Count -gt 0) {

        foreach ($Match in $Matches) {

            $UpdatedThisMonth = Test-CurrentMonth $Match.InstallDate

            if ($UpdatedThisMonth) {

                $Status = "Updated this month"
            }
            elseif ($Match.InstallDate -eq "Unknown") {

                $Status = "Installed - update date unavailable"
            }
            else {

                $Status = "Installed"
            }

            $AppStatusResults += [PSCustomObject]@{

                Application     = $AppDefinition.Name
                InstalledName   = $Match.DisplayName

                Version         = if ($Match.DisplayVersion) {
                    $Match.DisplayVersion
                }
                else {
                    "Unknown"
                }

                InstallDate     = $Match.InstallDate
                Status          = $Status
            }
        }
    }
}

# ============================================================
# NEW APPLICATION INVENTORY TABLE
# ============================================================

Write-Host "Building application inventory..." -ForegroundColor Yellow

$InventoryResults = @()

foreach ($InventoryDefinition in $InventoryApps) {

    $Matches = @(
        Find-Application `
            -InstalledApps $InstalledApps `
            -Patterns $InventoryDefinition.Matches
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

        # Keep requested application visible even when missing
        $InventoryResults += [PSCustomObject]@{

            ApplicationName = $InventoryDefinition.Name
            Version         = "Not Installed"
            InstallDate     = "-"
        }
    }
}

# Remove accidental duplicates

$InventoryResults = @(
    $InventoryResults |
        Sort-Object ApplicationName, Version -Unique
)

# ============================================================
# OTHER APPLICATIONS INSTALLED / UPDATED THIS MONTH
# ============================================================

Write-Host "Looking for other applications updated this month..." `
    -ForegroundColor Yellow

$OtherApplications = @()

foreach ($App in $InstalledApps) {

    if (Test-CurrentMonth $App.InstallDate) {

        $AlreadyTracked = $false

        # Check main applications
        foreach ($Definition in $AppsToCheck) {

            foreach ($Pattern in $Definition.Matches) {

                if ($App.DisplayName -like $Pattern) {
                    $AlreadyTracked = $true
                    break
                }
            }

            if ($AlreadyTracked) {
                break
            }
        }

        if (-not $AlreadyTracked) {

            $OtherApplications += [PSCustomObject]@{

                Application = $App.DisplayName

                Version = if ($App.DisplayVersion) {
                    $App.DisplayVersion
                }
                else {
                    "Unknown"
                }

                Publisher   = $App.Publisher
                InstallDate = $App.InstallDate
            }
        }
    }
}

$OtherApplications = @(
    $OtherApplications |
        Sort-Object Application -Unique
)

# ============================================================
# CSV OUTPUTS
# ============================================================

$Prefix = "$ComputerName-$DateStamp"

$AppStatusCSV = Join-Path `
    $LocalMonthFolder `
    "$Prefix-AppStatus.csv"

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

$AppStatusResults |
    Export-Csv `
        -Path $AppStatusCSV `
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
# HTML TABLE BUILDERS
# ============================================================

# ------------------------------------------------------------
# OS table
# ------------------------------------------------------------

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

# ------------------------------------------------------------
# Main application status table
# ------------------------------------------------------------

$AppHtml = @"
<table>
<tr>
    <th>Application</th>
    <th>Installed Name</th>
    <th>Version</th>
    <th>Install / Update Date</th>
    <th>Status</th>
</tr>
"@

foreach ($App in $AppStatusResults) {

    switch ($App.Status) {

        "Updated this month" {
            $Class = "good"
        }

        "Installed - update date unavailable" {
            $Class = "warning"
        }

        default {
            $Class = ""
        }
    }

    $AppHtml += @"
<tr>
    <td>$(HtmlEncode $App.Application)</td>
    <td>$(HtmlEncode $App.InstalledName)</td>
    <td>$(HtmlEncode $App.Version)</td>
    <td>$(HtmlEncode $App.InstallDate)</td>
    <td class="$Class">$(HtmlEncode $App.Status)</td>
</tr>
"@
}

$AppHtml += "</table>"

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
    <td>$(HtmlEncode $App.ApplicationName)</td>
    <td>$(HtmlEncode $App.Version)</td>
    <td>$(HtmlEncode $App.InstallDate)</td>
</tr>
"@
}

$InventoryHtml += "</table>"

# ============================================================
# WINDOWS UPDATE TABLE
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

        $UpdateDate = $_.InstalledOn

        if (
            $UpdateDate.Month -eq $Now.Month -and
            $UpdateDate.Year -eq $Now.Year
        ) {
            $UpdateClass = "good"
        }
        else {
            $UpdateClass = ""
        }

        $UpdateHtml += @"
<tr>
    <td>$(HtmlEncode $_.HotFixID)</td>
    <td>$(HtmlEncode $_.Description)</td>
    <td class="$UpdateClass">$($_.InstalledOn.ToString("dd/MM/yyyy"))</td>
</tr>
"@
    }

$UpdateHtml += "</table>"

# ============================================================
# OTHER APPLICATIONS HTML
# ============================================================

$OtherAppsHtml = ""

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
    <td>$(HtmlEncode $App.Application)</td>
    <td>$(HtmlEncode $App.Version)</td>
    <td>$(HtmlEncode $App.Publisher)</td>
    <td class="good">$(HtmlEncode $App.InstallDate)</td>
</tr>
"@
    }

    $OtherAppsHtml += "</table>"
}
else {

    $OtherAppsHtml = @"
<div class="info">
No additional applications were detected as installed or updated during $MonthName.
</div>
"@
}

# ============================================================
# COMPLETE HTML REPORT
# ============================================================

$HtmlReport = @"
<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<title>AVD Application Report - $ComputerName</title>

<style>

body {
    font-family: "Segoe UI", Arial, sans-serif;
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
    box-shadow: 0 2px 12px rgba(0,0,0,0.12);
    border-radius: 8px;
}

h1 {
    margin-top: 0;
    color: #1f4e78;
    border-bottom: 4px solid #1f4e78;
    padding-bottom: 12px;
}

h2 {
    margin-top: 35px;
    color: #1f4e78;
    border-bottom: 2px solid #d9e2f3;
    padding-bottom: 8px;
}

.subtitle {
    color: #666;
    margin-bottom: 25px;
}

.summary {
    background: #eef4fa;
    border-left: 5px solid #1f4e78;
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
    border: 1px solid #d0d7de;
}

td {
    padding: 9px 10px;
    border: 1px solid #d0d7de;
}

tr:nth-child(even) {
    background-color: #f7f9fb;
}

tr:hover {
    background-color: #eef4fa;
}

.good {
    background-color: #dff0d8 !important;
    color: #2e6b2e;
    font-weight: 600;
}

.warning {
    background-color: #fff3cd !important;
    color: #856404;
    font-weight: 600;
}

.bad {
    background-color: #f8d7da !important;
    color: #842029;
    font-weight: 600;
}

.notinstalled td {
    background-color: #f3f3f3 !important;
    color: #777;
}

.info {
    background-color: #eef4fa;
    padding: 15px;
    border-left: 5px solid #1f4e78;
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

<h1>AVD Master Image Application Report</h1>

<div class="subtitle">
Generated: $RunTime
</div>

<div class="summary">

<strong>Server:</strong> $ComputerName
<br>

<strong>Reporting Period:</strong> $MonthName
<br>

<strong>Installed Applications Detected:</strong>
$($InstalledApps.Count)

<br>

<strong>Tracked Inventory Applications:</strong>
$($InventoryApps.Count)

</div>


<h2>1. Operating System Information</h2>

$OSHtml


<h2>2. Recent Windows Updates</h2>

$UpdateHtml


<h2>3. Application Patching Status</h2>

$AppHtml


<h2>4. Application Inventory</h2>

<p>
The table below provides the installed version and installation /
update date for the specifically tracked business applications.
</p>

$InventoryHtml


<h2>5. Other Applications Installed or Updated in $MonthName</h2>

$OtherAppsHtml


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

$HtmlFileName = "$ComputerName-$DateStamp-AVD-Application-Report.html"

$LocalHtml = Join-Path `
    $LocalMonthFolder `
    $HtmlFileName

$HtmlReport |
    Out-File `
        -FilePath $LocalHtml `
        -Encoding UTF8 `
        -Force

# ============================================================
# COPY TO NETWORK SHARE
# ============================================================

if ($ShareAvailable) {

    Write-Host "Copying reports to network share..." `
        -ForegroundColor Yellow

    $FilesToCopy = @(

        $LocalHtml,
        $AppStatusCSV,
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

            Write-Warning "Unable to copy: $File"
        }
    }
}

# ============================================================
# CONSOLE SUMMARY
# ============================================================

Write-Host ""
Write-Host "===============================================" -ForegroundColor Green
Write-Host " REPORT COMPLETE" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green
Write-Host ""

Write-Host "Server:" -NoNewline
Write-Host " $ComputerName" -ForegroundColor Cyan

Write-Host "Applications discovered:" -NoNewline
Write-Host " $($InstalledApps.Count)" -ForegroundColor Cyan

Write-Host "Inventory applications:" -NoNewline
Write-Host " $($InventoryResults.Count)" -ForegroundColor Cyan

Write-Host ""
Write-Host "Local report:"
Write-Host $LocalHtml -ForegroundColor Cyan

if ($ShareAvailable) {

    Write-Host ""
    Write-Host "Network report folder:"
    Write-Host $ShareMonthFolder -ForegroundColor Cyan
}

Write-Host ""

# Open report automatically
Start-Process $LocalHtml
