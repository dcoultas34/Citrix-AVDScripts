<#
.SYNOPSIS
    AVD Master Image Monthly Application Report

.DESCRIPTION
    Generates:
      1. Operating System Information
      2. Recent Windows Updates
      3. Core Application Update Status
      4. Other Applications Installed or Updated This Month
      5. Application Inventory

    Duplicate uninstall-registry entries are handled automatically.
    Each configured Core or Inventory application returns ONE row only.

.OUTPUTS
    C:\Ultima\AVDPatchReports\<Month Year>
    \\transfer\transfer\Citrixapps\AV\<Month Year>
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

$LocalRoot = "C:\Ultima\AVDPatchReports"
$ShareRoot = "\\transfer\transfer\Citrixapps\AV"

$LocalMonthFolder = Join-Path $LocalRoot $MonthName
$ShareMonthFolder = Join-Path $ShareRoot $MonthName

# ============================================================
# CREATE OUTPUT FOLDERS
# ============================================================

if (-not (Test-Path $LocalMonthFolder)) {
    New-Item -Path $LocalMonthFolder -ItemType Directory -Force | Out-Null
}

$ShareAvailable = $false

try {
    if (-not (Test-Path $ShareMonthFolder)) {
        New-Item -Path $ShareMonthFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    $ShareAvailable = $true
}
catch {
    Write-Warning "Unable to access network report location: $ShareMonthFolder"
    Write-Warning $_.Exception.Message
}

# ============================================================
# CORE APPLICATIONS
#
# ONE ROW ONLY PER APPLICATION
# ============================================================

$CoreApplications = @(
    @{
        Name = "Microsoft 365 Apps"
        Matches = @(
            "Microsoft 365 Apps for business - en-gb",
            "Microsoft 365 Apps for business*"
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
            "Git"
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

function HtmlEncode {

    param(
        [object]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function Get-InstalledApplications {

    $RegistryLocations = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $Results = @()

    foreach ($RegistryPath in $RegistryLocations) {

        $Apps = Get-ItemProperty -Path $RegistryPath -ErrorAction SilentlyContinue |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.DisplayName)
            }

        foreach ($App in $Apps) {

            $Results += [PSCustomObject]@{
                DisplayName    = [string]$App.DisplayName
                DisplayVersion = [string]$App.DisplayVersion
                Publisher      = [string]$App.Publisher
                InstallDate    = Convert-InstallDate $App.InstallDate
                RawInstallDate = [string]$App.InstallDate
                RegistryPath   = [string]$App.PSPath
            }
        }
    }

    return @($Results)
}

# ============================================================
# FIND ONE BEST MATCH ONLY
#
# Preference:
#   1. Real InstallDate
#   2. Has DisplayVersion
#   3. Earlier/more specific matching pattern
#
# This is what stops duplicate Power BI / Office entries.
# ============================================================

function Find-BestApplication {

    param(
        [array]$InstalledApps,
        [array]$Patterns
    )

    $Candidates = @()

    for ($i = 0; $i -lt $Patterns.Count; $i++) {

        $Pattern = $Patterns[$i]

        $FoundApps = $InstalledApps |
            Where-Object {
                $_.DisplayName -like $Pattern
            }

        foreach ($App in $FoundApps) {

            $HasDate = 0
            $HasVersion = 0

            if (
                -not [string]::IsNullOrWhiteSpace($App.InstallDate) -and
                $App.InstallDate -ne "Unknown"
            ) {
                $HasDate = 1
            }

            if (-not [string]::IsNullOrWhiteSpace($App.DisplayVersion)) {
                $HasVersion = 1
            }

            $Candidates += [PSCustomObject]@{
                DisplayName    = $App.DisplayName
                DisplayVersion = $App.DisplayVersion
                Publisher      = $App.Publisher
                InstallDate    = $App.InstallDate
                RawInstallDate = $App.RawInstallDate
                RegistryPath   = $App.RegistryPath
                PatternIndex   = $i
                HasDate        = $HasDate
                HasVersion     = $HasVersion
            }
        }
    }

    if ($Candidates.Count -eq 0) {
        return $null
    }

    # Remove identical candidate records
    $Candidates = @(
        $Candidates |
            Sort-Object DisplayName, DisplayVersion, InstallDate -Unique
    )

    # Select ONE best result
    $Best = $Candidates |
        Sort-Object `
            @{Expression = {$_.HasDate}; Descending = $true},
            @{Expression = {$_.HasVersion}; Descending = $true},
            @{Expression = {$_.PatternIndex}; Descending = $false},
            @{Expression = {$_.DisplayName}; Descending = $false} |
        Select-Object -First 1

    return $Best
}

function Test-CurrentMonth {

    param(
        [string]$Date
    )

    if (
        [string]::IsNullOrWhiteSpace($Date) -or
        $Date -eq "Unknown" -or
        $Date -eq "-"
    ) {
        return $false
    }

    try {
        $ParsedDate = [datetime]::ParseExact(
            $Date,
            "dd/MM/yyyy",
            [System.Globalization.CultureInfo]::InvariantCulture
        )

        return (
            $ParsedDate.Month -eq $Now.Month -and
            $ParsedDate.Year -eq $Now.Year
        )
    }
    catch {
        return $false
    }
}

# ============================================================
# START
# ============================================================

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " AVD MASTER IMAGE APPLICATION REPORT" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Computer : $ComputerName"
Write-Host "Period   : $MonthName"
Write-Host ""

# ============================================================
# GET INSTALLED APPLICATIONS
# ============================================================

Write-Host "Collecting installed applications..." -ForegroundColor Yellow

$InstalledApps = @(Get-InstalledApplications)

Write-Host "Registry entries found: $($InstalledApps.Count)" -ForegroundColor Green

# ============================================================
# OPERATING SYSTEM INFORMATION
# ============================================================

Write-Host "Collecting operating system information..." -ForegroundColor Yellow

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
# WINDOWS UPDATES
# ============================================================

Write-Host "Collecting Windows Update information..." -ForegroundColor Yellow

$RecentUpdates = @()

try {
    $RecentUpdates = @(
        Get-HotFix |
            Where-Object {$_.InstalledOn} |
            Sort-Object InstalledOn -Descending
    )
}
catch {
    Write-Warning "Unable to retrieve Windows Update information."
}

# ============================================================
# CORE APPLICATION UPDATE STATUS
#
# EXACTLY ONE ROW PER APPLICATION
# ============================================================

Write-Host "Building Core Application Update Status..." -ForegroundColor Yellow

$CoreResults = @()

foreach ($Definition in $CoreApplications) {

    $Match = Find-BestApplication `
        -InstalledApps $InstalledApps `
        -Patterns $Definition.Matches

    if ($null -eq $Match) {

        $CoreResults += [PSCustomObject]@{
            Application   = $Definition.Name
            InstalledName = "Not Installed"
            Version       = "-"
            InstallDate   = "-"
            Status        = "Not Installed"
        }

        continue
    }

    if (Test-CurrentMonth $Match.InstallDate) {
        $Status = "Updated this month"
    }
    else {
        $Status = "Latest version already installed"
    }

    if ([string]::IsNullOrWhiteSpace($Match.DisplayVersion)) {
        $Version = "Unknown"
    }
    else {
        $Version = $Match.DisplayVersion
    }

    $CoreResults += [PSCustomObject]@{
        Application   = $Definition.Name
        InstalledName = $Match.DisplayName
        Version       = $Version
        InstallDate   = $Match.InstallDate
        Status        = $Status
    }
}

# ============================================================
# APPLICATION INVENTORY
#
# EXACTLY ONE ROW PER APPLICATION
# ============================================================

Write-Host "Building Application Inventory..." -ForegroundColor Yellow

$InventoryResults = @()

foreach ($Definition in $InventoryApps) {

    $Match = Find-BestApplication `
        -InstalledApps $InstalledApps `
        -Patterns $Definition.Matches

    if ($null -eq $Match) {

        $InventoryResults += [PSCustomObject]@{
            ApplicationName = $Definition.Name
            Version         = "Not Installed"
            InstallDate     = "-"
        }

        continue
    }

    if ([string]::IsNullOrWhiteSpace($Match.DisplayVersion)) {
        $Version = "Unknown"
    }
    else {
        $Version = $Match.DisplayVersion
    }

    $InventoryResults += [PSCustomObject]@{
        ApplicationName = $Match.DisplayName
        Version         = $Version
        InstallDate     = $Match.InstallDate
    }
}

# ============================================================
# OTHER APPLICATIONS UPDATED THIS MONTH
# ============================================================

Write-Host "Finding other applications updated this month..." -ForegroundColor Yellow

$OtherApplications = @()

foreach ($App in $InstalledApps) {

    if (-not (Test-CurrentMonth $App.InstallDate)) {
        continue
    }

    # --------------------------------------------------------
    # Exclude anything already in the Core Applications table
    # --------------------------------------------------------

    $IsCoreApp = $false

    foreach ($Definition in $CoreApplications) {

        foreach ($Pattern in $Definition.Matches) {

            if ($App.DisplayName -like $Pattern) {
                $IsCoreApp = $true
                break
            }
        }

        if ($IsCoreApp) {
            break
        }
    }

    if ($IsCoreApp) {
        continue
    }

    if ([string]::IsNullOrWhiteSpace($App.DisplayVersion)) {
        $Version = "Unknown"
    }
    else {
        $Version = $App.DisplayVersion
    }

    $OtherApplications += [PSCustomObject]@{
        Application = $App.DisplayName
        Version     = $Version
        Publisher   = $App.Publisher
        InstallDate = $App.InstallDate
    }
}

# Remove duplicate rows
$OtherApplications = @(
    $OtherApplications |
        Sort-Object Application, Version, InstallDate -Unique
)

# ============================================================
# OUTPUT FILE PATHS
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

# ============================================================
# EXPORT CSV FILES
# ============================================================

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
# BUILD OPERATING SYSTEM HTML
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
# BUILD WINDOWS UPDATE HTML
# ============================================================

$UpdateHtml = @"
<table>
<tr>
    <th>KB</th>
    <th>Description</th>
    <th>Installed On</th>
</tr>
"@

foreach ($Update in ($RecentUpdates | Select-Object -First 10)) {

    $UpdateClass = ""

    if (
        $Update.InstalledOn.Month -eq $Now.Month -and
        $Update.InstalledOn.Year -eq $Now.Year
    ) {
        $UpdateClass = "good"
    }

    $UpdateHtml += @"
<tr>
    <td>$(HtmlEncode $Update.HotFixID)</td>
    <td>$(HtmlEncode $Update.Description)</td>
    <td class="$UpdateClass">$($Update.InstalledOn.ToString("dd/MM/yyyy"))</td>
</tr>
"@
}

$UpdateHtml += "</table>"

# ============================================================
# BUILD CORE APPLICATION HTML
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

    if ($App.Status -eq "Not Installed") {
        $StatusClass = "bad"
    }
    else {
        $StatusClass = "good"
    }

    $CoreHtml += @"
<tr>
    <td>$(HtmlEncode $App.Application)</td>
    <td>$(HtmlEncode $App.InstalledName)</td>
    <td>$(HtmlEncode $App.Version)</td>
    <td>$(HtmlEncode $App.InstallDate)</td>
    <td class="$StatusClass">$(HtmlEncode $App.Status)</td>
</tr>
"@
}

$CoreHtml += "</table>"

# ============================================================
# BUILD OTHER APPLICATIONS HTML
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
# BUILD APPLICATION INVENTORY HTML
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

<strong>Server:</strong>
$ComputerName

<br>

<strong>Reporting Period:</strong>
$MonthName

<br>

<strong>Core Applications:</strong>
$($CoreApplications.Count)

<br>

<strong>Inventory Applications:</strong>
$($InventoryApps.Count)

</div>

<h2>1. Operating System Information</h2>

$OSHtml

<h2>2. Recent Windows Updates</h2>

$UpdateHtml

<h2>3. Core Application Update Status</h2>

$CoreHtml

<h2>4. Other Applications Installed or Updated in $MonthName</h2>

$OtherAppsHtml

<h2>5. Application Inventory</h2>

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
# SAVE HTML
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

    Write-Host ""
    Write-Host "Copying report files to network share..." -ForegroundColor Yellow

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
            Write-Warning "Unable to copy file: $File"
        }
    }
}

# ============================================================
# FINISH
# ============================================================

Write-Host ""
Write-Host "===============================================" -ForegroundColor Green
Write-Host " REPORT COMPLETE" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green
Write-Host ""

Write-Host "Server: $ComputerName"
Write-Host "Core applications: $($CoreApplications.Count)"
Write-Host "Inventory applications: $($InventoryApps.Count)"

Write-Host ""
Write-Host "Local report:"
Write-Host $LocalHtml -ForegroundColor Cyan

if ($ShareAvailable) {
    Write-Host ""
    Write-Host "Network report folder:"
    Write-Host $ShareMonthFolder -ForegroundColor Cyan
}

Write-Host ""

Start-Process $LocalHtml
