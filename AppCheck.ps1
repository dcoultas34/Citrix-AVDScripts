<#
.SYNOPSIS
    Checks nominated Citrix master images for legacy applications.

.DESCRIPTION
    Run this script locally on each master image.

    The script:
      1. Checks traditional installed applications.
      2. Checks installed and provisioned Windows Appx packages.
      3. Checks applicable Windows optional features.
      4. Uploads a CSV report for the current machine.
      5. Rebuilds a combined HTML report covering all expected machines.

.NOTES
    Run from an elevated PowerShell window.

    The account running the script requires Modify permissions to:

    \\Transfer\Transfer\CitrixApps\Citrix reporting\Appremoval
#>

[CmdletBinding()]
param (
    [string]$ReportFolder = "\\Transfer\Transfer\CitrixApps\Citrix reporting\Appremoval"
)

$ErrorActionPreference = "Stop"

# -------------------------------------------------------------------------
# Expected master images
# -------------------------------------------------------------------------

$ExpectedDevices = @(
    "UKST1MICTXMI3U"
    "UKST1MICTXMI5"
    "UKST1MICTXMI4C"
    "UKST1MICTXMI2D"
    "UKST1MICTXMI1D"
    "UKST1MICTXMI2U"
    "UKST1MICTXMI1U"
)

# -------------------------------------------------------------------------
# Software to check
#
# RegistryPatterns:
#   Matches traditional Add/Remove Programs entries.
#
# ExcludePatterns:
#   Excludes supporting components that might otherwise match.
#
# AppxPatterns:
#   Matches installed or provisioned Windows Apps.
#
# FeaturePatterns:
#   Matches enabled Windows optional features.
# -------------------------------------------------------------------------

$SoftwareChecks = @(

    [PSCustomObject]@{
        Name             = "Microsoft Visual C++ 2005 Redistributable"
        RegistryPatterns = @("*Microsoft Visual C++ 2005*Redistributable*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft Visual C++ 2010 Redistributable"
        RegistryPatterns = @("*Microsoft Visual C++ 2010*Redistributable*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft SQL Server 2012 Native Client"
        RegistryPatterns = @("*SQL Server 2012 Native Client*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft Visual C++ 2013 Redistributable"
        RegistryPatterns = @("*Microsoft Visual C++ 2013*Redistributable*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft SQL Server System CLR Types 2014"
        RegistryPatterns = @(
            "*SQL Server*System CLR Types*2014*"
            "*System CLR Types for SQL Server 2014*"
        )
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft Visual Studio 2015 Tools for Applications"
        RegistryPatterns = @("*Visual Studio 2015 Tools for Applications*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft Visual Studio 2015 Shell (Isolated)"
        RegistryPatterns = @("*Visual Studio 2015 Shell*Isolated*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft SQL Server 2016 Integration Services"
        RegistryPatterns = @("*SQL Server 2016 Integration Services*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft SQL Server System CLR Types 2016"
        RegistryPatterns = @(
            "*SQL Server*System CLR Types*2016*"
            "*System CLR Types for SQL Server 2016*"
        )
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft OData Source 13 for Microsoft SQL Server 2016"
        RegistryPatterns = @(
            "*OData Source*SQL Server 2016*"
            "*Microsoft OData Source 13*"
        )
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    # This checks only the main SQL Server Management Studio installation.
    # Supporting SQL Server components are explicitly excluded.
    [PSCustomObject]@{
        Name             = "Microsoft SQL Server 2016 Management Studio"
        RegistryPatterns = @(
            "Microsoft SQL Server 2016 Management Studio"
            "Microsoft SQL Server Management Studio*"
            "SQL Server Management Studio*"
        )
        ExcludePatterns  = @(
            "*Setup*"
            "*Language Pack*"
            "*Shared Management Objects*"
            "*Management Objects*"
            "*Management Components*"
            "*Tools*"
            "*Components*"
            "*Add-in*"
            "*Prerequisites*"
        )
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft Visual Studio 2017 Tools for Applications"
        RegistryPatterns = @("*Visual Studio 2017 Tools for Applications*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft .NET Framework 4.5 SDK"
        RegistryPatterns = @(
            "*.NET Framework 4.5 SDK*"
            "*NET Framework 4.5 SDK*"
        )
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft .NET Framework 4.5 Targeting Pack"
        RegistryPatterns = @(
            "*.NET Framework 4.5*Targeting Pack*"
            "*NET Framework 4.5*Targeting Pack*"
        )
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft Internet Explorer 11"
        RegistryPatterns = @("*Internet Explorer*11*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @(
            "Internet-Explorer-Optional*"
            "Internet-Explorer-Optional-amd64"
        )
    }

    [PSCustomObject]@{
        Name             = "ABBYY FineReader PDF 14"
        RegistryPatterns = @("*ABBYY FineReader*14*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft ODBC Driver 13 for SQL Server"
        RegistryPatterns = @("*ODBC Driver 13 for SQL Server*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Java SE 11 (OpenJDK)"
        RegistryPatterns = @(
            "*OpenJDK*11*"
            "*OpenJDK 11*"
            "*Temurin*11*"
            "*Adoptium*11*"
        )
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft Visual Studio 2019 Tools for Applications"
        RegistryPatterns = @("*Visual Studio 2019 Tools for Applications*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft Azure Information Protection Unified Labeling Client"
        RegistryPatterns = @(
            "*Azure Information Protection*Unified Labeling*"
            "*Azure Information Protection Client*"
        )
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Paint 3D"
        RegistryPatterns = @("*Paint 3D*")
        ExcludePatterns  = @()
        AppxPatterns     = @(
            "*Microsoft.MSPaint*"
            "*MSPaint*"
        )
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft .NET Runtime 6"
        RegistryPatterns = @(
            "*Microsoft .NET Runtime*6*"
            "*Microsoft.NETCore.App*6*"
        )
        ExcludePatterns  = @(
            "*Desktop Runtime*"
            "*Windows Desktop Runtime*"
            "*Hosting Bundle*"
            "*ASP.NET*"
        )
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft .NET Desktop Runtime 6"
        RegistryPatterns = @(
            "*Microsoft Windows Desktop Runtime*6*"
            "*Microsoft.WindowsDesktop.App*6*"
            "*Windows Desktop Runtime*6*"
        )
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "PSFTP 0.81"
        RegistryPatterns = @("*PSFTP*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "PuTTY 0.81"
        RegistryPatterns = @("*PuTTY*")
        ExcludePatterns  = @("*PSFTP*")
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft SQL Server System CLR Types 2019"
        RegistryPatterns = @(
            "*SQL Server*System CLR Types*2019*"
            "*System CLR Types for SQL Server 2019*"
        )
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft SQL Server 2019 LocalDB"
        RegistryPatterns = @(
            "*SQL Server 2019 LocalDB*"
            "*Microsoft SQL Server 2019 LocalDB*"
        )
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Skype 8"
        RegistryPatterns = @("*Skype*")
        ExcludePatterns  = @(
            "*Skype for Business*"
            "*Meeting Add-in*"
            "*Teams Meeting Add-in*"
        )
        AppxPatterns     = @(
            "*Microsoft.SkypeApp*"
            "*SkypeApp*"
        )
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft Remote Desktop"
        RegistryPatterns = @(
            "Microsoft Remote Desktop"
            "Microsoft Remote Desktop*"
        )
        ExcludePatterns  = @(
            "*WebRTC*"
            "*Redirector*"
            "*Multimedia*"
            "*Services Infrastructure*"
        )
        AppxPatterns     = @(
            "*MicrosoftCorporationII.RemoteDesktop*"
            "*Microsoft.RemoteDesktop*"
            "*RemoteDesktop*"
        )
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Microsoft Power BI Desktop"
        RegistryPatterns = @(
            "Microsoft Power BI Desktop"
            "Microsoft Power BI Desktop*"
        )
        ExcludePatterns  = @(
            "*Report Server*"
            "*Gateway*"
            "*Language Pack*"
        )
        AppxPatterns     = @(
            "*Microsoft.MicrosoftPowerBIDesktop*"
            "*PowerBIDesktop*"
            "*PowerBI*"
        )
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Citrix Workspace App"
        RegistryPatterns = @(
            "*Citrix Workspace*"
            "*Citrix Receiver*"
        )
        ExcludePatterns  = @(
            "*Browser Content Redirection*"
            "*HDX RealTime*"
            "*Teams Redirection*"
        )
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Chromium"
        RegistryPatterns = @("*Chromium*")
        ExcludePatterns  = @()
        AppxPatterns     = @("*Chromium*")
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "IntelliJ IDEA Ultimate Edition"
        RegistryPatterns = @("*IntelliJ IDEA*Ultimate*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "3D Viewer"
        RegistryPatterns = @("*3D Viewer*")
        ExcludePatterns  = @()
        AppxPatterns     = @(
            "*Microsoft.Microsoft3DViewer*"
            "*Microsoft3DViewer*"
        )
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Azure Data Studio"
        RegistryPatterns = @("*Azure Data Studio*")
        ExcludePatterns  = @()
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }

    [PSCustomObject]@{
        Name             = "Google Chrome"
        RegistryPatterns = @("*Google Chrome*")
        ExcludePatterns  = @(
            "*Update Helper*"
            "*Elevation Service*"
        )
        AppxPatterns     = @()
        FeaturePatterns  = @()
    }
)

# -------------------------------------------------------------------------
# Helper functions
# -------------------------------------------------------------------------

function ConvertTo-HtmlSafe {
    param (
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Test-WildcardMatch {
    param (
        [AllowNull()]
        [string]$Value,

        [AllowNull()]
        [string[]]$Patterns
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    if ($null -eq $Patterns -or $Patterns.Count -eq 0) {
        return $false
    }

    foreach ($Pattern in $Patterns) {
        if ($Value -like $Pattern) {
            return $true
        }
    }

    return $false
}

function Test-SoftwareNameMatch {
    param (
        [string]$Value,
        [string[]]$IncludePatterns,
        [string[]]$ExcludePatterns
    )

    $Included = Test-WildcardMatch `
        -Value $Value `
        -Patterns $IncludePatterns

    if (-not $Included) {
        return $false
    }

    $Excluded = Test-WildcardMatch `
        -Value $Value `
        -Patterns $ExcludePatterns

    return (-not $Excluded)
}

function Get-InstalledProgramInventory {

    $RegistryLocations = @(
        [PSCustomObject]@{
            Path         = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
            Architecture = "64-bit"
            Scope        = "Machine"
        }

        [PSCustomObject]@{
            Path         = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
            Architecture = "32-bit"
            Scope        = "Machine"
        }

        [PSCustomObject]@{
            Path         = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
            Architecture = "Current user"
            Scope        = "Current user"
        }
    )

    $Inventory = foreach ($Location in $RegistryLocations) {

        Get-ItemProperty `
            -Path $Location.Path `
            -ErrorAction SilentlyContinue |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.DisplayName)
            } |
            ForEach-Object {
                [PSCustomObject]@{
                    DetectedName = [string]$_.DisplayName
                    Version      = [string]$_.DisplayVersion
                    Publisher    = [string]$_.Publisher
                    Architecture = $Location.Architecture
                    Scope        = $Location.Scope
                    InstallType  = "Installed program"
                    Source       = [string]$_.PSPath
                }
            }
    }

    return @(
        $Inventory |
            Sort-Object `
                DetectedName,
                Version,
                Architecture,
                Scope `
                -Unique
    )
}

function Get-AppxInventory {

    $Inventory = @()

    try {
        $Inventory += Get-AppxPackage `
            -AllUsers `
            -ErrorAction Stop |
            ForEach-Object {
                [PSCustomObject]@{
                    DetectedName = [string]$_.Name
                    PackageName  = [string]$_.PackageFullName
                    Version      = [string]$_.Version
                    Publisher    = [string]$_.Publisher
                    Architecture = [string]$_.Architecture
                    Scope        = "Installed Appx"
                    InstallType  = "Windows App"
                    Source       = [string]$_.InstallLocation
                }
            }
    }
    catch {
        Write-Warning "Unable to read all-user Appx packages: $($_.Exception.Message)"
    }

    try {
        $Inventory += Get-AppxProvisionedPackage `
            -Online `
            -ErrorAction Stop |
            ForEach-Object {
                [PSCustomObject]@{
                    DetectedName = [string]$_.DisplayName
                    PackageName  = [string]$_.PackageName
                    Version      = [string]$_.Version
                    Publisher    = ""
                    Architecture = [string]$_.Architecture
                    Scope        = "Provisioned Appx"
                    InstallType  = "Provisioned Windows App"
                    Source       = [string]$_.PackagePath
                }
            }
    }
    catch {
        Write-Warning "Unable to read provisioned Appx packages: $($_.Exception.Message)"
    }

    return @(
        $Inventory |
            Sort-Object `
                DetectedName,
                PackageName,
                Version,
                Scope `
                -Unique
    )
}

function Get-OptionalFeatureInventory {

    try {
        return @(
            Get-WindowsOptionalFeature `
                -Online `
                -ErrorAction Stop |
                ForEach-Object {
                    [PSCustomObject]@{
                        DetectedName = [string]$_.FeatureName
                        Version      = ""
                        Publisher    = "Microsoft Corporation"
                        Architecture = "Operating system"
                        Scope        = "Machine"
                        InstallType  = "Windows optional feature"
                        State        = [string]$_.State
                        Source       = [string]$_.FeatureName
                    }
                }
        )
    }
    catch {
        Write-Warning "Unable to read Windows optional features: $($_.Exception.Message)"
        return @()
    }
}

function New-SoftwareResult {
    param (
        [string]$ExpectedItem,
        [bool]$Installed,
        [string]$DetectedName = "",
        [string]$Version = "",
        [string]$Publisher = "",
        [string]$Architecture = "",
        [string]$Scope = "",
        [string]$InstallType = "",
        [string]$Source = ""
    )

    [PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME.ToUpperInvariant()
        ExpectedItem = $ExpectedItem
        Installed    = if ($Installed) { "Yes" } else { "No" }
        DetectedName = $DetectedName
        Version      = $Version
        Publisher    = $Publisher
        Architecture = $Architecture
        Scope        = $Scope
        InstallType  = $InstallType
        Source       = $Source
        CheckedAt    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
}

function Write-CombinedHtmlReport {
    param (
        [string]$Folder,
        [string[]]$Devices,
        [object[]]$Checks
    )

    Write-Host "Reading submitted machine reports..." -ForegroundColor Cyan

    $AllResults = @()
    $DeviceReportInformation = @{}

    foreach ($Device in $Devices) {

        $DeviceName = $Device.ToUpperInvariant()
        $CsvPath = Join-Path $Folder "$DeviceName-AppRemoval.csv"

        if (Test-Path -LiteralPath $CsvPath) {

            try {
                $ImportedResults = @(
                    Import-Csv `
                        -LiteralPath $CsvPath `
                        -ErrorAction Stop
                )

                $AllResults += $ImportedResults

                $DeviceReportInformation[$DeviceName] = [PSCustomObject]@{
                    ReportFound = $true
                    LastWrite   = (Get-Item -LiteralPath $CsvPath).LastWriteTime
                    CheckedAt   = (
                        $ImportedResults |
                            Select-Object `
                                -First 1 `
                                -ExpandProperty CheckedAt
                    )
                }
            }
            catch {
                Write-Warning "Could not read report for $DeviceName. $($_.Exception.Message)"

                $DeviceReportInformation[$DeviceName] = [PSCustomObject]@{
                    ReportFound = $false
                    LastWrite   = $null
                    CheckedAt   = $null
                }
            }
        }
        else {
            $DeviceReportInformation[$DeviceName] = [PSCustomObject]@{
                ReportFound = $false
                LastWrite   = $null
                CheckedAt   = $null
            }
        }
    }

    $MachinesReported = @(
        $DeviceReportInformation.GetEnumerator() |
            Where-Object {
                $_.Value.ReportFound
            }
    ).Count

    $MachinesMissing = $Devices.Count - $MachinesReported

    $ApplicationsInstalled = @(
        $AllResults |
            Where-Object {
                $_.Installed -eq "Yes"
            } |
            Select-Object `
                ComputerName,
                ExpectedItem `
                -Unique
    ).Count

    # ---------------------------------------------------------------------
    # Machine reporting-status table
    # ---------------------------------------------------------------------

    $MachineStatusRows = foreach ($Device in $Devices) {

        $DeviceName = $Device.ToUpperInvariant()
        $Information = $DeviceReportInformation[$DeviceName]

        if ($Information.ReportFound) {
            $Status = "Report received"
            $RowClass = "status-good"
            $CheckedAt = $Information.CheckedAt
        }
        else {
            $Status = "No report received"
            $RowClass = "status-warning"
            $CheckedAt = ""
        }

        @"
<tr class="$RowClass">
    <td>$(ConvertTo-HtmlSafe $DeviceName)</td>
    <td>$(ConvertTo-HtmlSafe $Status)</td>
    <td>$(ConvertTo-HtmlSafe $CheckedAt)</td>
</tr>
"@
    }

    # ---------------------------------------------------------------------
    # Application matrix
    # ---------------------------------------------------------------------

    $DeviceHeadings = foreach ($Device in $Devices) {
        "<th>$(ConvertTo-HtmlSafe $Device)</th>"
    }

    $MatrixRows = foreach ($Check in $Checks) {

        $InstalledOnCount = 0

        $Cells = foreach ($Device in $Devices) {

            $DeviceName = $Device.ToUpperInvariant()
            $Information = $DeviceReportInformation[$DeviceName]

            if (-not $Information.ReportFound) {
                '<td class="no-report">No report</td>'
                continue
            }

            $Matches = @(
                $AllResults |
                    Where-Object {
                        $_.ComputerName -eq $DeviceName -and
                        $_.ExpectedItem -eq $Check.Name -and
                        $_.Installed -eq "Yes"
                    }
            )

            if ($Matches.Count -eq 0) {
                '<td class="not-installed">Not installed</td>'
                continue
            }

            $InstalledOnCount++

            $DetectedEntries = foreach ($Match in $Matches) {

                $VersionText = if (
                    [string]::IsNullOrWhiteSpace([string]$Match.Version)
                ) {
                    "Version not recorded"
                }
                else {
                    "Version: $($Match.Version)"
                }

                $ArchitectureText = if (
                    [string]::IsNullOrWhiteSpace([string]$Match.Architecture)
                ) {
                    ""
                }
                else {
                    " [$($Match.Architecture)]"
                }

                $ScopeText = if (
                    [string]::IsNullOrWhiteSpace([string]$Match.Scope)
                ) {
                    ""
                }
                else {
                    "<br><span class=`"small-text`">$(ConvertTo-HtmlSafe $Match.Scope)</span>"
                }

                @"
<strong>$(ConvertTo-HtmlSafe $Match.DetectedName)</strong><br>
$(ConvertTo-HtmlSafe $VersionText)$(ConvertTo-HtmlSafe $ArchitectureText)$ScopeText
"@
            }

            $UniqueEntries = @(
                $DetectedEntries |
                    Sort-Object -Unique
            )

            "<td class=`"installed`">$($UniqueEntries -join '<hr>')</td>"
        }

        $RowClass = if ($InstalledOnCount -gt 0) {
            "application-found"
        }
        else {
            "application-clear"
        }

        @"
<tr class="$RowClass">
    <td class="application-name">
        <strong>$(ConvertTo-HtmlSafe $Check.Name)</strong>
    </td>
    <td class="count-cell">$InstalledOnCount</td>
    $($Cells -join "`n")
</tr>
"@
    }

    # ---------------------------------------------------------------------
    # Detailed detections
    # ---------------------------------------------------------------------

    $InstalledResults = @(
        $AllResults |
            Where-Object {
                $_.Installed -eq "Yes"
            } |
            Sort-Object `
                ExpectedItem,
                ComputerName,
                DetectedName,
                Version
    )

    if ($InstalledResults.Count -gt 0) {

        $DetailedRows = foreach ($Result in $InstalledResults) {

            @"
<tr>
    <td>$(ConvertTo-HtmlSafe $Result.ExpectedItem)</td>
    <td>$(ConvertTo-HtmlSafe $Result.ComputerName)</td>
    <td>$(ConvertTo-HtmlSafe $Result.DetectedName)</td>
    <td>$(ConvertTo-HtmlSafe $Result.Version)</td>
    <td>$(ConvertTo-HtmlSafe $Result.Architecture)</td>
    <td>$(ConvertTo-HtmlSafe $Result.Scope)</td>
    <td>$(ConvertTo-HtmlSafe $Result.InstallType)</td>
    <td>$(ConvertTo-HtmlSafe $Result.CheckedAt)</td>
</tr>
"@
        }
    }
    else {
        $DetailedRows = @"
<tr>
    <td colspan="8" class="not-installed">
        No listed applications were detected in the submitted reports.
    </td>
</tr>
"@
    }

    $GeneratedAt = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

    $Html = @"
<!DOCTYPE html>
<html lang="en">

<head>

<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>Citrix Master Image Application Removal Report</title>

<style>

    body {
        margin: 24px;
        font-family: "Segoe UI", Arial, sans-serif;
        color: #202020;
        background-color: #f5f7fa;
    }

    h1 {
        margin-bottom: 4px;
        color: #183153;
    }

    h2 {
        margin-top: 32px;
        color: #183153;
    }

    .subtitle {
        margin-top: 0;
        color: #5b6573;
    }

    .summary-container {
        display: flex;
        flex-wrap: wrap;
        gap: 14px;
        margin-top: 22px;
        margin-bottom: 24px;
    }

    .summary-card {
        min-width: 180px;
        padding: 16px;
        border: 1px solid #d5dce5;
        border-radius: 6px;
        background-color: #ffffff;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
    }

    .summary-number {
        display: block;
        margin-top: 6px;
        font-size: 27px;
        font-weight: 700;
        color: #183153;
    }

    .report-section {
        margin-top: 20px;
        padding: 18px;
        overflow-x: auto;
        border: 1px solid #d5dce5;
        border-radius: 6px;
        background-color: #ffffff;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);
    }

    table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
    }

    th {
        position: sticky;
        top: 0;
        padding: 9px;
        border: 1px solid #bcc7d3;
        background-color: #dceaf7;
        text-align: left;
        white-space: nowrap;
    }

    td {
        padding: 8px;
        border: 1px solid #d4d9df;
        vertical-align: top;
    }

    .application-name {
        min-width: 260px;
    }

    .count-cell {
        text-align: center;
        font-weight: 700;
    }

    .installed {
        min-width: 175px;
        background-color: #dff2e1;
    }

    .not-installed {
        background-color: #f0f2f4;
        color: #5d6570;
        text-align: center;
    }

    .no-report {
        background-color: #fff1cc;
        color: #775500;
        font-weight: 600;
        text-align: center;
    }

    .application-found .application-name,
    .application-found .count-cell {
        background-color: #ffe2e2;
    }

    .application-clear .application-name,
    .application-clear .count-cell {
        background-color: #e4f3e5;
    }

    .status-good {
        background-color: #e4f3e5;
    }

    .status-warning {
        background-color: #fff1cc;
    }

    .small-text {
        color: #59636e;
        font-size: 11px;
    }

    hr {
        height: 1px;
        margin: 8px 0;
        border: 0;
        background-color: #b8c2cc;
    }

    .legend {
        margin-top: 16px;
        line-height: 1.8;
    }

    .legend-box {
        display: inline-block;
        width: 18px;
        height: 14px;
        margin-right: 5px;
        border: 1px solid #b0b0b0;
        vertical-align: middle;
    }

    .legend-installed {
        background-color: #dff2e1;
    }

    .legend-missing {
        background-color: #f0f2f4;
    }

    .legend-no-report {
        background-color: #fff1cc;
    }

    .footer {
        margin-top: 22px;
        color: #67717e;
        font-size: 12px;
    }

</style>

</head>

<body>

<h1>Citrix Master Image Application Removal Report</h1>

<p class="subtitle">
    Combined application inventory for the nominated Citrix master images.
</p>

<div class="summary-container">

    <div class="summary-card">
        Expected machines
        <span class="summary-number">$($Devices.Count)</span>
    </div>

    <div class="summary-card">
        Reports received
        <span class="summary-number">$MachinesReported</span>
    </div>

    <div class="summary-card">
        Reports outstanding
        <span class="summary-number">$MachinesMissing</span>
    </div>

    <div class="summary-card">
        Application detections
        <span class="summary-number">$ApplicationsInstalled</span>
    </div>

</div>

<h2>Machine reporting status</h2>

<div class="report-section">

<table>

<thead>
<tr>
    <th>Machine</th>
    <th>Status</th>
    <th>Last checked</th>
</tr>
</thead>

<tbody>
$($MachineStatusRows -join "`n")
</tbody>

</table>

</div>

<h2>Application summary</h2>

<div class="report-section">

<table>

<thead>
<tr>
    <th>Application</th>
    <th>Machines installed on</th>
    $($DeviceHeadings -join "`n")
</tr>
</thead>

<tbody>
$($MatrixRows -join "`n")
</tbody>

</table>

<div class="legend">

    <span class="legend-box legend-installed"></span>
    Installed

    &nbsp;&nbsp;

    <span class="legend-box legend-missing"></span>
    Not installed

    &nbsp;&nbsp;

    <span class="legend-box legend-no-report"></span>
    No report received

</div>

</div>

<h2>Detailed detections</h2>

<div class="report-section">

<table>

<thead>
<tr>
    <th>Expected application</th>
    <th>Machine</th>
    <th>Detected name</th>
    <th>Version</th>
    <th>Architecture</th>
    <th>Scope</th>
    <th>Installation type</th>
    <th>Checked</th>
</tr>
</thead>

<tbody>
$($DetailedRows -join "`n")
</tbody>

</table>

</div>

<p class="footer">
    Report generated: $GeneratedAt<br>
    Report location: $(ConvertTo-HtmlSafe $Folder)
</p>

</body>

</html>
"@

    $FinalHtmlPath = Join-Path `
        $Folder `
        "Combined-AppRemoval-Report.html"

    $TemporaryHtmlPath = Join-Path `
        $Folder `
        "Combined-AppRemoval-Report-$([guid]::NewGuid().Guid).tmp"

    $Html |
        Out-File `
            -LiteralPath $TemporaryHtmlPath `
            -Encoding UTF8 `
            -Force

    Move-Item `
        -LiteralPath $TemporaryHtmlPath `
        -Destination $FinalHtmlPath `
        -Force

    return $FinalHtmlPath
}

# -------------------------------------------------------------------------
# Validate machine and destination folder
# -------------------------------------------------------------------------

$CurrentComputer = $env:COMPUTERNAME.ToUpperInvariant()

Write-Host ""
Write-Host "Computer: $CurrentComputer" -ForegroundColor Cyan

if ($CurrentComputer -notin $ExpectedDevices) {

    Write-Warning @"
$CurrentComputer is not in the expected master-image list.

The inventory will still be collected, but the machine will not appear as
a column in the combined report unless it is added to `$ExpectedDevices.
"@
}

if (-not (Test-Path -LiteralPath $ReportFolder)) {

    Write-Host "Creating report folder: $ReportFolder" -ForegroundColor Cyan

    New-Item `
        -Path $ReportFolder `
        -ItemType Directory `
        -Force |
        Out-Null
}

# Test write access to the shared folder.
$WriteTestPath = Join-Path `
    $ReportFolder `
    "$CurrentComputer-$([guid]::NewGuid().Guid).write-test"

try {

    "Write test" |
        Out-File `
            -LiteralPath $WriteTestPath `
            -Encoding UTF8 `
            -Force

    Remove-Item `
        -LiteralPath $WriteTestPath `
        -Force
}
catch {

    throw @"
Unable to write to the reporting folder:

$ReportFolder

Check the share permissions, NTFS permissions and the account running
PowerShell.

Error: $($_.Exception.Message)
"@
}

# -------------------------------------------------------------------------
# Collect software inventories
# -------------------------------------------------------------------------

Write-Host "Reading traditional installed programs..." -ForegroundColor Cyan
$InstalledPrograms = Get-InstalledProgramInventory

Write-Host "Reading installed and provisioned Windows Apps..." -ForegroundColor Cyan
$AppxPackages = Get-AppxInventory

Write-Host "Reading Windows optional features..." -ForegroundColor Cyan
$OptionalFeatures = Get-OptionalFeatureInventory

# -------------------------------------------------------------------------
# Check each application
# -------------------------------------------------------------------------

$Results = foreach ($Check in $SoftwareChecks) {

    $Matches = @()

    # Traditional installed applications
    if ($Check.RegistryPatterns.Count -gt 0) {

        $RegistryMatches = @(
            $InstalledPrograms |
                Where-Object {

                    Test-SoftwareNameMatch `
                        -Value $_.DetectedName `
                        -IncludePatterns $Check.RegistryPatterns `
                        -ExcludePatterns $Check.ExcludePatterns
                }
        )

        foreach ($Match in $RegistryMatches) {

            $Matches += New-SoftwareResult `
                -ExpectedItem $Check.Name `
                -Installed $true `
                -DetectedName $Match.DetectedName `
                -Version $Match.Version `
                -Publisher $Match.Publisher `
                -Architecture $Match.Architecture `
                -Scope $Match.Scope `
                -InstallType $Match.InstallType `
                -Source $Match.Source
        }
    }

    # Installed and provisioned Windows Apps
    if ($Check.AppxPatterns.Count -gt 0) {

        $AppxMatches = @(
            $AppxPackages |
                Where-Object {

                    $NameMatch = Test-WildcardMatch `
                        -Value $_.DetectedName `
                        -Patterns $Check.AppxPatterns

                    $PackageMatch = Test-WildcardMatch `
                        -Value $_.PackageName `
                        -Patterns $Check.AppxPatterns

                    $NameMatch -or $PackageMatch
                }
        )

        foreach ($Match in $AppxMatches) {

            $Matches += New-SoftwareResult `
                -ExpectedItem $Check.Name `
                -Installed $true `
                -DetectedName $Match.DetectedName `
                -Version $Match.Version `
                -Publisher $Match.Publisher `
                -Architecture $Match.Architecture `
                -Scope $Match.Scope `
                -InstallType $Match.InstallType `
                -Source $Match.Source
        }
    }

    # Windows optional features
    if ($Check.FeaturePatterns.Count -gt 0) {

        $FeatureMatches = @(
            $OptionalFeatures |
                Where-Object {

                    $_.State -eq "Enabled" -and
                    (
                        Test-WildcardMatch `
                            -Value $_.DetectedName `
                            -Patterns $Check.FeaturePatterns
                    )
                }
        )

        foreach ($Match in $FeatureMatches) {

            $Matches += New-SoftwareResult `
                -ExpectedItem $Check.Name `
                -Installed $true `
                -DetectedName $Match.DetectedName `
                -Version $Match.Version `
                -Publisher $Match.Publisher `
                -Architecture $Match.Architecture `
                -Scope $Match.Scope `
                -InstallType $Match.InstallType `
                -Source $Match.Source
        }
    }

    $Matches = @(
        $Matches |
            Sort-Object `
                ExpectedItem,
                DetectedName,
                Version,
                Architecture,
                Scope `
                -Unique
    )

    if ($Matches.Count -gt 0) {
        $Matches
    }
    else {
        New-SoftwareResult `
            -ExpectedItem $Check.Name `
            -Installed $false
    }
}

# -------------------------------------------------------------------------
# Upload current machine's CSV report
# -------------------------------------------------------------------------

$MachineCsvPath = Join-Path `
    $ReportFolder `
    "$CurrentComputer-AppRemoval.csv"

$TemporaryCsvPath = Join-Path `
    $ReportFolder `
    "$CurrentComputer-AppRemoval-$([guid]::NewGuid().Guid).tmp"

$Results |
    Sort-Object `
        ExpectedItem,
        Installed,
        DetectedName,
        Version |
    Export-Csv `
        -LiteralPath $TemporaryCsvPath `
        -NoTypeInformation `
        -Encoding UTF8 `
        -Force

Move-Item `
    -LiteralPath $TemporaryCsvPath `
    -Destination $MachineCsvPath `
    -Force

Write-Host ""
Write-Host "Machine report uploaded:" -ForegroundColor Green
Write-Host $MachineCsvPath

# -------------------------------------------------------------------------
# Rebuild combined HTML report
# -------------------------------------------------------------------------

$CombinedReportPath = Write-CombinedHtmlReport `
    -Folder $ReportFolder `
    -Devices $ExpectedDevices `
    -Checks $SoftwareChecks

Write-Host ""
Write-Host "Combined report updated:" -ForegroundColor Green
Write-Host $CombinedReportPath

# -------------------------------------------------------------------------
# Display current machine's detections
# -------------------------------------------------------------------------

$LocalInstalledCount = @(
    $Results |
        Where-Object {
            $_.Installed -eq "Yes"
        } |
        Select-Object `
            ExpectedItem `
            -Unique
).Count

Write-Host ""
Write-Host "Applications detected on $CurrentComputer`: $LocalInstalledCount" `
    -ForegroundColor Yellow

$Results |
    Select-Object `
        ExpectedItem,
        Installed,
        DetectedName,
        Version,
        Architecture,
        InstallType |
    Sort-Object `
        ExpectedItem,
        DetectedName |
    Format-Table -AutoSize

Write-Host ""
Write-Host "Opening the combined HTML report..." -ForegroundColor Cyan

Start-Process $CombinedReportPath
