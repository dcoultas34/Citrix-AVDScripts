<#
===========================================================================
 CITRIX MASTER IMAGE AV / SECURITY AUDIT
===========================================================================

 Designed to work in PowerShell ConstrainedLanguage mode.

 Output:
   \\transfer\transfer\Citrixapps\AV\CSV
   \\transfer\transfer\Citrixapps\AV\Combined

 Creates:
   <SERVER>-SecurityStatus.csv
   Citrix-Security-Combined.csv
   Citrix-Security-Report.html

=========================================================================== 
#>


# ==========================================================================
# CONFIGURATION
# ==========================================================================

$ReportRoot = "\\transfer\transfer\Citrixapps\AV"

$CsvFolder = Join-Path $ReportRoot "CSV"

$CombinedFolder = Join-Path $ReportRoot "Combined"

$CombinedCsv = Join-Path $CombinedFolder "Citrix-Security-Combined.csv"

$HtmlReport = Join-Path $CombinedFolder "Citrix-Security-Report.html"


# ==========================================================================
# CREATE OUTPUT FOLDERS
# ==========================================================================

foreach ($Folder in @(
    $ReportRoot,
    $CsvFolder,
    $CombinedFolder
)) {

    if (!(Test-Path $Folder)) {

        try {

            New-Item `
                -Path $Folder `
                -ItemType Directory `
                -Force `
                -ErrorAction Stop |
                Out-Null

        }
        catch {

            Write-Host ""
            Write-Host "ERROR: Unable to create/access:" -ForegroundColor Red
            Write-Host $Folder -ForegroundColor Yellow

            Write-Host ""
            Write-Host $_.Exception.Message -ForegroundColor Red

            exit
        }
    }
}


# ==========================================================================
# FUNCTION - HTML SAFE TEXT
#
# Does not use .NET methods so it is safer under ConstrainedLanguage.
# ==========================================================================

function Convert-HtmlSafe {

    param(
        $Text
    )

    if ($null -eq $Text) {

        return ""
    }

    $Output = "$Text"

    $Output = $Output -replace '&', '&amp;'
    $Output = $Output -replace '<', '&lt;'
    $Output = $Output -replace '>', '&gt;'
    $Output = $Output -replace '"', '&quot;'

    return $Output
}


# ==========================================================================
# FUNCTION - DISPLAY BOOLEAN
# ==========================================================================

function Get-EnabledStatus {

    param(
        $Value
    )

    if ($Value -eq $true) {

        return "Enabled"
    }

    if ($Value -eq $false) {

        return "Disabled"
    }

    return "Unknown"
}


# ==========================================================================
# FUNCTION - DETECT SECURITY SOFTWARE
# ==========================================================================

function Get-SecuritySoftware {

    $RegistryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )


    $SecurityRegex = @(
        "Microsoft Defender",
        "Defender",
        "CrowdStrike",
        "Falcon",
        "SentinelOne",
        "Sophos",
        "McAfee",
        "Trellix",
        "Symantec",
        "Carbon Black",
        "Trend Micro",
        "Cylance",
        "FireEye",
        "Rapid7",
        "Qualys",
        "Nessus",
        "Zscaler",
        "Cisco Secure",
        "Cisco AMP",
        "Palo Alto",
        "Cortex",
        "Forcepoint"
    ) -join "|"


    $Products = Get-ItemProperty `
        -Path $RegistryPaths `
        -ErrorAction SilentlyContinue |
        Where-Object {

            $_.DisplayName -and
            $_.DisplayName -match $SecurityRegex

        } |
        Sort-Object DisplayName -Unique


    if (!$Products) {

        return "No additional security tooling detected"
    }


    $ProductList = @()


    foreach ($Product in $Products) {

        if ($Product.DisplayVersion) {

            $ProductList += "$($Product.DisplayName) [$($Product.DisplayVersion)]"

        }
        else {

            $ProductList += $Product.DisplayName
        }
    }


    return ($ProductList -join "; ")
}


# ==========================================================================
# FUNCTION - COLLECT SECURITY STATUS
# ==========================================================================

function Get-CitrixSecurityStatus {

    Write-Host ""
    Write-Host "============================================================" `
        -ForegroundColor Cyan

    Write-Host " Scanning $env:COMPUTERNAME" `
        -ForegroundColor Cyan

    Write-Host "============================================================" `
        -ForegroundColor Cyan

    Write-Host ""


    # ----------------------------------------------------------------------
    # PowerShell Language Mode
    # ----------------------------------------------------------------------

    $LanguageMode = $ExecutionContext.SessionState.LanguageMode


    # ----------------------------------------------------------------------
    # Microsoft Defender Status
    # ----------------------------------------------------------------------

    try {

        $Defender = Get-MpComputerStatus `
            -ErrorAction Stop

    }
    catch {

        $Defender = $null
    }


    # ----------------------------------------------------------------------
    # Defender for Endpoint Onboarding
    # ----------------------------------------------------------------------

    try {

        $MDE = Get-ItemProperty `
            -Path "HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status" `
            -ErrorAction Stop

    }
    catch {

        $MDE = $null
    }


    # ----------------------------------------------------------------------
    # Services
    # ----------------------------------------------------------------------

    $SenseService = Get-Service `
        -Name "Sense" `
        -ErrorAction SilentlyContinue


    $DefenderService = Get-Service `
        -Name "WinDefend" `
        -ErrorAction SilentlyContinue


    # ----------------------------------------------------------------------
    # Firewall Status
    # ----------------------------------------------------------------------

    $DomainFirewall = $null
    $PrivateFirewall = $null
    $PublicFirewall = $null


    try {

        $FirewallProfiles = Get-NetFirewallProfile `
            -ErrorAction Stop


        $DomainFirewall = (
            $FirewallProfiles |
            Where-Object {
                $_.Name -eq "Domain"
            }
        ).Enabled


        $PrivateFirewall = (
            $FirewallProfiles |
            Where-Object {
                $_.Name -eq "Private"
            }
        ).Enabled


        $PublicFirewall = (
            $FirewallProfiles |
            Where-Object {
                $_.Name -eq "Public"
            }
        ).Enabled

    }
    catch {

        $DomainFirewall = $null
        $PrivateFirewall = $null
        $PublicFirewall = $null
    }


    # ----------------------------------------------------------------------
    # MDE Onboarding Status
    # ----------------------------------------------------------------------

    if ($MDE) {

        if ($MDE.OnboardingState -eq 1) {

            $MDEStatus = "Onboarded"

        }
        elseif ($MDE.OnboardingState -eq 0) {

            $MDEStatus = "Not Onboarded"

        }
        else {

            $MDEStatus = "Unknown"
        }

    }
    else {

        $MDEStatus = "Unknown"
    }


    # ----------------------------------------------------------------------
    # Security Software
    # ----------------------------------------------------------------------

    $SecuritySoftware = Get-SecuritySoftware


    # ----------------------------------------------------------------------
    # BUILD REVIEW NOTES
    # ----------------------------------------------------------------------

    $Issues = @()


    if (!$Defender) {

        $Issues += "Unable to query Microsoft Defender"

    }
    else {

        if ($Defender.AntivirusEnabled -ne $true) {

            $Issues += "Defender Antivirus not enabled"
        }


        if ($Defender.RealTimeProtectionEnabled -ne $true) {

            $Issues += "Real-time protection disabled"
        }


        if ($Defender.BehaviorMonitorEnabled -ne $true) {

            $Issues += "Behaviour monitoring disabled"
        }


        if ($Defender.IoavProtectionEnabled -ne $true) {

            $Issues += "IOAV protection disabled"
        }


        if ($Defender.NISEnabled -ne $true) {

            $Issues += "Network inspection disabled"
        }


        if ($Defender.DefenderSignaturesOutOfDate -eq $true) {

            $Issues += "Defender signatures out of date"
        }
    }


    if ($MDEStatus -ne "Onboarded") {

        $Issues += "Defender for Endpoint onboarding not confirmed"
    }


    if (!$SenseService) {

        $Issues += "MDE Sense service not found"

    }
    elseif ($SenseService.Status -ne "Running") {

        $Issues += "MDE Sense service not running"
    }


    if (!$DefenderService) {

        $Issues += "WinDefend service not found"

    }
    elseif ($DefenderService.Status -ne "Running") {

        $Issues += "WinDefend service not running"
    }


    if ($DomainFirewall -eq $false) {

        $Issues += "Domain firewall disabled"
    }


    # ----------------------------------------------------------------------
    # OVERALL STATUS
    # ----------------------------------------------------------------------

    if ($Issues.Count -eq 0) {

        $ProtectionStatus = "Protected"

        $ReviewNotes = "No issues detected"

    }
    else {

        $ProtectionStatus = "Requires Review"

        $ReviewNotes = $Issues -join "; "
    }


    # ======================================================================
    # CREATE OUTPUT OBJECT
    #
    # IMPORTANT:
    # No [PSCustomObject] is used here.
    #
    # Select-Object creates the properties for us.
    # ======================================================================

    $Result = "" |
        Select-Object `
            ComputerName,
            ScanDate,
            PowerShellLanguageMode,
            DefenderAVEnabled,
            DefenderRunningMode,
            RealTimeProtection,
            BehaviorMonitoring,
            IOAVProtection,
            NetworkInspection,
            SignatureVersion,
            SignatureLastUpdated,
            SignaturesOutOfDate,
            TamperProtection,
            MDEOnboardingState,
            MDEOnboardingValue,
            MDESenseService,
            DefenderService,
            FirewallDomain,
            FirewallPrivate,
            FirewallPublic,
            SecuritySoftware,
            ProtectionStatus,
            ReviewNotes


    # ----------------------------------------------------------------------
    # General
    # ----------------------------------------------------------------------

    $Result.ComputerName = $env:COMPUTERNAME

    $Result.ScanDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $Result.PowerShellLanguageMode = "$LanguageMode"


    # ----------------------------------------------------------------------
    # Defender
    # ----------------------------------------------------------------------

    if ($Defender) {

        $Result.DefenderAVEnabled =
            Get-EnabledStatus $Defender.AntivirusEnabled


        $Result.DefenderRunningMode =
            "$($Defender.AMRunningMode)"


        $Result.RealTimeProtection =
            Get-EnabledStatus $Defender.RealTimeProtectionEnabled


        $Result.BehaviorMonitoring =
            Get-EnabledStatus $Defender.BehaviorMonitorEnabled


        $Result.IOAVProtection =
            Get-EnabledStatus $Defender.IoavProtectionEnabled


        $Result.NetworkInspection =
            Get-EnabledStatus $Defender.NISEnabled


        $Result.SignatureVersion =
            "$($Defender.AntivirusSignatureVersion)"


        if ($Defender.AntivirusSignatureLastUpdated) {

            $Result.SignatureLastUpdated =
                "$($Defender.AntivirusSignatureLastUpdated)"

        }
        else {

            $Result.SignatureLastUpdated = "Unknown"
        }


        if ($Defender.DefenderSignaturesOutOfDate -eq $true) {

            $Result.SignaturesOutOfDate = "Yes"

        }
        elseif ($Defender.DefenderSignaturesOutOfDate -eq $false) {

            $Result.SignaturesOutOfDate = "No"

        }
        else {

            $Result.SignaturesOutOfDate = "Unknown"
        }


        $Result.TamperProtection =
            Get-EnabledStatus $Defender.IsTamperProtected

    }
    else {

        $Result.DefenderAVEnabled = "Unknown"

        $Result.DefenderRunningMode = "Unknown"

        $Result.RealTimeProtection = "Unknown"

        $Result.BehaviorMonitoring = "Unknown"

        $Result.IOAVProtection = "Unknown"

        $Result.NetworkInspection = "Unknown"

        $Result.SignatureVersion = "Unknown"

        $Result.SignatureLastUpdated = "Unknown"

        $Result.SignaturesOutOfDate = "Unknown"

        $Result.TamperProtection = "Unknown"
    }


    # ----------------------------------------------------------------------
    # MDE
    # ----------------------------------------------------------------------

    $Result.MDEOnboardingState = $MDEStatus


    if ($MDE) {

        $Result.MDEOnboardingValue =
            "$($MDE.OnboardingState)"

    }
    else {

        $Result.MDEOnboardingValue = "Unknown"
    }


    # ----------------------------------------------------------------------
    # Services
    # ----------------------------------------------------------------------

    if ($SenseService) {

        $Result.MDESenseService =
            "$($SenseService.Status)"

    }
    else {

        $Result.MDESenseService = "Not Found"
    }


    if ($DefenderService) {

        $Result.DefenderService =
            "$($DefenderService.Status)"

    }
    else {

        $Result.DefenderService = "Not Found"
    }


    # ----------------------------------------------------------------------
    # Firewall
    # ----------------------------------------------------------------------

    $Result.FirewallDomain =
        Get-EnabledStatus $DomainFirewall


    $Result.FirewallPrivate =
        Get-EnabledStatus $PrivateFirewall


    $Result.FirewallPublic =
        Get-EnabledStatus $PublicFirewall


    # ----------------------------------------------------------------------
    # Security Tools
    # ----------------------------------------------------------------------

    $Result.SecuritySoftware =
        $SecuritySoftware


    # ----------------------------------------------------------------------
    # Overall status
    # ----------------------------------------------------------------------

    $Result.ProtectionStatus =
        $ProtectionStatus


    $Result.ReviewNotes =
        $ReviewNotes


    return $Result
}


# ==========================================================================
# FUNCTION - SCAN THIS MASTER IMAGE
# ==========================================================================

function Export-CurrentImage {

    $Result = Get-CitrixSecurityStatus


    if (!$Result) {

        Write-Host ""
        Write-Host "No results were returned." `
            -ForegroundColor Red

        return
    }


    $OutputFile = Join-Path `
        $CsvFolder `
        "$($env:COMPUTERNAME)-SecurityStatus.csv"


    try {

        $Result |
            Export-Csv `
                -Path $OutputFile `
                -NoTypeInformation `
                -Encoding UTF8 `
                -Force `
                -ErrorAction Stop

    }
    catch {

        Write-Host ""
        Write-Host "ERROR writing CSV:" `
            -ForegroundColor Red

        Write-Host $OutputFile `
            -ForegroundColor Yellow

        Write-Host ""
        Write-Host $_.Exception.Message `
            -ForegroundColor Red

        return
    }


    Write-Host ""
    Write-Host "============================================================" `
        -ForegroundColor Green

    Write-Host " SCAN COMPLETE" `
        -ForegroundColor Green

    Write-Host "============================================================" `
        -ForegroundColor Green

    Write-Host ""


    Write-Host "Server:" `
        -NoNewline

    Write-Host " $($Result.ComputerName)" `
        -ForegroundColor Cyan


    Write-Host "Protection Status:" `
        -NoNewline

    if ($Result.ProtectionStatus -eq "Protected") {

        Write-Host " $($Result.ProtectionStatus)" `
            -ForegroundColor Green

    }
    else {

        Write-Host " $($Result.ProtectionStatus)" `
            -ForegroundColor Yellow
    }


    Write-Host ""

    Write-Host "CSV created:" `
        -ForegroundColor Cyan

    Write-Host $OutputFile `
        -ForegroundColor Yellow

    Write-Host ""


    $Result |
        Format-List
}


# ==========================================================================
# FUNCTION - BUILD COMBINED REPORT
# ==========================================================================

function New-CombinedReport {

    Write-Host ""
    Write-Host "============================================================" `
        -ForegroundColor Cyan

    Write-Host " BUILDING COMBINED SECURITY REPORT" `
        -ForegroundColor Cyan

    Write-Host "============================================================" `
        -ForegroundColor Cyan

    Write-Host ""


    # ----------------------------------------------------------------------
    # Find CSV files
    # ----------------------------------------------------------------------

    $CsvFiles = Get-ChildItem `
        -Path $CsvFolder `
        -Filter "*-SecurityStatus.csv" `
        -File `
        -ErrorAction SilentlyContinue


    if (!$CsvFiles) {

        Write-Host "No master image CSV files found." `
            -ForegroundColor Red

        Write-Host ""

        Write-Host "Expected location:" `
            -ForegroundColor Cyan

        Write-Host $CsvFolder `
            -ForegroundColor Yellow

        return
    }


    # ----------------------------------------------------------------------
    # Import CSVs
    # ----------------------------------------------------------------------

    $Data = @()


    foreach ($File in $CsvFiles) {

        try {

            $ImportedData = Import-Csv `
                -Path $File.FullName `
                -ErrorAction Stop


            $Data += $ImportedData

        }
        catch {

            Write-Host ""
            Write-Host "Unable to read:" `
                -ForegroundColor Yellow

            Write-Host $File.FullName `
                -ForegroundColor Yellow
        }
    }


    if (!$Data) {

        Write-Host ""
        Write-Host "No usable CSV data was found." `
            -ForegroundColor Red

        return
    }


    # ----------------------------------------------------------------------
    # Sort
    # ----------------------------------------------------------------------

    $Data = $Data |
        Sort-Object ComputerName


    # ----------------------------------------------------------------------
    # Export combined CSV
    # ----------------------------------------------------------------------

    try {

        $Data |
            Export-Csv `
                -Path $CombinedCsv `
                -NoTypeInformation `
                -Encoding UTF8 `
                -Force `
                -ErrorAction Stop

    }
    catch {

        Write-Host ""
        Write-Host "Unable to create combined CSV." `
            -ForegroundColor Red

        Write-Host $_.Exception.Message `
            -ForegroundColor Red

        return
    }


    # ----------------------------------------------------------------------
    # Totals
    # ----------------------------------------------------------------------

    $TotalMachines = @($Data).Count


    $ProtectedMachines = @(
        $Data |
        Where-Object {
            $_.ProtectionStatus -eq "Protected"
        }
    ).Count


    $ReviewMachines = @(
        $Data |
        Where-Object {
            $_.ProtectionStatus -ne "Protected"
        }
    ).Count


    $MDEOnboarded = @(
        $Data |
        Where-Object {
            $_.MDEOnboardingState -eq "Onboarded"
        }
    ).Count


    $RealTimeEnabled = @(
        $Data |
        Where-Object {
            $_.RealTimeProtection -eq "Enabled"
        }
    ).Count


    $ReportDate = Get-Date `
        -Format "dd MMMM yyyy HH:mm"


    # ======================================================================
    # BUILD TABLE ROWS
    # ======================================================================

    $HtmlRows = @()


    foreach ($Row in $Data) {


        # ------------------------------------------------------------------
        # CSS Classes
        # ------------------------------------------------------------------

        if ($Row.ProtectionStatus -eq "Protected") {

            $ProtectionClass = "good"

        }
        else {

            $ProtectionClass = "bad"
        }


        if ($Row.RealTimeProtection -eq "Enabled") {

            $RealTimeClass = "good"

        }
        else {

            $RealTimeClass = "bad"
        }


        if ($Row.BehaviorMonitoring -eq "Enabled") {

            $BehaviourClass = "good"

        }
        else {

            $BehaviourClass = "bad"
        }


        if ($Row.MDEOnboardingState -eq "Onboarded") {

            $MDEClass = "good"

        }
        else {

            $MDEClass = "bad"
        }


        if ($Row.MDESenseService -eq "Running") {

            $SenseClass = "good"

        }
        else {

            $SenseClass = "bad"
        }


        if ($Row.SignaturesOutOfDate -eq "No") {

            $SignatureClass = "good"

            $SignatureStatus = "Current"

        }
        else {

            $SignatureClass = "bad"

            $SignatureStatus = "Review"
        }


        if ($Row.TamperProtection -eq "Enabled") {

            $TamperClass = "good"

        }
        else {

            $TamperClass = "warning"
        }


        # ------------------------------------------------------------------
        # HTML safe values
        # ------------------------------------------------------------------

        $ComputerName =
            Convert-HtmlSafe $Row.ComputerName

        $ScanDate =
            Convert-HtmlSafe $Row.ScanDate

        $LanguageMode =
            Convert-HtmlSafe $Row.PowerShellLanguageMode

        $SecuritySoftware =
            Convert-HtmlSafe $Row.SecuritySoftware

        $DefenderMode =
            Convert-HtmlSafe $Row.DefenderRunningMode

        $RealTime =
            Convert-HtmlSafe $Row.RealTimeProtection

        $Behaviour =
            Convert-HtmlSafe $Row.BehaviorMonitoring

        $MDEState =
            Convert-HtmlSafe $Row.MDEOnboardingState

        $Sense =
            Convert-HtmlSafe $Row.MDESenseService

        $SignatureVersion =
            Convert-HtmlSafe $Row.SignatureVersion

        $SignatureUpdated =
            Convert-HtmlSafe $Row.SignatureLastUpdated

        $Tamper =
            Convert-HtmlSafe $Row.TamperProtection

        $FirewallDomain =
            Convert-HtmlSafe $Row.FirewallDomain

        $FirewallPrivate =
            Convert-HtmlSafe $Row.FirewallPrivate

        $FirewallPublic =
            Convert-HtmlSafe $Row.FirewallPublic

        $Protection =
            Convert-HtmlSafe $Row.ProtectionStatus

        $ReviewNotes =
            Convert-HtmlSafe $Row.ReviewNotes


        # ------------------------------------------------------------------
        # Add table row
        # ------------------------------------------------------------------

        $HtmlRows += @"
<tr>

<td class="computer">
$ComputerName
</td>

<td>
$ScanDate
</td>

<td>
$LanguageMode
</td>

<td>
$SecuritySoftware
</td>

<td>
$DefenderMode
</td>

<td class="$RealTimeClass">
$RealTime
</td>

<td class="$BehaviourClass">
$Behaviour
</td>

<td class="$MDEClass">
$MDEState
</td>

<td class="$SenseClass">
$Sense
</td>

<td>
$SignatureVersion
</td>

<td class="$SignatureClass">
$SignatureStatus
</td>

<td>
$SignatureUpdated
</td>

<td class="$TamperClass">
$Tamper
</td>

<td>
<strong>Domain:</strong> $FirewallDomain<br>
<strong>Private:</strong> $FirewallPrivate<br>
<strong>Public:</strong> $FirewallPublic
</td>

<td class="$ProtectionClass">
$Protection
</td>

<td>
$ReviewNotes
</td>

</tr>
"@

    }


    # ======================================================================
    # HTML REPORT
    # ======================================================================

    $HTML = @"
<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<title>
Citrix Security Protection Report
</title>


<style>

body {

    font-family:
        Segoe UI,
        Arial,
        sans-serif;

    margin: 30px;

    background-color: #f3f6f9;

    color: #263238;
}


.header {

    background-color: #25364a;

    color: white;

    padding: 25px 30px;

    border-radius: 8px;

    margin-bottom: 25px;
}


.header h1 {

    margin: 0;

    font-size: 28px;
}


.header p {

    margin-top: 8px;

    margin-bottom: 0;

    color: #d9e1e8;
}


.summary {

    display: flex;

    flex-wrap: wrap;

    gap: 18px;

    margin-bottom: 25px;
}


.card {

    background-color: white;

    min-width: 180px;

    padding: 18px 22px;

    border-radius: 8px;

    box-shadow:
        0 2px 7px rgba(0,0,0,0.08);
}


.card-title {

    color: #687078;

    font-size: 12px;

    text-transform: uppercase;

    font-weight: 600;
}


.card-value {

    font-size: 30px;

    font-weight: 700;

    margin-top: 6px;
}


.green-number {

    color: #188038;
}


.red-number {

    color: #c5221f;
}


.blue-number {

    color: #1967d2;
}


.table-container {

    background: white;

    overflow-x: auto;

    border-radius: 8px;

    box-shadow:
        0 2px 7px rgba(0,0,0,0.08);
}


table {

    border-collapse: collapse;

    width: 100%;

    min-width: 1750px;

    font-size: 12px;
}


th {

    background-color: #34495e;

    color: white;

    padding: 11px;

    text-align: left;

    white-space: nowrap;

    border-right: 1px solid #53687b;
}


td {

    padding: 10px;

    vertical-align: top;

    border-bottom: 1px solid #e2e6ea;

    border-right: 1px solid #eef1f3;
}


tr:nth-child(even) {

    background-color: #fafbfc;
}


tr:hover {

    background-color: #edf4fa;
}


.computer {

    font-weight: bold;

    white-space: nowrap;
}


.good {

    background-color: #e6f4ea;

    color: #137333;

    font-weight: 600;
}


.bad {

    background-color: #fce8e6;

    color: #c5221f;

    font-weight: 600;
}


.warning {

    background-color: #fef7e0;

    color: #8a5800;

    font-weight: 600;
}


.footer {

    color: #687078;

    font-size: 12px;

    margin-top: 20px;

    line-height: 1.6;
}


</style>

</head>


<body>


<div class="header">

<h1>
Citrix Security Protection Report
</h1>

<p>
Microsoft Defender and Defender for Endpoint protection posture
across the Citrix master image estate.
</p>

<p>
Report generated: $ReportDate
</p>

</div>


<div class="summary">


<div class="card">

<div class="card-title">
Master Images Scanned
</div>

<div class="card-value">
$TotalMachines
</div>

</div>


<div class="card">

<div class="card-title">
Protected
</div>

<div class="card-value green-number">
$ProtectedMachines
</div>

</div>


<div class="card">

<div class="card-title">
Requires Review
</div>

<div class="card-value red-number">
$ReviewMachines
</div>

</div>


<div class="card">

<div class="card-title">
MDE Onboarded
</div>

<div class="card-value blue-number">
$MDEOnboarded
</div>

</div>


<div class="card">

<div class="card-title">
Real-Time Protection
</div>

<div class="card-value green-number">
$RealTimeEnabled
</div>

</div>


</div>


<div class="table-container">

<table>


<thead>

<tr>

<th>
Master Image
</th>

<th>
Scan Date
</th>

<th>
PowerShell Mode
</th>

<th>
Security Tooling
</th>

<th>
Defender Mode
</th>

<th>
Real-Time
</th>

<th>
Behaviour
</th>

<th>
MDE Status
</th>

<th>
Sense / EDR
</th>

<th>
Signature Version
</th>

<th>
Signature Status
</th>

<th>
Signature Updated
</th>

<th>
Tamper Protection
</th>

<th>
Windows Firewall
</th>

<th>
Protection Status
</th>

<th>
Review Notes
</th>

</tr>

</thead>


<tbody>

$($HtmlRows -join "`n")

</tbody>


</table>

</div>


<div class="footer">

<strong>
Report Location:
</strong>

$ReportRoot

<br>

<strong>
Protected
</strong>

means the expected Microsoft Defender Antivirus,
real-time protection and Microsoft Defender for Endpoint
components were detected and healthy when the image was scanned.

<br>

This report represents the state of each Citrix master image
at the time its individual scan was performed.

</div>


</body>

</html>
"@


    # ----------------------------------------------------------------------
    # Save HTML
    # ----------------------------------------------------------------------

    try {

        $HTML |
            Out-File `
                -FilePath $HtmlReport `
                -Encoding UTF8 `
                -Force `
                -ErrorAction Stop

    }
    catch {

        Write-Host ""
        Write-Host "ERROR: Unable to create HTML report." `
            -ForegroundColor Red

        Write-Host $_.Exception.Message `
            -ForegroundColor Red

        return
    }


    # ----------------------------------------------------------------------
    # Results
    # ----------------------------------------------------------------------

    Write-Host ""
    Write-Host "============================================================" `
        -ForegroundColor Green

    Write-Host " COMBINED REPORT COMPLETE" `
        -ForegroundColor Green

    Write-Host "============================================================" `
        -ForegroundColor Green

    Write-Host ""


    Write-Host "Master Images Scanned : $TotalMachines"

    Write-Host "Protected             : $ProtectedMachines" `
        -ForegroundColor Green

    Write-Host "Requires Review       : $ReviewMachines" `
        -ForegroundColor Yellow

    Write-Host "MDE Onboarded         : $MDEOnboarded" `
        -ForegroundColor Cyan

    Write-Host "Real-Time Enabled     : $RealTimeEnabled" `
        -ForegroundColor Cyan


    Write-Host ""

    Write-Host "Combined CSV:" `
        -ForegroundColor Cyan

    Write-Host $CombinedCsv `
        -ForegroundColor Yellow


    Write-Host ""

    Write-Host "HTML Report:" `
        -ForegroundColor Cyan

    Write-Host $HtmlReport `
        -ForegroundColor Yellow

    Write-Host ""
}


# ==========================================================================
# MAIN MENU
# ==========================================================================

do {

    Clear-Host


    Write-Host ""
    Write-Host "============================================================" `
        -ForegroundColor Cyan

    Write-Host " CITRIX MASTER IMAGE AV / SECURITY AUDIT" `
        -ForegroundColor Cyan

    Write-Host "============================================================" `
        -ForegroundColor Cyan

    Write-Host ""


    Write-Host "Current Server :" `
        -NoNewline

    Write-Host " $env:COMPUTERNAME" `
        -ForegroundColor Yellow


    Write-Host "PowerShell Mode:" `
        -NoNewline

    Write-Host " $($ExecutionContext.SessionState.LanguageMode)" `
        -ForegroundColor Yellow


    Write-Host "Output Folder :" `
        -NoNewline

    Write-Host " $ReportRoot" `
        -ForegroundColor Yellow


    Write-Host ""
    Write-Host "------------------------------------------------------------"
    Write-Host ""

    Write-Host "1 - Scan this master image and create CSV"

    Write-Host ""

    Write-Host "2 - Combine all existing CSVs and create HTML report"

    Write-Host ""

    Write-Host "3 - Scan this image AND rebuild combined HTML report"

    Write-Host ""

    Write-Host "4 - Open AV report folder"

    Write-Host ""

    Write-Host "5 - Exit"

    Write-Host ""


    $Choice = Read-Host "Select an option"


    switch ($Choice) {


        "1" {

            Export-CurrentImage

            Write-Host ""

            Read-Host "Press Enter to return to menu"
        }


        "2" {

            New-CombinedReport

            Write-Host ""

            Read-Host "Press Enter to return to menu"
        }


        "3" {

            Export-CurrentImage

            New-CombinedReport

            Write-Host ""

            Read-Host "Press Enter to return to menu"
        }


        "4" {

            try {

                Start-Process `
                    explorer.exe `
                    $ReportRoot

            }
            catch {

                Write-Host ""
                Write-Host "Unable to open Explorer." `
                    -ForegroundColor Yellow

                Write-Host $ReportRoot
            }


            Start-Sleep -Seconds 1
        }


        "5" {

            Write-Host ""
            Write-Host "Exiting." `
                -ForegroundColor Cyan

        }


        default {

            Write-Host ""
            Write-Host "Invalid selection." `
                -ForegroundColor Red

            Start-Sleep -Seconds 2
        }

    }


} until ($Choice -eq "5")
