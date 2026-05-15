# Download-CitrixApps-Direct.ps1


[CmdletBinding()]
param(
    [string]$ShareRoot = "C:\AppUpdates",
    [switch]$WhatIf
)

# =====================================================
# Direct vendor download URLs
# =====================================================
# NOTE:
# Some vendors use permanent redirect links which always
# point to the latest release.
#
# Teams intentionally excluded for now because Microsoft
# changes packaging/service model frequently.
# Office 365 and Windows Updates also excluded.
# =====================================================

$Apps = @(

    @{
        Name = "Google Chrome"
        Url  = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
        File = "GoogleChromeEnterprise64.msi"
    }

    @{
        Name = "Microsoft Edge"
        Url  = "https://go.microsoft.com/fwlink/?linkid=2109047"
        File = "MicrosoftEdgeEnterpriseX64.msi"
    }

    @{
        Name = "Visual Studio Code"
        Url  = "https://update.code.visualstudio.com/latest/win32-x64/stable"
        File = "VSCode-x64.exe"
    }

    @{
        Name = "Power BI Desktop"
        Url  = "https://download.microsoft.com/download/9/5/A/95A641C1-7D0F-4F9D-B9BE-8C3F4E708AA4/PBIDesktopSetup_x64.exe"
        File = "PBIDesktopSetup_x64.exe"
    }

    @{
        Name = "Azure Data Studio"
        Url  = "https://aka.ms/azuredatastudio-windows-user-setup"
        File = "AzureDataStudio.exe"
    }

    @{
        Name = "GitHub Desktop"
        Url  = "https://central.github.com/deployments/desktop/desktop/latest/win32"
        File = "GitHubDesktop.exe"
    }

    @{
        Name = "OneDrive"
        Url  = "https://go.microsoft.com/fwlink/p/?LinkID=2182910"
        File = "OneDriveSetup.exe"
    }

    @{
        Name = "Adobe Acrobat Reader"
        Url  = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2600121529/AcroRdrDCx642600121529_en_US.exe"
        File = "AdobeReader.exe"
    }
)

# =====================================================
# Paths
# =====================================================

$LogPath      = Join-Path $ShareRoot "DownloadLog.txt"
$ManifestPath = Join-Path $ShareRoot "CitrixAppsManifest.csv"
$Manifest     = @()

# =====================================================
# Functions
# =====================================================

function Write-Log {
    param([string]$Message)

    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message

    Write-Host $Message
    Add-Content -Path $LogPath -Value $line
}

function Get-SafeFolderName {
    param([string]$Name)

    return ($Name -replace '[\\/:*?"<>|]', '').Trim()
}

function Get-FileVersion {
    param([string]$Path)

    try {
        if (Test-Path $Path) {
            return (Get-Item $Path).VersionInfo.ProductVersion
        }
    } catch {}

    return "Unknown"
}

# =====================================================
# Startup
# =====================================================

if (-not (Test-Path $ShareRoot)) {
    New-Item -ItemType Directory -Path $ShareRoot -Force | Out-Null
}

Add-Content -Path $LogPath -Value ""
Add-Content -Path $LogPath -Value "==== Direct vendor download run started: $(Get-Date) ===="

# =====================================================
# Download loop
# =====================================================

foreach ($app in $Apps) {

    $name = $app.Name
    $url  = $app.Url
    $file = $app.File

    Write-Host ""
    Write-Host "Processing $name..." -ForegroundColor Cyan

    $safeName = Get-SafeFolderName -Name $name
    $appPath  = Join-Path $ShareRoot $safeName

    if (-not (Test-Path $appPath)) {
        New-Item -ItemType Directory -Path $appPath -Force | Out-Null
    }

    $targetFile = Join-Path $appPath $file

    try {

        if ($WhatIf) {
            Write-Host "WhatIf: would download $url" -ForegroundColor Yellow
            $status = "WhatIf"
        }
        else {

            Write-Host "Downloading from: $url" -ForegroundColor DarkGray

            Invoke-WebRequest `
                -Uri $url `
                -OutFile $targetFile `
                -UseBasicParsing `
                -MaximumRedirection 10 `
                -ErrorAction Stop

            $status = "Downloaded"
        }

        $version = if (Test-Path $targetFile) {
            Get-FileVersion -Path $targetFile
        }
        else {
            "Unknown"
        }

        $Manifest += [pscustomobject]@{
            Application = $name
            Version     = $version
            Status      = $status
            FilePath    = $targetFile
            Url         = $url
            Downloaded  = Get-Date
        }

        Write-Log "$name - $status - $version"
    }
    catch {

        $Manifest += [pscustomobject]@{
            Application = $name
            Version     = "-"
            Status      = "Failed"
            FilePath    = $targetFile
            Url         = $url
            Downloaded  = Get-Date
        }

        Write-Log "ERROR: $name failed - $($_.Exception.Message)"
    }
}

# =====================================================
# Export manifest
# =====================================================

$Manifest |
    Sort-Object Application |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ManifestPath

Write-Host ""
Write-Host "Download run complete." -ForegroundColor Green
Write-Host "Share:    $ShareRoot"
Write-Host "Manifest: $ManifestPath"
Write-Host "Log:      $LogPath"
```

## Notes

* This version avoids `winget download` completely.
* It uses direct vendor download URLs instead.
* Teams intentionally excluded because Microsoft packaging changes frequently.
* Office 365 and Windows Updates should be handled separately.
* Adobe Reader URL will need occasional maintenance because Adobe changes build numbers.
* The script is designed to stage installers centrally for Citrix image patching.
