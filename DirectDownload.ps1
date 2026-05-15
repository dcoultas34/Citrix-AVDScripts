[CmdletBinding()]
param(
    [string]$ShareRoot = "\\transfer\transfer\CitrixApps",
    [switch]$WhatIf
)

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
        Url  = "https://www.microsoft.com/en-us/download/details.aspx?id=58494"
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

$LogPath      = Join-Path $ShareRoot "DownloadLog.txt"
$ManifestPath = Join-Path $ShareRoot "CitrixAppsManifest.csv"
$Manifest     = @()

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

function Resolve-PowerBiDownloadUrl {
    param([string]$DownloadPageUrl)

    $page = Invoke-WebRequest `
        -Uri $DownloadPageUrl `
        -UseBasicParsing `
        -MaximumRedirection 10 `
        -ErrorAction Stop

    $realUrl = ($page.Links |
        Where-Object { $_.href -match "PBIDesktopSetup_x64\.exe" } |
        Select-Object -First 1 -ExpandProperty href)

    if (-not $realUrl) {
        $realUrl = ($page.Content -split '"' |
            Where-Object { $_ -match "https://download\.microsoft\.com/.+PBIDesktopSetup_x64\.exe" } |
            Select-Object -First 1)
    }

    if (-not $realUrl) {
        throw "Could not resolve Power BI Desktop x64 download URL from Microsoft Download Center."
    }

    return $realUrl
}

function Invoke-AppDownload {
    param(
        [string]$Name,
        [string]$Url,
        [string]$TargetFile
    )

    if ($Name -eq "Power BI Desktop") {
        Write-Host "Resolving Power BI Desktop download URL..." -ForegroundColor DarkGray
        $Url = Resolve-PowerBiDownloadUrl -DownloadPageUrl $Url
        Write-Host "Resolved Power BI URL: $Url" -ForegroundColor DarkGray
    }

    Invoke-WebRequest `
        -Uri $Url `
        -OutFile $TargetFile `
        -UseBasicParsing `
        -MaximumRedirection 10 `
        -ErrorAction Stop

    $fileInfo = Get-Item $TargetFile -ErrorAction Stop

    if ($Name -eq "Power BI Desktop" -and $fileInfo.Length -lt 100MB) {
        Remove-Item $TargetFile -Force -ErrorAction SilentlyContinue
        throw "Power BI download is too small and likely invalid: $($fileInfo.Length) bytes"
    }

    if ($fileInfo.Length -lt 1KB) {
        Remove-Item $TargetFile -Force -ErrorAction SilentlyContinue
        throw "$Name download is too small and likely invalid: $($fileInfo.Length) bytes"
    }
}

if (-not (Test-Path $ShareRoot)) {
    New-Item -ItemType Directory -Path $ShareRoot -Force | Out-Null
}

Add-Content -Path $LogPath -Value ""
Add-Content -Path $LogPath -Value "==== Direct vendor download run started: $(Get-Date) ===="

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
            Write-Host "WhatIf: would download $url to $targetFile" -ForegroundColor Yellow
            $status = "WhatIf"
        }
        else {
            if (Test-Path $targetFile) {
                Remove-Item $targetFile -Force -ErrorAction SilentlyContinue
            }

            Write-Host "Downloading from: $url" -ForegroundColor DarkGray

            Invoke-AppDownload `
                -Name $name `
                -Url $url `
                -TargetFile $targetFile

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

$Manifest |
    Sort-Object Application |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ManifestPath

Write-Host ""
Write-Host "Download run complete." -ForegroundColor Green
Write-Host "Share:    $ShareRoot"
Write-Host "Manifest: $ManifestPath"
Write-Host "Log:      $LogPath"
