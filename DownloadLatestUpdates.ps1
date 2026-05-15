[CmdletBinding()]
param(
    [string]$ShareRoot = "\\transfer\transfer\CitrixApps",
    [switch]$WhatIf
)

$AppsToDownload = @(
    @{ Name="Adobe Acrobat Reader"; WingetId="Adobe.Acrobat.Reader.64-bit" }
    @{ Name="Azure Data Studio";    WingetId="Microsoft.AzureDataStudio" }
    @{ Name="GitHub Desktop";       WingetId="GitHub.GitHubDesktop" }
    @{ Name="Google Chrome";        WingetId="Google.Chrome" }
    @{ Name="Microsoft Edge";       WingetId="Microsoft.Edge" }
    @{ Name="Microsoft Teams";      WingetId="Microsoft.Teams" }
    @{ Name="OneDrive";             WingetId="Microsoft.OneDrive" }
    @{ Name="Power BI Desktop";     WingetId="Microsoft.PowerBI" }
    @{ Name="Visual Studio Code";   WingetId="Microsoft.VisualStudioCode" }
)

$RunDate = Get-Date
$Manifest = @()
$LogPath = Join-Path $ShareRoot "DownloadLog.txt"
$ManifestPath = Join-Path $ShareRoot "CitrixAppsManifest.csv"

function Test-WingetReady {
    try {
        $null = Get-Command winget -ErrorAction Stop
        winget --version | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Get-WingetLatestVersion {
    param([string]$WingetId)

    try {
        $out = winget show --id $WingetId --exact --source winget --accept-source-agreements 2>$null

        $verLine = ($out -split "`r?`n") |
            Where-Object { $_ -match "^\s*Version\s*:" } |
            Select-Object -First 1

        if ($verLine) {
            return ($verLine -split ":\s*", 2)[1].Trim()
        }
    } catch {}

    return $null
}

if (-not (Test-WingetReady)) {
    throw "winget is not installed or not working on this machine."
}

if (-not (Test-Path $ShareRoot)) {
    New-Item -ItemType Directory -Path $ShareRoot -Force | Out-Null
}

Add-Content -Path $LogPath -Value "`n==== Citrix app download started: $RunDate ===="

foreach ($app in $AppsToDownload) {
    $name = $app.Name
    $id   = $app.WingetId

    Write-Host "`nProcessing $name..." -ForegroundColor Cyan

    $safeName = $name -replace '[\\/:*?"<>|]', ''
    $appFolder = Join-Path $ShareRoot $safeName

    if (-not (Test-Path $appFolder)) {
        New-Item -ItemType Directory -Path $appFolder -Force | Out-Null
    }

    $latest = Get-WingetLatestVersion -WingetId $id
    if (-not $latest) {
        Write-Warning "$name - could not determine latest version."
        Add-Content -Path $LogPath -Value "$name - failed to determine latest version."
        continue
    }

    $versionFolder = Join-Path $appFolder $latest
    if (-not (Test-Path $versionFolder)) {
        New-Item -ItemType Directory -Path $versionFolder -Force | Out-Null
    }

    $before = Get-ChildItem -Path $versionFolder -Recurse -File -ErrorAction SilentlyContinue

    $args = @(
        "download",
        "--id", $id,
        "--exact",
        "--source", "winget",
        "--download-directory", $versionFolder,
        "--accept-package-agreements",
        "--accept-source-agreements"
    )

    if ($WhatIf) {
        Write-Host "WhatIf: winget $($args -join ' ')" -ForegroundColor Yellow
        $exitCode = 0
        $status = "WhatIf"
    } else {
        $proc = Start-Process -FilePath "winget" -ArgumentList $args -Wait -PassThru -NoNewWindow
        $exitCode = $proc.ExitCode
        $status = if ($exitCode -eq 0) { "Downloaded" } else { "Failed" }
    }

    $after = Get-ChildItem -Path $versionFolder -Recurse -File -ErrorAction SilentlyContinue
    $newFiles = Compare-Object $before.FullName $after.FullName -PassThru |
        Where-Object { $_ }

    if (-not $newFiles) {
        $newFiles = $after.FullName
    }

    $Manifest += [pscustomobject]@{
        Application = $name
        WingetId    = $id
        Version     = $latest
        Status      = $status
        ExitCode    = $exitCode
        Folder      = $versionFolder
        Files       = ($newFiles -join "; ")
        Downloaded  = Get-Date
    }

    Add-Content -Path $LogPath -Value "$name $latest - $status - ExitCode $exitCode"
}

$Manifest |
    Sort-Object Application |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ManifestPath

Write-Host "`nDownload complete." -ForegroundColor Green
Write-Host "Share path: $ShareRoot"
Write-Host "Manifest: $ManifestPath"
Write-Host "Log: $LogPath"