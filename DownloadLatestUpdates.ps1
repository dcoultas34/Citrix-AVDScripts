[CmdletBinding()]
param(
    [string]$ShareRoot = "\\transfer\transfer\CitrixApps",
    [switch]$WhatIf,
    [switch]$CleanVersionFolder
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

$RunDate      = Get-Date
$ManifestPath = Join-Path $ShareRoot "CitrixAppsManifest.csv"
$LogPath      = Join-Path $ShareRoot "DownloadLog.txt"
$Manifest     = @()

function Write-Log {
    param([string]$Message)

    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $Message
    Add-Content -Path $LogPath -Value $line
}

function Test-WingetReady {
    try {
        $null = Get-Command winget -ErrorAction Stop
        $null = winget --version
        return $true
    } catch {
        return $false
    }
}

function Get-SafeFolderName {
    param([string]$Name)
    return ($Name -replace '[\\/:*?"<>|]', '').Trim()
}

function Get-WingetLatestVersion {
    param([string]$WingetId)

    try {
        $out = winget show --id $WingetId --exact --source winget --accept-source-agreements 2>$null

        if (-not $out) {
            return $null
        }

        $verLine = ($out -split "`r?`n") |
            Where-Object { $_ -match "^\s*Version\s*:" } |
            Select-Object -First 1

        if ($verLine) {
            return ($verLine -split ":\s*", 2)[1].Trim()
        }
    } catch {}

    return $null
}

function Invoke-WingetDownload {
    param(
        [string]$WingetId,
        [string]$TargetFolder,
        [switch]$WhatIf
    )

    $args = @(
        "download",
        "--id", $WingetId,
        "--exact",
        "--source", "winget",
        "--download-directory", $TargetFolder,
        "--accept-package-agreements",
        "--accept-source-agreements"
    )

    if ($WhatIf) {
        Write-Log "WhatIf: winget $($args -join ' ')"
        return 0
    }

    try {
        $proc = Start-Process `
            -FilePath "winget" `
            -ArgumentList $args `
            -Wait `
            -PassThru `
            -NoNewWindow

        return $proc.ExitCode
    } catch {
        Write-Log "ERROR: winget download failed for $WingetId - $($_.Exception.Message)"
        return -1
    }
}

if (-not (Test-WingetReady)) {
    throw "winget is not installed or not working on this machine."
}

if (-not (Test-Path $ShareRoot)) {
    New-Item -ItemType Directory -Path $ShareRoot -Force | Out-Null
}

Add-Content -Path $LogPath -Value ""
Add-Content -Path $LogPath -Value "==== Citrix app download started: $RunDate ===="

foreach ($app in $AppsToDownload) {
    $name = $app.Name
    $id   = $app.WingetId

    Write-Host ""
    Write-Host "Processing $name..." -ForegroundColor Cyan

    $safeName  = Get-SafeFolderName -Name $name
    $appFolder = Join-Path $ShareRoot $safeName

    if (-not (Test-Path $appFolder)) {
        New-Item -ItemType Directory -Path $appFolder -Force | Out-Null
    }

    $latest = Get-WingetLatestVersion -WingetId $id

    if (-not $latest) {
        Write-Log "WARNING: $name - could not determine latest version."

        $Manifest += [pscustomobject]@{
            Application = $name
            WingetId    = $id
            Version     = "-"
            Status      = "Failed - version unknown"
            ExitCode    = "-"
            Folder      = $appFolder
            Files       = "-"
            Downloaded  = Get-Date
        }

        continue
    }

    $safeVersion   = Get-SafeFolderName -Name $latest
    $versionFolder = Join-Path $appFolder $safeVersion

    if ($CleanVersionFolder -and (Test-Path $versionFolder) -and -not $WhatIf) {
        Remove-Item -Path $versionFolder -Recurse -Force
    }

    if (-not (Test-Path $versionFolder)) {
        New-Item -ItemType Directory -Path $versionFolder -Force | Out-Null
    }

    $before = @(Get-ChildItem -Path $versionFolder -Recurse -File -ErrorAction SilentlyContinue)

    $exitCode = Invoke-WingetDownload `
        -WingetId $id `
        -TargetFolder $versionFolder `
        -WhatIf:$WhatIf

    $after = @(Get-ChildItem -Path $versionFolder -Recurse -File -ErrorAction SilentlyContinue)

    if ($after.Count -gt 0) {
        if ($before.Count -gt 0) {
            $beforeNames = @($before | ForEach-Object { $_.FullName })
            $afterNames  = @($after  | ForEach-Object { $_.FullName })

            $newFiles = @(
                Compare-Object `
                    -ReferenceObject $beforeNames `
                    -DifferenceObject $afterNames `
                    -PassThru |
                Where-Object { $_ }
            )

            if (-not $newFiles -or $newFiles.Count -eq 0) {
                $newFiles = $afterNames
            }
        } else {
            $newFiles = @($after | ForEach-Object { $_.FullName })
        }
    } else {
        $newFiles = @()
    }

    if ($WhatIf) {
        $status = "WhatIf"
    } elseif ($exitCode -eq 0 -and $after.Count -gt 0) {
        $status = "Downloaded"
    } elseif ($exitCode -eq 0 -and $after.Count -eq 0) {
        $status = "Completed but no files found"
    } else {
        $status = "Failed"
    }

    $Manifest += [pscustomobject]@{
        Application = $name
        WingetId    = $id
        Version     = $latest
        Status      = $status
        ExitCode    = $exitCode
        Folder      = $versionFolder
        Files       = if ($newFiles.Count -gt 0) { ($newFiles -join "; ") } else { "-" }
        Downloaded  = Get-Date
    }

    Write-Log "$name $latest - $status - ExitCode $exitCode"
}

$Manifest |
    Sort-Object Application |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ManifestPath

Write-Host ""
Write-Host "Download run complete." -ForegroundColor Green
Write-Host "Share path: $ShareRoot"
Write-Host "Manifest: $ManifestPath"
Write-Host "Log: $LogPath"
