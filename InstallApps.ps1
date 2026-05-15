[CmdletBinding()]
param(
    [string]$ShareRoot = "\\transfer\transfer\CitrixApps",
    [switch]$WhatIf,
    [switch]$InstalledOnly
)

$ManifestPath = Join-Path $ShareRoot "CitrixAppsManifest.csv"
$LogPath      = Join-Path $env:PUBLIC "CitrixApps-OfflineInstall.log"

$SilentArgs = @{
    "Google Chrome"        = "/qn /norestart"
    "Microsoft Edge"       = "/qn /norestart"
    "Visual Studio Code"   = "/VERYSILENT /NORESTART /MERGETASKS=!runcode"
    "Power BI Desktop"     = "/quiet ACCEPT_EULA=1"
    "Azure Data Studio"    = "/VERYSILENT /NORESTART"
    "GitHub Desktop"       = "/silent"
    "OneDrive"             = "/silent /allusers"
    "Adobe Acrobat Reader" = "/sAll /rs /rps /msi EULA_ACCEPT=YES"
}

function Write-Log {
    param([string]$Message)

    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Write-Host $Message
    Add-Content -Path $LogPath -Value $line
}

function Get-InstalledAppVersion {
    param([string]$DisplayNameMatch)

    $roots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }

        foreach ($sub in Get-ChildItem $root -ErrorAction SilentlyContinue) {
            try {
                $p = Get-ItemProperty $sub.PSPath -ErrorAction Stop
                if ($p.DisplayName -like "*$DisplayNameMatch*") {
                    return $p.DisplayVersion
                }
            } catch {}
        }
    }

    return $null
}

function Get-AppMatchName {
    param([string]$Application)

    switch ($Application) {
        "Google Chrome"        { "Google Chrome" }
        "Microsoft Edge"       { "Microsoft Edge" }
        "Visual Studio Code"   { "Microsoft Visual Studio Code" }
        "Power BI Desktop"     { "Microsoft Power BI Desktop" }
        "Azure Data Studio"    { "Azure Data Studio" }
        "GitHub Desktop"       { "GitHub Desktop" }
        "OneDrive"             { "Microsoft OneDrive" }
        "Adobe Acrobat Reader" { "Adobe Acrobat Reader" }
        default                { $Application }
    }
}

function Invoke-Installer {
    param(
        [string]$Application,
        [string]$FilePath,
        [switch]$WhatIf
    )

    if (-not (Test-Path $FilePath)) {
        Write-Log "ERROR: Installer not found for $Application - $FilePath"
        return -1
    }

    $ext = [IO.Path]::GetExtension($FilePath).ToLowerInvariant()

    if ($ext -eq ".msi") {
        $args = "/i `"$FilePath`" $($SilentArgs[$Application])"

        if ($WhatIf) {
            Write-Log "WhatIf: msiexec.exe $args"
            return 0
        }

        $proc = Start-Process `
            -FilePath "msiexec.exe" `
            -ArgumentList $args `
            -Wait `
            -PassThru `
            -NoNewWindow

        return $proc.ExitCode
    }

    if ($ext -eq ".exe") {
        $args = $SilentArgs[$Application]

        if (-not $args) {
            Write-Log "ERROR: No silent arguments defined for $Application"
            return -1
        }

        if ($WhatIf) {
            Write-Log "WhatIf: `"$FilePath`" $args"
            return 0
        }

        $proc = Start-Process `
            -FilePath $FilePath `
            -ArgumentList $args `
            -Wait `
            -PassThru `
            -NoNewWindow

        return $proc.ExitCode
    }

    Write-Log "ERROR: Unsupported installer type for $Application - $FilePath"
    return -1
}

if (-not (Test-Path $ManifestPath)) {
    throw "Manifest not found: $ManifestPath"
}

$manifest = Import-Csv $ManifestPath

Add-Content -Path $LogPath -Value ""
Add-Content -Path $LogPath -Value "==== Offline install run started: $(Get-Date) ===="

foreach ($item in $manifest) {
    $app = $item.Application
    $filePath = $item.FilePath

    Write-Host ""
    Write-Host "Processing $app..." -ForegroundColor Cyan

    if (-not $filePath -or $filePath -eq "-") {
        Write-Log "Skipping $app - no installer path in manifest."
        continue
    }

    $matchName = Get-AppMatchName -Application $app
    $installedBefore = Get-InstalledAppVersion -DisplayNameMatch $matchName

    if ($InstalledOnly -and -not $installedBefore) {
        Write-Log "Skipping $app - not currently installed on this image."
        continue
    }

    $exitCode = Invoke-Installer `
        -Application $app `
        -FilePath $filePath `
        -WhatIf:$WhatIf

    Start-Sleep -Seconds 3

    $installedAfter = Get-InstalledAppVersion -DisplayNameMatch $matchName

    if ($exitCode -eq 0 -or $exitCode -eq 3010) {
        Write-Log "$app installed/updated successfully. Before: $installedBefore After: $installedAfter ExitCode: $exitCode"
    }
    else {
        Write-Log "$app failed. Before: $installedBefore After: $installedAfter ExitCode: $exitCode"
    }
}

Write-Host ""
Write-Host "Offline install run complete." -ForegroundColor Green
Write-Host "Log: $LogPath"