# This script will check the below defined apps and update each to the latest version of Evergreen / winget

param(
  [switch]$Upgrade,                 # perform upgrades (installed apps only)
  [switch]$IncludeWindowsUpdate,    # try OS updates (requires PSWindowsUpdate)
  [switch]$WindowsUpdateOnly,       # do only Windows Updates, skip app upgrades
  [switch]$WhatIf,                  # dry-run the actions
  [switch]$NoHtml,                  # never write/open the HTML report
  [switch]$HtmlOnlyWhenGreen,       # write/open HTML only when all installed apps are green
  [switch]$NoCsv,                   # skip CSV export
  [switch]$NoReports,               # shorthand: implies -NoCsv and -NoHtml
  [switch]$DisableMsStore           # disable 'msstore' source in winget for this session/user
)

# Shorthand
if ($NoReports) { $NoCsv = $true; $NoHtml = $true }

<# =======================
   Apps to check / update
   ======================= #>
$AppsToCheck = @(
    @{ Name="Windows Updates";         LocalMatch="__WINDOWS_UPDATE__";           LatestProvider="None";      EvergreenName=$null;                       WingetId=$null }

    @{ Name="Microsoft Office 365";    LocalMatch="__OFFICE_C2R__";               LatestProvider="Winget";    EvergreenName=$null;                       WingetId="Microsoft.Office" }

    # Evergreen-first (vendor data), winget fallback
    @{ Name="Google Chrome";           LocalMatch="Google Chrome";                LatestProvider="Evergreen"; EvergreenName="GoogleChrome";              WingetId="Google.Chrome" }
    @{ Name="Microsoft Edge";          LocalMatch="Microsoft Edge";               LatestProvider="Evergreen"; EvergreenName="MicrosoftEdge";              WingetId="Microsoft.Edge" }
    @{ Name="Visual Studio Code";      LocalMatch="Microsoft Visual Studio Code"; LatestProvider="Evergreen"; EvergreenName="MicrosoftVisualStudioCode"; WingetId="Microsoft.VisualStudioCode" }
    @{ Name="Power BI Desktop";        LocalMatch="Microsoft Power BI Desktop";   LatestProvider="Evergreen"; EvergreenName="MicrosoftPowerBIDesktop";    WingetId="Microsoft.PowerBI" }
    @{ Name="Visual Studio";           LocalMatch="Microsoft Visual Studio";      LatestProvider="Evergreen"; EvergreenName="MicrosoftVisualStudio";      WingetId="Microsoft.VisualStudio" }

    # Winget for latest
    @{ Name="Adobe Acrobat Reader";    LocalMatch="Adobe Acrobat Reader";         LatestProvider="Winget";    EvergreenName=$null;                       WingetId="Adobe.Acrobat.Reader.64-bit" }

    # Full Acrobat tracked separately (no latest lookup)
    @{ Name="Adobe Acrobat (full)";    LocalMatch="Adobe Acrobat";                LatestProvider="None";      EvergreenName=$null;                       WingetId=$null }

    @{ Name="Microsoft Teams";         LocalMatch="Microsoft Teams";              LatestProvider="Winget";    EvergreenName=$null;                       WingetId="Microsoft.Teams" }
    @{ Name="OneDrive";                LocalMatch="Microsoft OneDrive";           LatestProvider="Winget";    EvergreenName=$null;                       WingetId="Microsoft.OneDrive" }
    @{ Name="GitHub Desktop";          LocalMatch="GitHub Desktop";               LatestProvider="Winget";    EvergreenName=$null;                       WingetId="GitHub.GitHubDesktop" }
    @{ Name="Azure Data Studio";       LocalMatch="Azure Data Studio";            LatestProvider="Winget";    EvergreenName=$null;                       WingetId="Microsoft.AzureDataStudio" }
)

$ReportPath  = Join-Path $env:PUBLIC "AVD-AppUpdateReport.html"
$CsvPath     = Join-Path $env:PUBLIC "AVD-AppUpdateReport.csv"
$LogPath     = Join-Path $env:PUBLIC "AVD-AppUpdateActions.log"

# Prefer D:\_Source\_AppUpdates for Evergreen downloads, with fallbacks
$DownloadDirCandidates = @(
  "D:\_Source\_AppUpdates",
  "C:\_Source\_AppUpdates",
  (Join-Path $env:ProgramData "AVD-AppDownloads")
)
$DownloadDir = $null
foreach ($p in $DownloadDirCandidates) {
    try {
        if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
        $DownloadDir = $p; break
    } catch { }
}
if (-not $DownloadDir) { throw "Could not create a download directory in any candidate path." }
Write-Host "Evergreen download directory: $DownloadDir" -ForegroundColor DarkGray

# ---------------- Environment / winget helpers ----------------
$WingetMsStoreTlsErr = -1978335138  # 0x8A15005E

function Test-WingetReady {
    try { $null = Get-Command winget -ErrorAction Stop } catch { return $false }
    try { winget --version | Out-Null; return $true } catch { return $false }
}

function Test-EvergreenReady {
    try {
        $m = Get-Module -ListAvailable -Name Evergreen
        if ($m) { Import-Module Evergreen -ErrorAction SilentlyContinue | Out-Null; return $true }
    } catch {}
    return $false
}

function Disable-MsStoreSource {
    try {
        $src = winget source list 2>$null
        if ($src -match '^\s*msstore\s+' -and $src -match 'Enabled') {
            Write-Host "winget: disabling 'msstore' source for this user..." -ForegroundColor DarkGray
            winget source disable msstore | Out-Null
        }
    } catch {}
}

# ---------------- App-specific helpers ----------------
function Get-GitHubDesktopVersion {
    try {
        $root = Join-Path $env:LOCALAPPDATA "GitHubDesktop"
        if (-not (Test-Path $root)) { return $null }
        $exe = Get-ChildItem -Path $root -Filter "GitHubDesktop.exe" -Recurse -ErrorAction SilentlyContinue |
               Sort-Object FullName -Descending | Select-Object -First 1
        if ($exe) { return (Get-Item $exe.FullName).VersionInfo.ProductVersion }
    } catch {}
    return $null
}

function Get-VSCodeVersion {
    $paths = @(
        "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe",     # user setup
        "$env:ProgramFiles\Microsoft VS Code\Code.exe",              # system (x64)
        "${env:ProgramFiles(x86)}\Microsoft VS Code\Code.exe"        # system (x86)
    )
    foreach ($p in $paths) {
        if (Test-Path $p) {
            try { return (Get-Item $p).VersionInfo.ProductVersion } catch {}
        }
    }
    return $null
}

function Get-VisualStudioVersion {
    param(
        [int]$MajorVersion = 17  # VS 2022
    )

    $roots = @(
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $candidates = @()

    foreach ($rt in $roots) {
        if (-not (Test-Path $rt)) { continue }

        foreach ($sub in Get-ChildItem $rt -ErrorAction SilentlyContinue) {
            $p = Get-ItemProperty $sub.PSPath -ErrorAction SilentlyContinue
            if (-not $p) { continue }

            if ($null -eq $p.VersionMajor -or $p.VersionMajor -ne $MajorVersion) { continue }

            $dn = $p.DisplayName
            if (-not $dn) { continue }

            if ($dn -notmatch 'Visual Studio' -and $dn -notmatch '^vs_') { continue }

            $dv   = $p.DisplayVersion
            $vObj = $null
            if ($dv) {
                try { $vObj = [version]($dv -replace '[^\d\.]','') } catch {}
            }

            $candidates += [pscustomobject]@{
                DisplayName    = $dn
                DisplayVersion = $dv
                VersionObj     = $vObj
                KeyPath        = $sub.PSPath
            }
        }
    }

    if ($candidates.Count -eq 0) { return $null }

    ($candidates |
        Sort-Object @{ Expression = "VersionObj"; Descending = $true } |
        Select-Object -First 1).DisplayVersion
}

# ---------------- Installed version (smart) ----------------
function Get-InstalledAppVersion {
    param(
        [Parameter(Mandatory)][string]$DisplayNameMatch,
        [string[]]$ExcludePatterns = $null
    )

    # --- Microsoft Teams (MSIX) - prefer package version (pick newest if many) ---
    if ($DisplayNameMatch -eq "Microsoft Teams") {
        $packages = @()
        foreach ($n in 'MSTeams','MicrosoftTeams') {
            try { $packages += Get-AppxPackage -AllUsers -Name $n -ErrorAction SilentlyContinue } catch {}
            try { $packages += Get-AppxPackage          -Name $n -ErrorAction SilentlyContinue } catch {}
        }
        $pkg = $packages | Where-Object { $_ } | Sort-Object Version -Descending | Select-Object -First 1
        if ($pkg) { return $pkg.Version.ToString() }
        # fall through to registry for legacy Teams
    }

    # --- Office C2R direct registry read ---
    if ($DisplayNameMatch -eq "__OFFICE_C2R__") {
        $key = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        if (Test-Path $key) {
            $v = (Get-ItemProperty $key -ErrorAction SilentlyContinue).VersionToReport
            if ($v) { return $v }
        }
        return $null
    }

    # --- Visual Studio 2022 - use dedicated helper based on VersionMajor ---
    if ($DisplayNameMatch -eq "Microsoft Visual Studio") {
        $vsver = Get-VisualStudioVersion -MajorVersion 17
        if ($vsver) { return $vsver }
        # fall through to generic registry scan if helper fails
    }

    # --- VS Code - prefer Code.exe file version ---
    if ($DisplayNameMatch -eq "Microsoft Visual Studio Code") {
        $vsver = Get-VSCodeVersion
        if ($vsver) { return $vsver }
        # fall through to registry if not found
    }

    # --- GitHub Desktop - prefer app file version ---
    if ($DisplayNameMatch -eq "GitHub Desktop") {
        $gh = Get-GitHubDesktopVersion
        if ($gh) { return $gh }
    }

    # --- Generic registry scan (noise filters applied) ---
    $roots = @(
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $nameExcludes = @(
        'WebView','Runtime','WebView2','Updater','Update','AutoUpdate',
        'Maintenance','Service','Helper','Crashpad','Stub',
        'Machine-wide','User Installer','System Installer','Setup'
    )

    $candidates = @()

    foreach ($rt in $roots) {
        if (-
