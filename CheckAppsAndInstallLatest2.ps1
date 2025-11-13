# This script will check the below defined apps and update each to the latest version of evergreen

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
        if ($m) {
            Import-Module Evergreen -ErrorAction SilentlyContinue | Out-Null
            return $true
        }
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

    # --- Microsoft Teams (MSIX) ---
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

    # --- Visual Studio 2022 via VersionMajor ---
    if ($DisplayNameMatch -eq "Microsoft Visual Studio") {
        $vsver = Get-VisualStudioVersion -MajorVersion 17
        if ($vsver) { return $vsver }
        # if not found, fall through to generic scan
    }

    # --- VS Code via Code.exe ---
    if ($DisplayNameMatch -eq "Microsoft Visual Studio Code") {
        $vsver = Get-VSCodeVersion
        if ($vsver) { return $vsver }
    }

    # --- GitHub Desktop via file version ---
    if ($DisplayNameMatch -eq "GitHub Desktop") {
        $gh = Get-GitHubDesktopVersion
        if ($gh) { return $gh }
    }

    # --- Generic registry scan ---
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
        if (-not (Test-Path $rt)) { continue }
        foreach ($sub in Get-ChildItem $rt -ErrorAction SilentlyContinue) {
            $p = Get-ItemProperty $sub.PSPath -ErrorAction SilentlyContinue
            $dn = $p.DisplayName
            if (-not $dn) { continue }
            if ($dn -notlike "*$DisplayNameMatch*") { continue }

            $skip = $false
            foreach ($ex in $nameExcludes) {
                if ($dn -match [regex]::Escape($ex)) { $skip = $true; break }
            }
            if ($skip) { continue }
            if ($ExcludePatterns) {
                foreach ($ex in $ExcludePatterns) {
                    if ($dn -match [regex]::Escape($ex)) { $skip = $true; break }
                }
            }
            if ($skip) { continue }

            $dv = $p.DisplayVersion
            $score = if ($dn -ieq $DisplayNameMatch) { 3 }
                     elseif ($dn -ilike "$DisplayNameMatch*") { 2 }
                     else { 1 }
            $vObj = $null
            if ($dv) {
                try { $vObj = [version]($dv -replace '[^\d\.]','') } catch {}
            }
            $candidates += [pscustomobject]@{
                DisplayName    = $dn
                DisplayVersion = $dv
                Score          = $score
                VersionObj     = $vObj
                KeyPath        = $sub.PSPath
            }
        }
    }
    if ($candidates.Count -eq 0) { return $null }

    # Sort by version first, then by score (so newest wins)
    ($candidates |
        Sort-Object @{ Expression = "VersionObj"; Descending = $true },
                    @{ Expression = "Score";      Descending = $true } |
        Select-Object -First 1).DisplayVersion
}

# ---------------- Latest version (forces winget source) ----------------
function Get-LatestWingetVersion {
    param([Parameter(Mandatory)][string]$WingetId)
    $out = winget search --id $WingetId --exact --source winget --accept-source-agreements 2>$null
    if ($out) {
        $line = ($out -split "`r?`n" | Where-Object { $_ -match [regex]::Escape($WingetId) -and $_ -match "\|" } | Select-Object -First 1)
        if ($line) {
            $parts = $line -split "\|"
            if ($parts.Count -ge 3) { return $parts[2].Trim() }
        }
    }
    $out2 = winget show --id $WingetId --exact --source winget --accept-source-agreements 2>$null
    if ($out2) {
        $verLine = ($out2 -split "`r?`n") | Where-Object { $_ -match "^\s*Version\s*:" } | Select-Object -First 1
        if ($verLine) { return ($verLine -split ":\s*",2)[1].Trim() }
    }
    return $null
}
function Get-LatestEvergreenVersion {
    param([Parameter(Mandatory)][string]$EvergreenName)
    try {
        $data = Get-EvergreenApp -Name $EvergreenName -ErrorAction Stop
        if (-not $data) { return $null }
        if ($null -ne ($data | Get-Member -Name Channel -ErrorAction SilentlyContinue)) {
            $stable = $data | Where-Object { $_.Channel -match 'Stable' }
            if ($stable) { $data = $stable }
        }
        $data | Where-Object { $_.Version } |
            Sort-Object { try { [version]($_.Version -replace '[^\d\.]','') } catch { $_.Version } } -Descending |
            Select-Object -First 1 -ExpandProperty Version
    } catch { return $null }
}

# ---------------- Version compare ----------------
function Compare-VersionSmart {
    param([string]$Installed,[string]$Latest)
    if (-not $Installed -or -not $Latest) { return $null }
    try {
        $v1 = [version]($Installed -replace '[^\d\.]','')
        $v2 = [version]($Latest    -replace '[^\d\.]','')
        [Math]::Sign($v1.CompareTo($v2))
    } catch {
        if ($Installed -eq $Latest) { 0 } else { -1 }
    }
}

# ---------------- Windows Update row ----------------
function Get-WindowsUpdateStatus {
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $result = $searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
        $pending = $result.Updates.Count
    } catch { $pending = $null }

    $os = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    $installedTxt = "$($os.DisplayVersion) (Build $($os.CurrentBuild).$($os.UBR))"

    if ($pending -gt 0) {
        [PSCustomObject]@{
            Installed = $installedTxt
            Latest    = "Updates available ($pending)"
            Status    = "Update available"
            UpgradeTo = "Apply Windows Updates"
        }
    }
    elseif ($pending -eq 0) {
        [PSCustomObject]@{
            Installed = $installedTxt
            Latest    = "-"
            Status    = "Up-to-date"
            UpgradeTo = "-"
        }
    }
    else {
        [PSCustomObject]@{
            Installed = $installedTxt
            Latest    = "-"
            Status    = "Unknown (WU query failed)"
            UpgradeTo = "-"
        }
    }
}

# ---------------- Evergreen download/install ----------------
$EvergreenSilentArgs = @{
  "MicrosoftVisualStudioCode" = "/VERYSILENT /NORESTART"
  "MicrosoftPowerBIDesktop"   = "/quiet ACCEPT_EULA=1"
}
function Get-EvergreenAsset {
    param([Parameter(Mandatory)][string]$EvergreenName)
    try {
        $data = Get-EvergreenApp -Name $EvergreenName -ErrorAction Stop
        if (-not $data) { return $null }
        if ($null -ne ($data | Get-Member -Name Channel -ErrorAction SilentlyContinue)) {
            $data = $data | Where-Object { $_.Channel -match 'Stable' -or -not $_.Channel }
        }
        $arch = if ([Environment]::Is64BitOperatingSystem) { "x64|amd64" } else { "x86|i386" }
        if ($null -ne ($data | Get-Member -Name Architecture -ErrorAction SilentlyContinue)) {
            $c = $data | Where-Object { $_.Architecture -match $arch }
            if ($c) { $data = $c }
        }
        $msi = $data | Where-Object { $_.Uri -match '\.msi($|\?)' } | Select-Object -First 1
        if ($msi) { return $msi }
        $exe = $data | Where-Object { $_.Uri -match '\.exe($|\?)' } | Select-Object -First 1
        if ($exe) { return $exe }
    } catch {}
    return $null
}
function Install-AppEvergreen {
    param(
      [Parameter(Mandatory)][string]$EvergreenName,
      [Parameter(Mandatory)][string]$AppDisplayName,
      [switch]$WhatIf
    )
    $asset = Get-EvergreenAsset -EvergreenName $EvergreenName
    if (-not $asset -or -not $asset.Uri) { return [pscustomobject]@{ Used="Evergreen"; Result="NoAsset"; ExitCode=-1 } }
    $uri  = $asset.Uri
    $file = Join-Path $DownloadDir ([IO.Path]::GetFileName((($uri -split '\?')[0])))

    Write-Host "$AppDisplayName: downloading from $uri" -ForegroundColor DarkCyan
    if ($WhatIf) {
        Write-Host "WhatIf: would download to $file" -ForegroundColor Gray
    } else {
        try {
            try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
            Invoke-WebRequest -UseBasicParsing -Uri $uri -OutFile $file -ErrorAction Stop
        } catch {
            Write-Warning "$AppDisplayName: download failed: $($_.Exception.Message)"
            return [pscustomobject]@{ Used="Evergreen"; Result="DownloadFailed"; ExitCode=-1 }
        }
    }

    $ext = [IO.Path]::GetExtension($file).ToLowerInvariant()
    if ($ext -eq ".msi") {
        $args = "/i `"$file`" /qn /norestart"
        if ($WhatIf) {
            Write-Host "WhatIf: msiexec $args" -ForegroundColor Gray
            return [pscustomobject]@{ Used="Evergreen"; Result="WouldRun"; ExitCode=0 }
        }
        $p = Start-Process msiexec -ArgumentList $args -Wait -PassThru -NoNewWindow
        return [pscustomobject]@{ Used="Evergreen"; Result="RanMSI"; ExitCode=$p.ExitCode }
    }
    elseif ($ext -eq ".exe") {
        $exeArgs = $null
        if ($EvergreenSilentArgs.ContainsKey($EvergreenName)) {
            $exeArgs = "$($EvergreenSilentArgs[$EvergreenName])"
        } else {
            Write-Host "$AppDisplayName: unknown silent args for Evergreen EXE; will fall back to winget." -ForegroundColor Yellow
            return [pscustomobject]@{ Used="Evergreen"; Result="UnknownExeArgs"; ExitCode=-1; File=$file }
        }
        if ($WhatIf) {
            Write-Host "WhatIf: `"$file`" $exeArgs" -ForegroundColor Gray
            return [pscustomobject]@{ Used="Evergreen"; Result="WouldRun"; ExitCode=0 }
        }
        $p = Start-Process -FilePath $file -ArgumentList $exeArgs -Wait -PassThru -NoNewWindow
        return [pscustomobject]@{ Used="Evergreen"; Result="RanEXE"; ExitCode=$p.ExitCode }
    }
    else {
        [pscustomobject]@{ Used="Evergreen"; Result="UnsupportedType"; ExitCode=-1 }
    }
}

# --------- Office & Teams helpers ---------
function Update-OfficeC2R {
    param([switch]$WhatIf)
    $c2r = Join-Path "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun" "OfficeC2RClient.exe"
    if (-not (Test-Path $c2r)) {
        return [pscustomobject]@{ Found=$false; ExitCode=-1; Result="C2R client not found" }
    }
    $args = "/update user displaylevel=false forceappshutdown=true"
    if ($WhatIf) {
        Write-Host "WhatIf: `"$c2r`" $args" -ForegroundColor Gray
        return [pscustomobject]@{ Found=$true; ExitCode=0; Result="WouldRun" }
    }
    Write-Host "Microsoft Office 365: running OfficeC2RClient update..." -ForegroundColor Cyan
    $p = Start-Process -FilePath $c2r -ArgumentList $args -NoNewWindow -PassThru -Wait
    [pscustomobject]@{
        Found    = $true
        ExitCode = $p.ExitCode
        Result   = (if ($p.ExitCode -eq 0) { "Success" } else { "ExitCode $($p.ExitCode)" })
    }
}
function Stop-TeamsIfRunning {
    $names = @("ms-teams","ms-teams-updater","Teams","MSTeams")
    foreach ($n in $names) {
        Get-Process -Name $n -ErrorAction SilentlyContinue | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch {}
        }
    }
}

# ---------------- Upgrades WITH PROGRESS ----------------
function Invoke-AppUpgrades {
    param(
        [Parameter(Mandatory)][array]$Results,
        [Parameter(Mandatory)][array]$AppConfig,
        [switch]$IncludeWindowsUpdate,
        [switch]$WhatIf,
        [switch]$WindowsOnly
    )

    Add-Content -Path $LogPath -Value "`n==== $(Get-Date) ===="

    if ($DisableMsStore -and (Test-WingetReady)) { Disable-MsStoreSource }

    if ($IncludeWindowsUpdate) {
        $wu = $Results | Where-Object Application -eq "Windows Updates"
        if ($wu -and $wu.Status -eq "Update available") {
            Write-Progress -Activity "Windows Update" -Status "Installing available updates..." -PercentComplete 0
            try {
                if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
                    Import-Module PSWindowsUpdate -ErrorAction SilentlyContinue | Out-Null
                    if ($WhatIf) {
                        Write-Host "WhatIf: Get-WindowsUpdate -AcceptAll -Install -AutoReboot:`$false" -ForegroundColor Gray
                    } else {
                        Get-WindowsUpdate -AcceptAll -Install -AutoReboot:$false | Tee-Object -FilePath $LogPath -Append | Out-Null
                    }
                } else {
                    Write-Warning "PSWindowsUpdate not installed; skipping OS updates."
                }
            } catch {
                Write-Warning "Windows Update install failed: $($_.Exception.Message)"
            }
            Write-Progress -Activity "Windows Update" -Completed
        } else {
            Write-Host "Windows Update: no pending updates detected." -ForegroundColor DarkGray
        }
        if ($WindowsOnly) {
            Write-Host "Windows Update-only run: skipping application upgrades." -ForegroundColor DarkGray
            return
        }
    }

    $queue = @($Results | Where-Object { $_.Application -ne "Windows Updates" -and $_.Status -eq "Update available" })
    $total = $queue.Count
    if ($total -eq 0) {
        Write-Host "No application upgrades required." -ForegroundColor Green
        return
    }

    Write-Host ("Upgrading {0} application(s)..." -f $total) -ForegroundColor Cyan
    $i = 0
    $actions = @()

    foreach ($r in $queue) {
        $i++
        $cfg = $AppConfig | Where-Object { $_.Name -eq $r.Application } | Select-Object -First 1
        if (-not $cfg) { continue }
        $label = "[{0}/{1}] {2}" -f $i,$total,$cfg.Name
        $den = [Math]::Max(1, $total)
        $pct = [int]([Math]::Max(0, [Math]::Min(100, ((($i - 1) * 100.0) / $den)) ))

        Write-Progress -Activity "Upgrading applications" -Status ("{0} -- preparing" -f $label) -PercentComplete $pct
        $fromVer = $r.Installed
        $method = "-"
        $exit = $null
        $result = "Skipped"

        if ($cfg.Name -eq "Microsoft Office 365") {
            $office = Update-OfficeC2R -WhatIf:$WhatIf
            $method = "OfficeC2R"
            $exit = $office.ExitCode
            $result = $office.Result
            Add-Content -Path $LogPath -Value ("Office 365: OfficeC2R {0}, exit {1}" -f $result,$exit)
        }
        else {
            $evergreenOk = (Get-Module -Name Evergreen -ListAvailable) -ne $null
            if ($cfg.LatestProvider -eq 'Evergreen' -and $cfg.EvergreenName -and $evergreenOk) {
                Write-Host ("{0}: Evergreen -> download/install..." -f $cfg.Name) -ForegroundColor Cyan
                Write-Progress -Activity "Upgrading applications" -Status ("{0} -- downloading (Evergreen)" -f $label) -PercentComplete $pct
                $ev = Install-AppEvergreen -EvergreenName $cfg.EvergreenName -AppDisplayName $cfg.Name -WhatIf:$WhatIf
                if ($ev.Result -in @("RanMSI","RanEXE","WouldRun")) {
                    $method = "Evergreen"
                    $exit = $ev.ExitCode
                    $result = if ($WhatIf) { "Simulated" } else { if ($exit -eq 0) { "Success" } else { "ExitCode $exit" } }
                    Add-Content -Path $LogPath -Value ("{0}: Evergreen {1} exit {2}" -f $cfg.Name,$ev.Result,$ev.ExitCode)
                }
            }

            if ($cfg.Name -eq "Microsoft Teams" -and -not $WhatIf) { Stop-TeamsIfRunning }

            if ($method -eq "-" -and $cfg.WingetId -and (Test-WingetReady)) {
                Write-Host ("{0}: winget -> upgrade..." -f $cfg.Name) -ForegroundColor Cyan
                Write-Progress -Activity "Upgrading applications" -Status ("{0} -- running winget" -f $label) -PercentComplete $pct

                $args = @(
                    "upgrade","--id",$cfg.WingetId,"--exact","--source","winget",
                    "--accept-package-agreements","--accept-source-agreements","--silent"
                )
                if ($WhatIf) {
                    Write-Host ("WhatIf: winget {0}" -f ($args -join ' ')) -ForegroundColor Gray
                    Add-Content -Path $LogPath -Value ("WhatIf: winget {0}" -f ($args -join ' '))
                    $method = "winget"
                    $exit = 0
                    $result = "Simulated"
                } else {
                    try {
                        $proc = Start-Process -FilePath "winget" -ArgumentList $args -NoNewWindow -PassThru -Wait
                        $method = "winget"
                        $exit = $proc.ExitCode
                        if ($proc.ExitCode -eq 0) {
                            $result = "Success"
                        }
                        elseif ($proc.ExitCode -eq $WingetMsStoreTlsErr) {
                            $result = "Source error (msstore/TLS). Use -DisableMsStore or fix proxy certs."
                        }
                        else {
                            $result = "ExitCode $($proc.ExitCode)"
                        }

                        if ($cfg.Name -eq "Microsoft Teams" -and $proc.ExitCode -ne 0) {
                            Write-Host "Microsoft Teams: retry with 'winget install --force'..." -ForegroundColor DarkYellow
                            $args2 = @(
                                "install","--id",$cfg.WingetId,"--exact","--source","winget",
                                "--accept-package-agreements","--accept-source-agreements","--silent","--force"
                            )
                            $proc2 = Start-Process -FilePath "winget" -ArgumentList $args2 -NoNewWindow -PassThru -Wait
                            if ($proc2.ExitCode -eq 0) {
                                $result = "Success (repair/install)"
                                $exit = 0
                                $method = "winget-install"
                            }
                            Add-Content -Path $LogPath -Value ("Teams install --force exit {0}" -f $proc2.ExitCode)
                        }

                        Add-Content -Path $LogPath -Value ("{0}: winget exitcode {1}" -f $cfg.Name,$proc.ExitCode)
                    } catch {
                        $method = "winget"
                        $exit = -1
                        $result = "Error"
                        Write-Warning ("{0}: winget failed - {1}" -f $cfg.Name,$_.Exception.Message)
                        Add-Content -Path $LogPath -Value ("{0}: winget error {1}" -f $cfg.Name,$_.Exception.Message)
                    }
                }
            }
            elseif ($method -eq "-" -and -not $cfg.WingetId) {
                Write-Warning ("{0}: no Evergreen path and no WingetId - cannot upgrade." -f $cfg.Name)
                $result = "No installer path"
            }
        }

        Write-Progress -Activity "Upgrading applications" -Status ("{0} -- verifying" -f $label) -PercentComplete $pct
        $toVer = $fromVer
        if (-not $WhatIf) {
            $exclude = $null
            if ($cfg.Name -eq "Adobe Acrobat (full)") { $exclude=@('Reader') }
            $toVer = Get-InstalledAppVersion -DisplayNameMatch $cfg.LocalMatch -ExcludePatterns $exclude
        }

        if ($result -eq "Success" -and $toVer -and $toVer -ne $fromVer) {
            Write-Host ("{0}: upgraded {1} -> {2} via {3}" -f $cfg.Name,$fromVer,$toVer,$method) -ForegroundColor Green
        }
        elseif ($result -like "Success (repair/install)*") {
            Write-Host ("{0}: repaired/installed to {1} via {2}" -f $cfg.Name,$toVer,$method) -ForegroundColor Green
        }
        elseif ($result -eq "Simulated") {
            Write-Host ("{0}: would upgrade {1} -> {2} via {3}" -f $cfg.Name,$fromVer,$r.UpgradeTo,$method) -ForegroundColor Yellow
        }
        else {
            Write-Host ("{0}: {1} (method: {2}, from: {3}, to: {4})" -f $cfg.Name,$result,$method,$fromVer,$toVer) -ForegroundColor DarkYellow
        }

        $actions += [pscustomobject]@{
            Application = $cfg.Name
            Method      = $method
            From        = $fromVer
            To          = $toVer
            Result      = $result
            ExitCode    = (if ($exit -ne $null) { $exit } else { "-" })
        }
    }

    Write-Progress -Activity "Upgrading applications" -Completed
    Write-Host "`nUpgrade summary:" -ForegroundColor White
    $actions | Format-Table Application,Method,From,To,Result,ExitCode -AutoSize
    foreach ($a in $actions) {
        Add-Content -Path $LogPath -Value ("{0}: {1} via {2} ({3} -> {4}) exit {5}" -f $a.Application,$a.Result,$a.Method,$a.From,$a.To,$a.ExitCode)
    }
    Write-Host "Action log: $LogPath"
}

# -------------- Build results --------------
function Get-AppResults {
    param([Parameter(Mandatory)][array]$Apps)
    $wingetOk = Test-WingetReady
    $evergreenOk = Test-EvergreenReady

    if ($wingetOk -and $DisableMsStore) { Disable-MsStoreSource }
    if (-not $wingetOk) {
        Write-Warning "winget not found/working. Winget lookups/installs will be skipped."
    }
    if (-not $evergreenOk) {
        Write-Warning "Evergreen module not available. Evergreen lookups/installs will be skipped."
    }

    $out = foreach ($app in $Apps) {
        if ($app.LocalMatch -eq "__WINDOWS_UPDATE__") {
            $wu = Get-WindowsUpdateStatus
            [PSCustomObject]@{
                Application = $app.Name
                Installed   = $wu.Installed
                Latest      = $wu.Latest
                Status      = $wu.Status
                UpgradeTo   = $wu.UpgradeTo
                WingetId    = "-"
            }
            continue
        }
        $exclude = $null
        if ($app.Name -eq "Adobe Acrobat (full)") { $exclude=@('Reader') }
        $installed = Get-InstalledAppVersion -DisplayNameMatch $app.LocalMatch -ExcludePatterns $exclude
        if (-not $installed) {
            [PSCustomObject]@{
                Application = $app.Name
                Installed   = "-"
                Latest      = "-"
                Status      = "Application not installed"
                UpgradeTo   = "-"
                WingetId    = (if ($app.WingetId) { $app.WingetId } else { "-" })
            }
            continue
        }

        $latest = $null
        if ($app.LatestProvider -eq 'Evergreen' -and $evergreenOk -and $app.EvergreenName) {
            $latest = Get-LatestEvergreenVersion -EvergreenName $app.EvergreenName
        }
        if (-not $latest -and $wingetOk -and $app.WingetId) {
            $latest = Get-LatestWingetVersion -WingetId $app.WingetId
        }

        $cmp = Compare-VersionSmart -Installed $installed -Latest $latest
        if (-not $latest) {
            $status = "Unknown (no latest info)"
        }
        elseif ($cmp -lt 0) {
            $status = "Update available"
        }
        elseif ($cmp -eq 0) {
            $status = "Up-to-date"
        }
        else {
            $status = "Ahead of catalog"
        }

        [PSCustomObject]@{
            Application = $app.Name
            Installed   = $installed
            Latest      = (if ($latest) { $latest } else { "-" })
            Status      = $status
            UpgradeTo   = (if ($status -eq "Update available" -and $latest) { $latest } else { "-" })
            WingetId    = (if ($app.WingetId) { $app.WingetId } else { "-" })
        }
    }
    $out | Sort-Object Application
}

# -------------- HTML writer --------------
function Write-ReportHtml {
    param([Parameter(Mandatory)][array]$Results)
    $css = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; }
h1 { font-size: 20px; margin-bottom: 12px; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
th { background: #f6f6f6; }
tr.status-uptodate { background: #e8f5e9; }
tr.status-update   { background: #ffebee; }
tr.status-missing  { background: #eceff1; }
tr.status-unknown  { background: #f3e5f5; }
.badge { padding: 2px 8px; border-radius: 999px; font-size: 12px; }
.badge-ok      { background: #2e7d32; color: white; }
.badge-update  { background: #c62828; color: white; }
.badge-missing { background: #607d8b; color: white; }
.badge-unknown { background: #6a1b9a; color: white; }
.small { color: #666; font-size: 12px; }
</style>
"@
    $rows = foreach ($r in $Results) {
        $class="status-unknown"; $badgeClass="badge-unknown"
        switch -Regex ($r.Status) {
            "Up-to-date"               { $class="status-uptodate"; $badgeClass="badge-ok"; break }
            "Update available"         { $class="status-update";   $badgeClass="badge-update"; break }
            "Application not installed"{ $class="status-missing";  $badgeClass="badge-missing"; break }
            default                    { $class="status-unknown";  $badgeClass="badge-unknown" }
        }
        "<tr class='$class'><td>$($r.Application)</td><td>$($r.Installed)</td><td>$($r.Latest)</td><td><span class='badge $badgeClass'>$($r.Status)</span></td><td>$($r.UpgradeTo)</td><td class='small'>$($r.WingetId)</td></tr>"
    }
    $html = @"
<html>
<head><meta charset='utf-8'><title>AVD App Update Report</title>$css</head>
<body>
<h1>AVD App Update Report</h1>
<table>
<thead><tr><th>Application</th><th>Installed Version</th><th>Latest Version</th><th>Status</th><th>Upgrade To</th><th>Winget ID</th></tr></thead>
<tbody>
$(($rows -join "`n"))
</tbody>
</table>
<p class='small'>Generated: $(Get-Date)</p>
</body>
</html>
"@
    $html | Set-Content -Path $ReportPath -Encoding UTF8
    Write-Host "HTML report saved to: $ReportPath"
    try {
        if ($env:USERNAME -ne "SYSTEM") {
            Start-Process -FilePath $ReportPath -ErrorAction Stop
        }
    } catch {
        Write-Host "Could not auto-open the report: $($_.Exception.Message)" -ForegroundColor DarkGray
    }
}

# ----------------- MAIN -----------------
$results = Get-AppResults -Apps $AppsToCheck

# Console output
$fmt = "{0,-28} {1,-20} {2,-20} {3,-24} {4,-14}"
Write-Host ""
Write-Host ($fmt -f "Application","Installed","Latest","Status","Upgrade To")
Write-Host ("-" * 110)
foreach ($r in $results) {
    $fg = "Gray"
    if     ($r.Status -eq "Up-to-date")                { $fg = "Green" }
    elseif ($r.Status -eq "Update available")          { $fg = "Red" }
    elseif ($r.Status -eq "Application not installed") { $fg = "DarkGray" }
    elseif ($r.Status -like "Unknown*")                { $fg = "Yellow" }
    Write-Host ($fmt -f $r.Application,$r.Installed,$r.Latest,$r.Status,$r.UpgradeTo) -ForegroundColor $fg
}

# CSV (respect -NoCsv)
if (-not $NoCsv) {
    $results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $CsvPath
    Write-Host "`nCSV saved to: $CsvPath"
} else {
    Write-Host "Skipping CSV output." -ForegroundColor DarkGray
}

# HTML gating BEFORE upgrades (respect -NoHtml)
$pendingHtml = $false
if (-not $NoHtml) {
    # ignore "Application not installed" when deciding if it's green
    $issuesNow = $results |
      Where-Object { $_.Status -ne "Application not installed" } |
      Where-Object { $_.Status -eq "Update available" -or $_.Status -like "Unknown*" }

    if ($HtmlOnlyWhenGreen) {
        if ($issuesNow.Count -eq 0 -and -not $Upgrade) {
            Write-ReportHtml -Results $results
        }
        elseif ($Upgrade) {
            $pendingHtml = $true
            Write-Host "Skipping HTML for now; will generate after upgrades if everything goes green." -ForegroundColor DarkYellow
        }
    } else {
        Write-ReportHtml -Results $results
    }
} else {
    Write-Host "Skipping HTML report." -ForegroundColor DarkGray
}

# Upgrades
if ($Upgrade) {
    Invoke-AppUpgrades -Results $results -AppConfig $AppsToCheck -IncludeWindowsUpdate:$IncludeWindowsUpdate -WhatIf:$WhatIf -WindowsOnly:$WindowsUpdateOnly
    if ($pendingHtml -and -not $NoHtml) {
        $resultsPost = Get-AppResults -Apps $AppsToCheck
        $issuesPost  = $resultsPost |
          Where-Object { $_.Status -ne "Application not installed" } |
          Where-Object { $_.Status -eq "Update available" -or $_.Status -like "Unknown*" }
        if ($issuesPost.Count -eq 0) {
            Write-ReportHtml -Results $resultsPost
        }
        else {
            Write-Host ("Still not fully green; HTML not generated. ({0} item(s) still need attention.)" -f $issuesPost.Count) -ForegroundColor DarkYellow
        }
    }
}
