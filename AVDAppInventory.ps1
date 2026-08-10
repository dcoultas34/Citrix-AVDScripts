param(
  [string]$CsvPath = (Join-Path $PSScriptRoot "Widgets_installed_programs_updated.csv"),
  [string]$ReportRoot = "C:\Ultima\AVDAppInventory"
)

$ErrorActionPreference = 'SilentlyContinue'

# ---------------------- Paths ----------------------
$today = Get-Date
$period = $today.ToString("MMMM yyyy")
$monthFolder = Join-Path $ReportRoot $period

if (-not (Test-Path -LiteralPath $monthFolder)) {
  New-Item -ItemType Directory -Path $monthFolder -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $CsvPath)) {
  Write-Error "CSV file not found: $CsvPath"
  exit 1
}

$machine = $env:COMPUTERNAME
$todayIso = $today.ToString('yyyy-MM-dd')

# ---------------------- Evergreen map ----------------------
# Only applications with a known Evergreen application name are queried.
$EvergreenMap = @{
  '7-Zip 25.01 (x64)' = '7Zip'
  'Adobe Acrobat (64-bit)' = 'AdobeAcrobat'
  'Audacity 3.7.7' = 'Audacity'
  'FileZilla 3.69.5' = 'FileZilla'
  'Google Chrome' = 'GoogleChrome'
  'Microsoft 365 Apps for enterprise - en-us' = 'Microsoft365Apps'
  'Microsoft 365 Apps for enterprise - fr-fr' = 'Microsoft365Apps'
  'Microsoft Edge' = 'MicrosoftEdge'
  'Microsoft FSLogix Apps' = 'MicrosoftFSLogixApps'
  'Microsoft Power BI Desktop (x64)' = 'MicrosoftPowerBIDesktop'
  'Mozilla Firefox ESR (x64 en-US)' = 'MozillaFirefox'
  'Notepad++ (64-bit x64)' = 'NotepadPlusPlus'
  'VLC media player' = 'VideoLanVlc'
  'Zoom Workplace (64-bit)' = 'Zoom'
}

function Test-EvergreenReady {
  try {
    $module = Get-Module -ListAvailable -Name Evergreen | Select-Object -First 1
    if (-not $module) { return $false }
    Import-Module Evergreen -ErrorAction Stop | Out-Null
    return $true
  }
  catch {
    return $false
  }
}

function Get-LatestEvergreenVersion {
  param([string]$EvergreenName)

  if (-not $EvergreenName) {
    return $null
  }

  try {
    $data = Get-EvergreenApp -Name $EvergreenName -ErrorAction Stop

    if (-not $data) {
      return $null
    }

    # Prefer x64 / Windows / stable style entries where those properties exist.
    $filtered = @($data)

    if ($filtered[0].PSObject.Properties.Name -contains 'Architecture') {
      $x64 = @($filtered | Where-Object { $_.Architecture -match 'x64|64' })
      if ($x64.Count -gt 0) {
        $filtered = $x64
      }
    }

    if ($filtered[0].PSObject.Properties.Name -contains 'Channel') {
      $stable = @($filtered | Where-Object { $_.Channel -match 'Stable|Enterprise|Monthly' })
      if ($stable.Count -gt 0) {
        $filtered = $stable
      }
    }

    $versions = foreach ($item in $filtered) {
      if ($item.Version) {
        $versionObject = $null
        try {
          $versionObject = [version](("$($item.Version)" -replace '[^\d\.]','').Trim('.'))
        } catch {}

        [pscustomobject]@{
          VersionText = "$($item.Version)"
          VersionObj  = $versionObject
        }
      }
    }

    if (-not $versions) {
      return $null
    }

    $latest = $versions |
      Sort-Object @{Expression={ if ($_.VersionObj) { $_.VersionObj } else { [version]'0.0' } }; Descending=$true} |
      Select-Object -First 1

    return $latest.VersionText
  }
  catch {
    return $null
  }
}

function Get-AllInstalledPrograms {
  $items = @()

  $targets = @(
    @{ Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' },
    @{ Hive='LocalMachine'; View='Registry32'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' },
    @{ Hive='CurrentUser';  View='Default';    Path='Software\Microsoft\Windows\CurrentVersion\Uninstall' }
  )

  foreach ($t in $targets) {
    try {
      $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::$($t.Hive),
        [Microsoft.Win32.RegistryView]::$($t.View)
      )

      $key = $base.OpenSubKey($t.Path)
      if (-not $key) { continue }

      foreach ($subName in $key.GetSubKeyNames()) {
        try {
          $sub = $key.OpenSubKey($subName)
          $displayName = $sub.GetValue('DisplayName')
          if (-not $displayName) { continue }

          $items += [pscustomobject]@{
            DisplayName    = "$displayName"
            DisplayVersion = "$($sub.GetValue('DisplayVersion'))"
            Publisher      = "$($sub.GetValue('Publisher'))"
          }
        }
        catch {}
      }
    }
    catch {}
  }

  return $items
}

function Get-InstalledVersionForCsvApp {
  param(
    [string]$CsvName,
    [string]$CsvPublisher,
    [array]$InstalledPrograms
  )

  # Special handling for Office Click-to-Run language entries.
  if ($CsvName -match 'Microsoft 365 Apps for enterprise|Aplicaciones de Microsoft 365') {
    try {
      $office = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop
      if ($office.VersionToReport) {
        return "$($office.VersionToReport)"
      }
    }
    catch {}
  }

  # First try exact display name.
  $match = $InstalledPrograms |
    Where-Object { $_.DisplayName -ieq $CsvName } |
    Select-Object -First 1

  if ($match -and $match.DisplayVersion) {
    return $match.DisplayVersion
  }

  # Remove trailing version text / architecture text from the CSV name
  # to make matching resilient to version changes.
  $baseName = $CsvName

  $baseName = $baseName -replace '\s+\d+(\.\d+)+.*$',''
  $baseName = $baseName -replace '\s+\(64-bit\)$',''
  $baseName = $baseName -replace '\s+\(x64\)$',''
  $baseName = $baseName -replace '\s+\(64-bit x64\)$',''
  $baseName = $baseName.Trim()

  $candidates = @(
    $InstalledPrograms | Where-Object {
      $_.DisplayName -ilike "$baseName*" -or
      $_.DisplayName -ilike "*$baseName*"
    }
  )

  # Publisher helps avoid false matches where possible.
  if ($CsvPublisher -and $CsvPublisher -ne 'nan') {
    $publisherCandidates = @(
      $candidates | Where-Object {
        $_.Publisher -and (
          $_.Publisher -ilike "*$CsvPublisher*" -or
          $CsvPublisher -ilike "*$($_.Publisher)*"
        )
      }
    )

    if ($publisherCandidates.Count -gt 0) {
      $candidates = $publisherCandidates
    }
  }

  $best = $candidates |
    Where-Object { $_.DisplayVersion } |
    Select-Object -First 1

  if ($best) {
    return $best.DisplayVersion
  }

  # Appx package fallback for modern apps.
  try {
    $appx = Get-AppxPackage -AllUsers |
      Where-Object {
        $_.Name -ilike "*$baseName*" -or
        $_.PackageFullName -ilike "*$baseName*"
      } |
      Sort-Object Version -Descending |
      Select-Object -First 1

    if ($appx) {
      return "$($appx.Version)"
    }
  }
  catch {}

  return '-'
}

# ---------------------- Run inventory ----------------------
$csvApps = Import-Csv -LiteralPath $CsvPath
$installedPrograms = @(Get-AllInstalledPrograms)
$evergreenReady = Test-EvergreenReady

if (-not $evergreenReady) {
  Write-Warning "Evergreen module is not installed or could not be loaded. Latest Evergreen versions will show as '-'."
}

$reportRows = foreach ($app in $csvApps) {
  $name = "$($app.Name)".Trim()
  $publisher = "$($app.Publisher)".Trim()

  $installedVersion = Get-InstalledVersionForCsvApp `
    -CsvName $name `
    -CsvPublisher $publisher `
    -InstalledPrograms $installedPrograms

  $evergreenName = $null
  if ($EvergreenMap.ContainsKey($name)) {
    $evergreenName = $EvergreenMap[$name]
  }

  $latestEvergreen = $null
  if ($evergreenReady -and $evergreenName) {
    $latestEvergreen = Get-LatestEvergreenVersion -EvergreenName $evergreenName
  }

  [pscustomobject]@{
    Application       = $name
    InstalledVersion  = if ($installedVersion) { $installedVersion } else { '-' }
    LatestEvergreen   = if ($latestEvergreen) { $latestEvergreen } else { '-' }
  }
}

# ---------------------- Output ----------------------
$baseName = "$machine AVD Application Inventory $todayIso"
$outCsv   = Join-Path $monthFolder ($baseName + '.csv')
$outHtml  = Join-Path $monthFolder ($baseName + '.html')

$reportRows |
  Sort-Object Application |
  Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outCsv

$css = @"
<style>
body {
  font-family: Segoe UI, Arial, sans-serif;
  margin: 24px;
  color: #222;
}
h1 {
  margin-bottom: 4px;
}
.subtitle {
  color: #666;
  margin-bottom: 18px;
}
table {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
}
th {
  background: #f2f4f7;
  border: 1px solid #d9dde3;
  padding: 9px;
  text-align: left;
}
td {
  border: 1px solid #d9dde3;
  padding: 9px;
}
tr:nth-child(even) {
  background: #fafafa;
}
.missing {
  color: #b42318;
  font-weight: 600;
}
.evergreen {
  color: #175cd3;
  font-weight: 600;
}
.footer {
  margin-top: 16px;
  color: #777;
  font-size: 12px;
}
</style>
"@

$htmlRows = foreach ($row in ($reportRows | Sort-Object Application)) {
  $installedClass = if ($row.InstalledVersion -eq '-') { 'missing' } else { '' }
  $evergreenClass = if ($row.LatestEvergreen -ne '-') { 'evergreen' } else { '' }

  "<tr><td>$($row.Application)</td><td class='$installedClass'>$($row.InstalledVersion)</td><td class='$evergreenClass'>$($row.LatestEvergreen)</td></tr>"
}

$html = @"
<html>
<head>
<meta charset='utf-8'>
<title>$baseName</title>
$css
</head>
<body>

<h1>AVD Application Inventory</h1>
<div class='subtitle'>
Machine: <strong>$machine</strong> &nbsp; | &nbsp;
Generated: $(Get-Date -Format 'dd/MM/yyyy HH:mm')
</div>

<table>
<thead>
<tr>
<th>Application</th>
<th>Installed Version</th>
<th>Latest Evergreen Version</th>
</tr>
</thead>
<tbody>
$(($htmlRows -join "`r`n"))
</tbody>
</table>

<div class='footer'>
Latest Evergreen Version is populated only for applications mapped to a supported Evergreen application.
</div>

</body>
</html>
"@

$html | Set-Content -LiteralPath $outHtml -Encoding UTF8

Write-Host ""
Write-Host "AVD application inventory completed." -ForegroundColor Green
Write-Host "CSV : $outCsv" -ForegroundColor Cyan
Write-Host "HTML: $outHtml" -ForegroundColor Cyan
