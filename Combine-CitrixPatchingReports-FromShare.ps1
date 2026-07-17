[CmdletBinding()]
param(
  [string]$ShareRoot = "\\stprodukstcitrix01.file.core.windows.net\profiles\Citrix Reporting\Image Updates",
  [string]$Period    = (Get-Date).ToString("MMMM yyyy"),
  [switch]$Recurse,
  [switch]$GeneratePdf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

<#
.SYNOPSIS
  Builds the combined Citrix patching report using CSV files already stored on
  the configured network share.

.DESCRIPTION
  This version does not inspect the computer on which it is executed. It does
  not query the registry, installed applications, AppX packages, hotfixes,
  CIM/WMI, Evergreen, winget, or Adobe web feeds.

  It only reads these existing files from the selected network-share folder:
    * <Computer>-AppStatus-<date>.csv
    * <Computer>-OSUpdates-<date>.csv
    * <Computer>-OSInfo-<date>.csv

  When more than one dated file exists for a computer and report type, the
  newest file is used. Application rows are normalised to Installed Before,
  Installed After and Status fields. Application comparison uses the complete
  application name, so similarly named products remain separate.
#>

# ---------------------- Master -> Citrix mapping ----------------------
$MasterMap = @(
  @{ Master='UKST1MICTXMI3U'; Environment='Non prod' ; Citrix='Windows-10-MS-NonProd-Standard' }
  @{ Master='UKST1MICTXMI5' ; Environment='Production'; Citrix='Windows-10-MS-Prod-Standard' }
  @{ Master='UKST1MICTXMI4C' ; Environment='Admin tools'; Citrix='Windows-10-MS-Prod-Access' }
  @{ Master='UKST1MICTXMI1D'; Environment='Dev'        ; Citrix='Windows-10-MS-DEVL' }
  @{ Master='UKST1MICTXMI1U'; Environment='CTE'        ; Citrix='Windows-10-MS-Test' }
) | ForEach-Object { [pscustomobject]$_ }  # coerce to PSCustomObject

# ---------------------- Citrix Machine Catalogues -> Master mapping ----------------------
$CatalogMap = @(
  # ---- Multi Session ----
  @{ Name='Windows 10 Multi Session UK South 3E Azure DR Prod' ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South AuditTrack Prod'  ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South DEVQA Prod'       ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South Ent Apps Prod'    ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South HR Systems Prod'  ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South iManage Prod'     ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South Standard Prod'    ;  Master='UKST1MICTXMI5' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South CTE'              ;  Master='UKST1MICTXMI1U'; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South CTE - UAT'        ;  Master='UKST1MICTXMI1U'; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South 3E Azure DR UAT'  ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South AuditTrack UAT'   ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South DEVQA UAT'        ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South Ent Apps UAT'     ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South HR Systems UAT'   ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South iManage UAT'      ;  Master='UKST1MICTXMI4C' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South DEV'              ;  Master='UKST1MICTXMI1D'; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South DEV - UAT'        ;  Master='UKST1MICTXMI1D'; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South Standard Non-Prod';  Master='UKST1MICTXMI3U'; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South Standard Non-Prod IT UAT'; Master='UKST1MICTXMI3U'; Type='Multi Session' }

) | ForEach-Object { [pscustomobject]$_ }

# ---------------------- Network-share input ----------------------
function Get-ComputerNameFromReportFile {
  param(
    [Parameter(Mandatory)] [System.IO.FileInfo]$File,
    [Parameter(Mandatory)] [ValidateSet('AppStatus','OSUpdates','OSInfo')] [string]$ReportType
  )

  $escapedType = [regex]::Escape($ReportType)
  if ($File.BaseName -match "^(?<Computer>.+?)-$escapedType-\d{4}-\d{2}-\d{2}$") {
    return $Matches.Computer.Trim()
  }
  if ($File.BaseName -match "^(?<Computer>.+?)-$escapedType(?:-.*)?$") {
    return $Matches.Computer.Trim()
  }
  return $null
}

function Get-LatestReportFiles {
  param(
    [Parameter(Mandatory)] [string]$Folder,
    [Parameter(Mandatory)] [ValidateSet('AppStatus','OSUpdates','OSInfo')] [string]$ReportType,
    [switch]$Recurse
  )

  $params = @{
    Path        = $Folder
    File        = $true
    Filter      = "*-$ReportType-*.csv"
    ErrorAction = 'SilentlyContinue'
  }
  if ($Recurse) { $params.Recurse = $true }

  $files = Get-ChildItem @params
  $identified = foreach ($file in $files) {
    $computer = Get-ComputerNameFromReportFile -File $file -ReportType $ReportType
    if ($computer) {
      [pscustomobject]@{
        Computer     = $computer
        File         = $file
        LastWriteTime = $file.LastWriteTime
      }
    }
    else {
      Write-Warning "Ignoring file with an unexpected name: $($file.FullName)"
    }
  }

  $identified |
    Group-Object Computer |
    ForEach-Object {
      $_.Group | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
}

function Convert-ToDate {
  param($Value)

  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace("$Value")) { return $null }
  if ($Value -is [datetime]) { return $Value }

  foreach ($format in @(
    'dd/MM/yyyy HH:mm:ss', 'dd/MM/yyyy',
    'MM/dd/yyyy HH:mm:ss', 'MM/dd/yyyy',
    'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd'
  )) {
    try {
      return [datetime]::ParseExact(
        "$Value",
        $format,
        [System.Globalization.CultureInfo]::InvariantCulture
      )
    }
    catch {}
  }

  try { return [datetime]::Parse("$Value", [System.Globalization.CultureInfo]::GetCultureInfo('en-GB')) }
  catch { return $null }
}


function Get-FirstPropertyValue {
  param(
    [Parameter(Mandatory)] $InputObject,
    [Parameter(Mandatory)] [string[]]$PropertyNames,
    $DefaultValue = $null
  )

  # First try exact property-name matches.
  foreach ($name in $PropertyNames) {
    $property = $InputObject.PSObject.Properties[$name]
    if ($property -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace("$($property.Value)")) {
      return $property.Value
    }
  }

  # Then compare normalised names. This allows headings such as:
  # Installed Version, Installed-Version, installed_version and InstalledVersion.
  $normalisedWanted = @(
    $PropertyNames | ForEach-Object {
      ("$_" -replace '[^a-zA-Z0-9]', '').ToLowerInvariant()
    }
  )

  foreach ($property in $InputObject.PSObject.Properties) {
    $normalisedActual = (($property.Name -replace '[^a-zA-Z0-9]', '').ToLowerInvariant())
    if ($normalisedWanted -contains $normalisedActual) {
      if ($null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace("$($property.Value)")) {
        return $property.Value
      }
    }
  }

  return $DefaultValue
}


function Convert-ToPlainText {
  param($Value)

  if ($null -eq $Value) { return $null }
  $text = [string]$Value

  # Existing CSV data may contain HTML anchors (for example, rewritten
  # Mimecast links). Keep only the visible text for the report.
  $text = [regex]::Replace($text, '<[^>]+>', '')
  $text = [System.Net.WebUtility]::HtmlDecode($text)
  return $text.Trim()
}

function Convert-ToHtmlSafe {
  param($Value)
  if ($null -eq $Value) { return '' }
  return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-ApplicationIdentityKey {
  param(
    [Parameter(Mandatory)]
    [string]$Application
  )

  # Match applications using the complete application name only.
  # We normalise harmless formatting differences such as repeated whitespace
  # and letter case, but never use partial, prefix, contains, or fuzzy matching.
  #
  # Therefore these remain separate:
  #   Adobe Acrobat
  #   Adobe Acrobat Reader
  #   Git for Windows
  #   GitHub Desktop (machine)
  $name = (Convert-ToPlainText $Application)
  if ([string]::IsNullOrWhiteSpace($name)) {
    return 'unknown application'
  }

  $name = [regex]::Replace($name.Trim(), '\s+', ' ')
  return $name.ToLowerInvariant()
}

function Test-AppNeedsAttention {
  param($App)

  $status = [string]$App.Status
  if ($status -match '(?i)update available|failed|failure|error|unknown|not available|could not query|requires attention') {
    return $true
  }

  if ([string]::IsNullOrWhiteSpace([string]$App.InstalledAfter) -or [string]$App.InstalledAfter -eq '-') {
    return $true
  }

  return $false
}

function Get-AppDisplayStatus {
  param($App)

  if (Test-AppNeedsAttention -App $App) { return 'Requires attention' }
  if ([string]$App.InstalledBefore -ne [string]$App.InstalledAfter) { return 'Updated' }
  return 'Already current'
}

function Get-AppStatusReason {
  param($App)

  $sourceStatus = ([string]$App.Status).Trim()
  $before = ([string]$App.InstalledBefore).Trim()
  $after  = ([string]$App.InstalledAfter).Trim()

  if ([string]::IsNullOrWhiteSpace($after) -or $after -eq '-') {
    return 'The post-patching installed version could not be determined.'
  }

  if ($sourceStatus -match '(?i)update available') {
    if ($sourceStatus -match '(?i)latest\s+([0-9][0-9A-Za-z\.\-_]+)') {
      return "An update is available. Latest reported version: $($Matches[1])."
    }
    return 'An update is available, but the newer version was not installed.'
  }

  if ($sourceStatus -match '(?i)could not query|unable to query|feed.*(failed|unavailable)|verification failed') {
    return 'Version compliance could not be verified because the update source could not be queried.'
  }

  if ($sourceStatus -match '(?i)failed|failure|error') {
    return "The source report recorded an update error: $sourceStatus"
  }

  if ($sourceStatus -match '(?i)unknown|not available') {
    return "The source report could not determine compliance: $sourceStatus"
  }

  if ($sourceStatus -match '(?i)requires attention') {
    return $sourceStatus
  }

  if ($before -ne $after) {
    return 'Successfully updated during this patching cycle.'
  }

  return 'No update was required; the installed version remained current.'
}

function Convert-ToStandardAppRow {
  param(
    [Parameter(Mandatory)] $Row,
    [Parameter(Mandatory)] [string]$FallbackComputer
  )

  [pscustomobject]@{
    Computer   = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Computer','ComputerName','Server','ServerName','Machine') -DefaultValue $FallbackComputer
    Application = Convert-ToPlainText (Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Application','Name','DisplayName','App') -DefaultValue 'Unknown application')
    InstalledBefore = Convert-ToPlainText (Get-FirstPropertyValue -InputObject $Row -PropertyNames @('InstalledBefore','Installed Before','Before','BeforeVersion','PreviousVersion','OriginalVersion') -DefaultValue '-')
    InstalledAfter  = Convert-ToPlainText (Get-FirstPropertyValue -InputObject $Row -PropertyNames @('InstalledAfter','Installed After','After','AfterVersion','Installed','InstalledVersion','CurrentVersion','Version') -DefaultValue '-')
    Status     = Convert-ToPlainText (Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Status','Result','State') -DefaultValue 'Unknown')
    Css        = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Css','CSS','RowCss') -DefaultValue ''
  }
}

function Convert-ToStandardOsRow {
  param(
    [Parameter(Mandatory)] $Row,
    [Parameter(Mandatory)] [string]$FallbackComputer
  )

  [pscustomobject]@{
    Computer    = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Computer','ComputerName','Server','ServerName','Machine') -DefaultValue $FallbackComputer
    KB          = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('KB','HotFixID','KBNumber') -DefaultValue '-'
    Description = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Description','UpdateDescription','Title') -DefaultValue '-'
    InstalledOn = Convert-ToDate (Get-FirstPropertyValue -InputObject $Row -PropertyNames @('InstalledOn','InstalledDate','DateInstalled') -DefaultValue $null)
  }
}

function Convert-ToStandardOsInfoRow {
  param(
    [Parameter(Mandatory)] $Row,
    [Parameter(Mandatory)] [string]$FallbackComputer
  )

  [pscustomobject]@{
    Computer  = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Computer','ComputerName','Server','ServerName','Machine') -DefaultValue $FallbackComputer
    OSName    = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('OSName','Caption','OperatingSystem') -DefaultValue '-'
    OSVersion = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('OSVersion','Version') -DefaultValue '-'
    OSBuild   = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('OSBuild','Build','BuildNumber') -DefaultValue '-'
  }
}

# ---------------------- Combined HTML ----------------------
function New-CombinedHtml {
  param(
    [array]$AppData,
    [array]$OsData,
    [hashtable]$OsInfoMap,
    [string]$Title,
    [string]$OutPath,
    [array]$Map,
    [array]$CatalogMap,
    [string]$Period
  )

  $css = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color:#111827; }
h1 { margin: 0 0 16px 0; }
h2 { margin: 8px 0 14px 0; font-size:20px; }
h3 { margin: 14px 0 10px 0; color:#374151; }
.panel { border:1px solid #dfe3e8; border-radius:10px; padding:16px; margin:18px 0; page-break-inside:avoid; }
.tag { display:inline-block; background:#e5e7eb; color:#111827; padding:3px 10px; border-radius:999px; font-size:12px; margin:3px 0 3px 8px; }
.small { color:#6b7280; font-size:12px; margin-top:16px; }
table { border-collapse:collapse; width:100%; margin-bottom:18px; }
th,td { border:1px solid #d7dce1; padding:8px; text-align:left; vertical-align:top; }
th { background:#f3f4f6; }
.headerline { font-size:18px; font-weight:700; }
ul { margin:6px 0 12px 18px; }
.catlabel { font-weight:600; }
.warn { color:#b91c1c; font-weight:600; }
.good { color:#166534; font-weight:700; }
.amber { color:#92400e; font-weight:700; }
.bad { color:#b91c1c; font-weight:700; }
.row-updated td { background:#ecfdf3; }
.row-current td { background:#ffffff; }
.row-attention td { background:#fef2f2; color:#991b1b; font-weight:600; }
.os-in-month { color:#166534; font-weight:600; }
.summary-grid { display:flex; flex-wrap:wrap; gap:12px; margin-top:12px; }
.summary-card { min-width:180px; border:1px solid #dfe3e8; border-radius:9px; padding:12px 14px; background:#fafafa; }
.summary-number { font-size:26px; font-weight:750; line-height:1.1; }
.summary-label { color:#4b5563; font-size:13px; margin-top:4px; }
.status-pill { display:inline-block; border-radius:999px; padding:3px 9px; font-size:12px; font-weight:700; }
.status-green { background:#dcfce7; color:#166534; }
.status-amber { background:#fef3c7; color:#92400e; }
.status-red { background:#fee2e2; color:#991b1b; }
.note { color:#4b5563; font-size:12px; }
</style>
"@

  $repMonth   = [datetime]::ParseExact($Period,'MMMM yyyy',[System.Globalization.CultureInfo]::InvariantCulture)
  $monthStart = Get-Date -Year $repMonth.Year -Month $repMonth.Month -Day 1 -Hour 0 -Minute 0 -Second 0
  $monthEnd   = $monthStart.AddMonths(1)

  $computers = @($AppData.Computer + $OsData.Computer | Where-Object { $_ } | Select-Object -Unique | Sort-Object)
  $allMasters = @($Map | Select-Object -ExpandProperty Master -Unique | Sort-Object)
  $presentMasters = @($computers | Where-Object { $allMasters -contains $_ } | Sort-Object)
  $missingMasters = @($allMasters | Where-Object { $presentMasters -notcontains $_ } | Sort-Object)

  $reportApps = @($AppData | Where-Object { $_.Status -ne 'Application not installed' })
  $updatedCount = @($reportApps | Where-Object { -not (Test-AppNeedsAttention -App $_) -and [string]$_.InstalledBefore -ne [string]$_.InstalledAfter }).Count
  $currentCount = @($reportApps | Where-Object { -not (Test-AppNeedsAttention -App $_) -and [string]$_.InstalledBefore -eq [string]$_.InstalledAfter }).Count
  $attentionCount = @($reportApps | Where-Object { Test-AppNeedsAttention -App $_ }).Count
  $monthlyOsCount = @($OsData | Where-Object { $_.InstalledOn -ge $monthStart -and $_.InstalledOn -lt $monthEnd }).Count

  $countLine = "Master image count: {0} of {1}" -f $presentMasters.Count, $allMasters.Count
  $missingLine = if ($missingMasters.Count -gt 0) { "Missing reports for: " + ($missingMasters -join ', ') } else { '' }

  $matrixRows = foreach ($comp in $computers) {
    $machineApps = @($reportApps | Where-Object Computer -eq $comp)
    $machineOsMonth = @($OsData | Where-Object { $_.Computer -eq $comp -and $_.InstalledOn -ge $monthStart -and $_.InstalledOn -lt $monthEnd })
    $needsAttention = @($machineApps | Where-Object { Test-AppNeedsAttention -App $_ }).Count -gt 0

    $appClass = if ($needsAttention) { 'status-red' } else { 'status-green' }
    $appText  = if ($needsAttention) { 'Attention' } else { 'Pass' }
    $osClass  = if ($machineOsMonth.Count -gt 0) { 'status-green' } else { 'status-amber' }
    $osText   = if ($machineOsMonth.Count -gt 0) { 'Pass' } else { 'No monthly updates' }
    $overallClass = if ($needsAttention) { 'status-red' } elseif ($machineOsMonth.Count -eq 0) { 'status-amber' } else { 'status-green' }
    $overallText  = if ($needsAttention) { 'Attention' } elseif ($machineOsMonth.Count -eq 0) { 'Review' } else { 'Pass' }

    "<tr><td>$(Convert-ToHtmlSafe $comp)</td><td><span class='status-pill $appClass'>$appText</span></td><td><span class='status-pill $osClass'>$osText</span></td><td><span class='status-pill $overallClass'>$overallText</span></td></tr>"
  }

  # Only flag genuine inconsistencies. An application being absent from an
  # image is intentionally ignored because each image can have a different role.
  $applicationGroups = @(
    $reportApps |
      Group-Object -Property { Get-ApplicationIdentityKey -Application ([string]$_.Application) } |
      Sort-Object {
        $firstRow = @($_.Group)[0]
        [string]$firstRow.Application
      }
  )

  $consistencyRows = foreach ($group in $applicationGroups) {
    $rows = @($group.Group)
    $displayApplication = [string]($rows | Select-Object -First 1 -ExpandProperty Application)
    $versions = @($rows.InstalledAfter | Where-Object { $_ -and $_ -ne '-' } | Select-Object -Unique | Sort-Object)
    $attentionMachines = @($rows | Where-Object { Test-AppNeedsAttention -App $_ } | Select-Object -ExpandProperty Computer -Unique | Sort-Object)

    $notes = @()

    if ($versions.Count -gt 1) {
      $versionDetails = @(
        $rows |
          Where-Object { $_.InstalledAfter -and $_.InstalledAfter -ne '-' } |
          Sort-Object Computer |
          ForEach-Object { "$($_.Computer): $($_.InstalledAfter)" }
      )
      $notes += ('Version mismatch after patching — ' + ($versionDetails -join '; '))
    }

    foreach ($attentionRow in ($rows | Where-Object { Test-AppNeedsAttention -App $_ } | Sort-Object Computer)) {
      $reason = Get-AppStatusReason -App $attentionRow
      $notes += ("$($attentionRow.Computer): $reason")
    }

    # Do not add a row when the application is consistent. This keeps the
    # review section focused only on items that genuinely need investigation.
    if ($notes.Count -gt 0) {
      $noteClass = if ($attentionMachines.Count -gt 0) { 'bad' } else { 'amber' }
      $noteHtml = ($notes | ForEach-Object { '<div>' + (Convert-ToHtmlSafe $_) + '</div>' }) -join ''
      "<tr><td>$(Convert-ToHtmlSafe $displayApplication)</td><td>$(Convert-ToHtmlSafe ($versions -join ', '))</td><td class='$noteClass'>$noteHtml</td></tr>"
    }
  }
  $consistencyRows = @($consistencyRows | Where-Object { $_ })

  $consistencyBody = if ($consistencyRows.Count -gt 0) {
    ($consistencyRows -join "`n")
  }
  else {
    "<tr><td colspan='3' class='good'>No application inconsistencies detected.</td></tr>"
  }

  $headerBlock = @"
<div class='panel'>
  <div class='headerline'>$countLine</div>
  $(if($missingLine){ "<div class='warn'>$missingLine</div>" } else { '' })
  <div class='summary-grid'>
    <div class='summary-card'><div class='summary-number'>$updatedCount</div><div class='summary-label'>Applications updated</div></div>
    <div class='summary-card'><div class='summary-number'>$currentCount</div><div class='summary-label'>Already current</div></div>
    <div class='summary-card'><div class='summary-number'>$attentionCount</div><div class='summary-label'>Require attention</div></div>
    <div class='summary-card'><div class='summary-number'>$monthlyOsCount</div><div class='summary-label'>OS updates recorded this month</div></div>
  </div>
</div>

<div class='panel'>
  <h2>Compliance overview</h2>
  <table>
    <thead><tr><th>Master image</th><th>Applications</th><th>OS updates</th><th>Overall</th></tr></thead>
    <tbody>$(($matrixRows -join "`n"))</tbody>
  </table>
</div>

<div class='panel'>
  <h2>Items requiring attention</h2>
  <p class='note'>Applications that are intentionally absent from an image are not included. Only version differences and application results requiring investigation are shown.</p>
  <table>
    <thead><tr><th>Application</th><th>Installed After version(s)</th><th>Review note</th></tr></thead>
    <tbody>$consistencyBody</tbody>
  </table>
</div>
"@

  $sections = foreach ($comp in $computers) {
    $apps = @($reportApps | Where-Object Computer -eq $comp | Sort-Object Application)
    $osMonth = @($OsData | Where-Object { $_.Computer -eq $comp -and $_.InstalledOn -ge $monthStart -and $_.InstalledOn -lt $monthEnd } | Sort-Object InstalledOn -Descending)
    $osPrev3 = @($OsData | Where-Object { $_.Computer -eq $comp -and $_.InstalledOn -lt $monthStart } | Sort-Object InstalledOn -Descending | Select-Object -First 3)

    $mapRows = @($Map | Where-Object { $_.Master -ieq $comp })
    $envs = ($mapRows.Environment | ForEach-Object { "$($_)".Trim() } | Select-Object -Unique) -join ', '
    $ctxs = ($mapRows.Citrix | ForEach-Object { "$($_)".Trim() } | Select-Object -Unique) -join ', '
    if (-not $envs) { $envs = '-' }
    if (-not $ctxs) { $ctxs = '-' }

    $osTag = ''
    if ($OsInfoMap.ContainsKey($comp)) {
      $oi = $OsInfoMap[$comp]
      $osTag = "OS: $($oi.OSName) $($oi.OSVersion) (Build $($oi.OSBuild))"
    }

    $cats = @(
      foreach ($c in ($CatalogMap | Where-Object { $_.Master -ieq $comp })) {
        $type = ([string]$c.Type).Trim()
        if ($type -match '^(multi\s*session)$') { $type = 'Multi Session' }
        elseif ($type -match '^(persistent|persistant|single.*)$') { $type = 'Persistent' }
        [pscustomobject]@{ Name=([string]$c.Name).Trim(); Type=$type }
      }
    )
    $catTotal = $cats.Count
    $ms = @($cats | Where-Object Type -eq 'Multi Session' | Select-Object -ExpandProperty Name)
    $ps = @($cats | Where-Object Type -eq 'Persistent' | Select-Object -ExpandProperty Name)
    $catHtml = @()
    if ($ms.Count -gt 0) { $catHtml += "<div class='catlabel'>Multi Session</div><ul>$(( $ms | ForEach-Object { '<li>' + (Convert-ToHtmlSafe $_) + '</li>' } ) -join '')</ul>" }
    if ($ps.Count -gt 0) { $catHtml += "<div class='catlabel'>Persistent</div><ul>$(( $ps | ForEach-Object { '<li>' + (Convert-ToHtmlSafe $_) + '</li>' } ) -join '')</ul>" }
    if ($catHtml.Count -eq 0) { $catHtml = @('<em>No catalogues mapped.</em>') }

    $appRows = foreach ($a in $apps) {
      $attention = Test-AppNeedsAttention -App $a
      $changed = [string]$a.InstalledBefore -ne [string]$a.InstalledAfter
      $rowClass = if ($attention) { 'row-attention' } elseif ($changed) { 'row-updated' } else { 'row-current' }
      $displayStatus = Get-AppDisplayStatus -App $a
      $statusReason = Get-AppStatusReason -App $a
      $statusClass = if ($attention) { 'status-red' } elseif ($changed) { 'status-green' } else { 'status-green' }

      "<tr class='$rowClass'><td>$(Convert-ToHtmlSafe $a.Application)</td><td>$(Convert-ToHtmlSafe $a.InstalledBefore)</td><td>$(Convert-ToHtmlSafe $a.InstalledAfter)</td><td><span class='status-pill $statusClass'>$displayStatus</span></td><td>$(Convert-ToHtmlSafe $statusReason)</td></tr>"
    }

    $fmt = 'dd/MM/yyyy'
    $osRowsMonth = if ($osMonth.Count -gt 0) {
      foreach ($o in $osMonth) {
        $date = if ($o.InstalledOn) { $o.InstalledOn.ToString($fmt) } else { '-' }
        "<tr><td class='os-in-month'>$(Convert-ToHtmlSafe $o.KB)</td><td class='os-in-month'>$(Convert-ToHtmlSafe $o.Description)</td><td class='os-in-month'>$date</td></tr>"
      }
    } else { @("<tr><td colspan='3'><em>No OS updates were installed in $(Convert-ToHtmlSafe $Period).</em></td></tr>") }

    $osRowsPrev3 = if ($osPrev3.Count -gt 0) {
      foreach ($o in $osPrev3) {
        $date = if ($o.InstalledOn) { $o.InstalledOn.ToString($fmt) } else { '-' }
        "<tr><td>$(Convert-ToHtmlSafe $o.KB)</td><td>$(Convert-ToHtmlSafe $o.Description)</td><td>$date</td></tr>"
      }
    } else { @("<tr><td colspan='3'><em>No updates prior to this month were found.</em></td></tr>") }

@"
<div class='panel'>
  <div class='headerline'>$(Convert-ToHtmlSafe $comp)
    <span class='tag'>Env: $(Convert-ToHtmlSafe $envs)</span>
    <span class='tag'>Citrix: $(Convert-ToHtmlSafe $ctxs)</span>
    <span class='tag'>Catalogues: $catTotal</span>
    $(if($osTag){ "<span class='tag'>$(Convert-ToHtmlSafe $osTag)</span>" } else { '' })
  </div>

  <div><span class='catlabel'>Citrix Machine catalogues:</span><br/>$(($catHtml -join ''))</div>

  <h3>Applications</h3>
  <table>
    <thead><tr><th>Application</th><th>Installed Before</th><th>Installed After</th><th>Status</th><th>Reason</th></tr></thead>
    <tbody>$(($appRows -join "`n"))</tbody>
  </table>

  <h3>OS updates applied in $(Convert-ToHtmlSafe $Period)</h3>
  <table>
    <thead><tr><th>KB</th><th>Description</th><th>Installed On</th></tr></thead>
    <tbody>$(($osRowsMonth -join "`n"))</tbody>
  </table>

  <h3>OS updates – recent 3 (excluding this month)</h3>
  <table>
    <thead><tr><th>KB</th><th>Description</th><th>Installed On</th></tr></thead>
    <tbody>$(($osRowsPrev3 -join "`n"))</tbody>
  </table>
</div>
"@
  }

  $generated = (Get-Date).ToString('dd/MM/yyyy HH:mm')
  $html = @"
<html>
<head><meta charset='utf-8'><title>$(Convert-ToHtmlSafe $Title)</title>$css</head>
<body>
<h1>$(Convert-ToHtmlSafe $Title)</h1>
$headerBlock
$(($sections -join "`n"))
<p class='small'>Application rows: light green = updated, white = already current, light red = requires attention. The Reason column explains the result, including why an item requires attention.</p>
<p class='small'>The attention review ignores applications that are intentionally absent from particular images. It only flags differing post-patch versions and results that require investigation.</p>
<p class='small'>OS dates use DD/MM/YYYY. The report shows updates applied during the selected month and the previous three updates.</p>
<p class='small'>Generated: $generated</p>
</body>
</html>
"@

  $html | Set-Content -Path $OutPath -Encoding UTF8
  return $OutPath
}

# ---------------------- PDF (optional) ----------------------
function Convert-HtmlToPdf {
  param([string]$HtmlPath,[string]$PdfPath)
  $edge = @("$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe","${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
  if ($edge) {
    foreach ($flags in @(
      @("--headless","--disable-gpu","--print-to-pdf=`"$PdfPath`"","`"$HtmlPath`""),
      @("--headless=new","--disable-gpu","--print-to-pdf=`"$PdfPath`"","`"$HtmlPath`"")
    )) { Start-Process -FilePath $edge -ArgumentList $flags -PassThru -Wait -NoNewWindow | Out-Null; if (Test-Path $PdfPath) { return $true } }
  }
  $chrome = @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe","${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
  if ($chrome) { Start-Process -FilePath $chrome -ArgumentList @("--headless=new","--disable-gpu","--print-to-pdf=`"$PdfPath`"","`"$HtmlPath`"") -PassThru -Wait -NoNewWindow | Out-Null; if (Test-Path $PdfPath) { return $true } }
  $wk = @("C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe","C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
  if ($wk) { Start-Process -FilePath $wk -ArgumentList "`"$HtmlPath`" `"$PdfPath`"" -PassThru -Wait -NoNewWindow | Out-Null; if (Test-Path $PdfPath) { return $true } }
  Write-Warning "PDF step skipped (Edge/Chrome/wkhtmltopdf not found)."; return $false
}

# ---------------------- Robust CSV import ----------------------
function Import-CsvRobust {
  param([string]$Path,[int]$Retries=5,[int]$DelayMs=500)
  for ($i=1; $i -le $Retries; $i++) {
    try { return (Import-Csv -Path $Path -ErrorAction Stop) }
    catch { if ($i -eq $Retries){ Write-Warning "Failed to read '$Path': $($_.Exception.Message)"; return @() } ; Start-Sleep -Milliseconds $DelayMs }
  }
}

# ---------------------- Build combined report ----------------------
try {
  $reportFolder = Join-Path $ShareRoot $Period

  if (-not (Test-Path -LiteralPath $ShareRoot -PathType Container)) {
    throw "The network share cannot be reached: $ShareRoot"
  }
  if (-not (Test-Path -LiteralPath $reportFolder -PathType Container)) {
    throw "The reporting period folder does not exist: $reportFolder"
  }

  Write-Host "Reading report files only from: $reportFolder" -ForegroundColor Cyan

  $appFileRecords = @(Get-LatestReportFiles -Folder $reportFolder -ReportType AppStatus -Recurse:$Recurse)
  if ($appFileRecords.Count -eq 0) {
    throw "No AppStatus CSV files were found in '$reportFolder'."
  }

  $allApps = foreach ($record in $appFileRecords) {
    foreach ($row in @(Import-CsvRobust -Path $record.File.FullName)) {
      if ($row) {
        Convert-ToStandardAppRow -Row $row -FallbackComputer $record.Computer
      }
    }
  }
  $allApps = @($allApps | Where-Object { $_ })
  if ($allApps.Count -eq 0) {
    throw 'AppStatus CSV files were found, but they contained no usable rows.'
  }

  $allOs = foreach ($record in @(Get-LatestReportFiles -Folder $reportFolder -ReportType OSUpdates -Recurse:$Recurse)) {
    foreach ($row in @(Import-CsvRobust -Path $record.File.FullName)) {
      if ($row) {
        Convert-ToStandardOsRow -Row $row -FallbackComputer $record.Computer
      }
    }
  }
  $allOs = @($allOs | Where-Object { $_ })

  $osInfoMap = @{}
  foreach ($record in @(Get-LatestReportFiles -Folder $reportFolder -ReportType OSInfo -Recurse:$Recurse)) {
    foreach ($row in @(Import-CsvRobust -Path $record.File.FullName)) {
      if ($row) {
        $standardRow = Convert-ToStandardOsInfoRow -Row $row -FallbackComputer $record.Computer
        $osInfoMap[$standardRow.Computer] = $standardRow
      }
    }
  }

  $todayIso     = (Get-Date).ToString('yyyy-MM-dd')
  $baseName     = "Citrix Monthly patching report $todayIso"
  $combinedCsv  = Join-Path $reportFolder ($baseName + '.csv')
  $combinedHtml = Join-Path $reportFolder ($baseName + '.html')
  $combinedPdf  = Join-Path $reportFolder ($baseName + '.pdf')

  $allApps |
    Sort-Object Computer, Application |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $combinedCsv
  Write-Host "Combined CSV written:  $combinedCsv" -ForegroundColor Green

  $null = New-CombinedHtml `
    -AppData $allApps `
    -OsData $allOs `
    -OsInfoMap $osInfoMap `
    -Title "Citrix Monthly patching report ($Period)" `
    -OutPath $combinedHtml `
    -Map $MasterMap `
    -CatalogMap $CatalogMap `
    -Period $Period
  Write-Host "Combined HTML written: $combinedHtml" -ForegroundColor Green

  if ($GeneratePdf) {
    if (Convert-HtmlToPdf -HtmlPath $combinedHtml -PdfPath $combinedPdf) {
      Write-Host "Combined PDF written:  $combinedPdf" -ForegroundColor Green
    }
  }

  Write-Host "Completed. No local application, registry, OS, hotfix, CIM/WMI, winget, or Evergreen scan was performed." -ForegroundColor Cyan
}
catch {
  Write-Error $_.Exception.Message
  exit 1
}
