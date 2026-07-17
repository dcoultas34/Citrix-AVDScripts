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
  Installed After, Latest and Status fields.
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

function Convert-ToStandardAppRow {
  param(
    [Parameter(Mandatory)] $Row,
    [Parameter(Mandatory)] [string]$FallbackComputer
  )

  [pscustomobject]@{
    Computer   = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Computer','ComputerName','Server','ServerName','Machine') -DefaultValue $FallbackComputer
    Application = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Application','Name','DisplayName','App') -DefaultValue 'Unknown application'
    InstalledBefore = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('InstalledBefore','Installed Before','Before','BeforeVersion','PreviousVersion','OriginalVersion') -DefaultValue '-'
    InstalledAfter  = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('InstalledAfter','Installed After','After','AfterVersion','Installed','InstalledVersion','CurrentVersion','Version') -DefaultValue '-'
    Latest          = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Latest','LatestVersion','AvailableVersion','TargetVersion') -DefaultValue '-'
    Status     = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Status','Result','State') -DefaultValue 'Unknown'
    Css        = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('Css','CSS','RowCss') -DefaultValue ''
    LatestCss  = Get-FirstPropertyValue -InputObject $Row -PropertyNames @('LatestCss','LatestCSS') -DefaultValue ''
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
    [hashtable]$OsInfoMap,      # NEW: machine -> OS info
    [string]$Title,
    [string]$OutPath,
    [array]$Map,
    [array]$CatalogMap,
    [string]$Period
  )

  $css = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; }
h1 { margin: 0 0 16px 0; }
h3 { margin: 6px 0 12px 0; color:#374151; }
.panel { border:1px solid #e0e0e0; border-radius:10px; padding:16px; margin:18px 0; }
.tag  { display:inline-block; background:#e5e7eb; color:#111827; padding:2px 10px; border-radius:999px; font-size:12px; margin-left:8px; }
.small { color:#6b7280; font-size:12px; margin-top: 16px; }
table { border-collapse: collapse; width: 100%; margin-bottom: 18px; }
th,td { border: 1px solid #ddd; padding: 8px; text-align: left; background:#ffffff; }
th { background: #f3f3f3; }
.ok   { color:#2e7d32; font-weight:600; }
.bad  { color:#c62828; font-weight:600; }
.info { color:#1d4ed8; font-weight:600; }
.os-in-month { color:#2e7d32; font-weight:600; }
.headerline { font-size:18px; font-weight:700; }
ul { margin: 6px 0 12px 18px; }
.catlabel { font-weight:600; }
.warn { color:#c62828; }
</style>
"@

  $repMonth    = [datetime]::ParseExact($Period,'MMMM yyyy',$null)
  $monthStart  = Get-Date -Year $repMonth.Year -Month $repMonth.Month -Day 1 -Hour 0 -Minute 0 -Second 0
  $monthEnd    = $monthStart.AddMonths(1)

  $computers = @($AppData.Computer + $OsData.Computer | Where-Object { $_ } | Select-Object -Unique | Sort-Object)

  $allMasters = @($Map | Select-Object -ExpandProperty Master -Unique | Sort-Object)
  $presentMasters = @($computers | Where-Object { $allMasters -contains $_ } | Sort-Object)
  $missingMasters = @($allMasters | Where-Object { $presentMasters -notcontains $_ } | Sort-Object)
  $countLine = "Master image count: {0} of {1}" -f ($presentMasters.Count), ($allMasters.Count)
  $missingLine = if ($missingMasters.Count -gt 0) { "Please run the report script on: " + ($missingMasters -join ', ') } else { "" }

  $headerBlock = @"
<div class='panel'>
  <div class='headerline'>$countLine</div>
  $(if($missingLine){ "<div class='warn'>$missingLine</div>" } else { "" })
</div>
"@

  $sections = foreach ($comp in $computers) {
    $apps = $AppData | Where-Object { $_.Computer -eq $comp -and $_.Status -ne 'Application not installed' } | Sort-Object Application

    $osMonth = @($OsData | Where-Object {
      $_.Computer -eq $comp -and $_.InstalledOn -ge $monthStart -and $_.InstalledOn -lt $monthEnd
    } | Sort-Object InstalledOn -Descending)

    $osPrev3 = @($OsData | Where-Object {
      $_.Computer -eq $comp -and $_.InstalledOn -lt $monthStart
    } | Sort-Object InstalledOn -Descending | Select-Object -First 3)

    $rows = $Map | Where-Object { $_.Master -ieq $comp }
    $envs = ($rows.Environment | ForEach-Object { "$_".Trim() } | Select-Object -Unique) -join ', '
    $ctxs = ($rows.Citrix      | ForEach-Object { "$_".Trim() } | Select-Object -Unique) -join ', '
    if (-not $envs) { $envs = "-" }
    if (-not $ctxs) { $ctxs = "-" }

    # NEW: OS tag
    $osTag = ''
    if ($OsInfoMap.ContainsKey($comp)) {
      $oi = $OsInfoMap[$comp]
      $osTag = "OS: $($oi.OSName) $($oi.OSVersion) (Build $($oi.OSBuild))"
    }

    $cats = foreach ($c in ($CatalogMap | Where-Object { $_.Master -ieq $comp })) {
      $t = ($c.Type -as [string]).Trim()
      if     ($t -match '^(multi\s*session)$')                { $t = 'Multi Session' }
      elseif ($t -match '^(persistent|persistant|single.*)$') { $t = 'Persistent' }
      [pscustomobject]@{ Name=($c.Name -as [string]).Trim(); Type=$t }
    }
    $catTotal = ($cats | Measure-Object).Count
    $ms = $cats | Where-Object Type -eq 'Multi Session' | Select-Object -ExpandProperty Name
    $ps = $cats | Where-Object Type -eq 'Persistent'    | Select-Object -ExpandProperty Name
    $catHtml = @()
    if ($ms) { $catHtml += "<div class='catlabel'>Multi Session</div><ul>$(( $ms | ForEach-Object { "<li>$_</li>" } ) -join '')</ul>" }
    if ($ps) { $catHtml += "<div class='catlabel'>Persistent</div><ul>$(( $ps | ForEach-Object { "<li>$_</li>" } ) -join '')</ul>" }
    if (-not $catHtml) { $catHtml = @('<em>No catalogues mapped.</em>') }

    $fmt = 'dd/MM/yyyy'
    $osRowsMonth = if ($osMonth -and $osMonth.Count -gt 0) {
      foreach ($o in $osMonth) {
        $d = if ($o.InstalledOn) { $o.InstalledOn.ToString($fmt) } else { '-' }
        "<tr><td class='os-in-month'>$($o.KB)</td><td class='os-in-month'>$($o.Description)</td><td class='os-in-month'>$d</td></tr>"
      }
    } else { @("<tr><td colspan='3'><em>No OS updates were installed in $Period.</em></td></tr>") }

    $osRowsPrev3 = if ($osPrev3 -and $osPrev3.Count -gt 0) {
      foreach ($o in $osPrev3) {
        $d = if ($o.InstalledOn) { $o.InstalledOn.ToString($fmt) } else { '-' }
        "<tr><td>$($o.KB)</td><td>$($o.Description)</td><td>$d</td></tr>"
      }
    } else { @("<tr><td colspan='3'><em>No updates prior to this month were found.</em></td></tr>") }

    $appRows = foreach ($a in $apps) {
      # Css and LatestCss were present in the original collector output, but some
      # existing CSV files do not contain those columns. Read them safely and,
      # when absent, derive the formatting from Status instead.
      $rowCls = ''
      if ($a.PSObject.Properties['Css']) {
        $rowCls = [string]$a.Css
      }
      elseif ([string]$a.Status -match 'Update available|Evergreen not available') {
        $rowCls = 'bad'
      }
      elseif ([string]$a.Status -match 'Latest.*installed|Ahead of') {
        $rowCls = 'ok'
      }

      $latestCls = $rowCls
      if ($a.PSObject.Properties['LatestCss'] -and $a.LatestCss) {
        $latestCls = [string]$a.LatestCss
      }
      elseif ([string]$a.Latest -match 'waiting on Evergreen release') {
        $latestCls = 'info'
      }

      "<tr><td class='$rowCls'>$($a.Application)</td><td class='$rowCls'>$($a.InstalledBefore)</td><td class='$rowCls'>$($a.InstalledAfter)</td><td class='$latestCls'>$($a.Latest)</td><td class='$rowCls'>$($a.Status)</td></tr>"
    }

@"
<div class='panel'>
  <div class='headerline'>$comp
    <span class='tag'>Env: $envs</span>
    <span class='tag'>Citrix: $ctxs</span>
    <span class='tag'>Catalogues: $catTotal</span>
    $(if($osTag){ "<span class='tag'>$osTag</span>" } )
  </div>

  <div><span class='catlabel'>Citrix Machine catalogues:</span><br/>
    $(($catHtml -join ""))  
  </div>

  <h3>Applications</h3>
  <table>
    <thead><tr><th>Application</th><th>Installed Before</th><th>Installed After</th><th>Latest</th><th>Status</th></tr></thead>
    <tbody>
      $(($appRows -join "`n"))
    </tbody>
  </table>

  <h3>OS updates applied in $Period</h3>
  <table>
    <thead><tr><th>KB</th><th>Description</th><th>Installed On</th></tr></thead>
    <tbody>
      $(($osRowsMonth -join "`n"))
    </tbody>
  </table>

  <h3>OS updates – recent 3 (excluding this month)</h3>
  <table>
    <thead><tr><th>KB</th><th>Description</th><th>Installed On</th></tr></thead>
    <tbody>
      $(($osRowsPrev3 -join "`n"))
    </tbody>
  </table>
</div>
"@
  }

  $html = @"
<html>
<head><meta charset='utf-8'><title>$Title</title>$css</head>
<body>
<h1>$Title</h1>
$headerBlock
$(($sections -join "`n"))
<p class='small'>Apps: green = latest installed; red = update available. If Evergreen Stable exists we use it; otherwise winget. If winget is newer than Evergreen, the <em>Latest</em> cell shows that version in <span class='info'>blue</span> with “waiting on Evergreen release”.</p>
<p class='small'>OS: dates shown as DD/MM/YYYY. We show updates applied <strong>this month</strong> (green) and the <strong>previous 3 updates</strong> excluding this month.</p>
<p class='small'>Generated: $(Get-Date)</p>
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
