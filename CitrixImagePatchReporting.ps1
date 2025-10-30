#This script runs on each master image server, this uploads to a central repo, this uploads various csv files of patches and OS updates applied. 
#Run the script switch to then combine and create a full HTML Report of work completed to the central repository

 
 [CmdletBinding()]
param(
  [string]$ShareRoot = "\\Transfer\transfer\CitrixApps\Citrix Reporting\Image updates",
  [string]$Period    = (Get-Date).ToString("MMMM yyyy"),
  [switch]$Combine,
  [switch]$NoCombine
)

# ---------------------- Master -> Citrix mapping ----------------------
$MasterMap = @(
  @{ Master='UKST1MICTXMI3U'; Environment='Non prod' ; Citrix='Windows-10-MS-NonProd-Standard' }
  @{ Master='UKST1MICTXMI1' ; Environment='Production'; Citrix='Windows-10-MS-Prod-Standard' }
  @{ Master='UKST1MICTXMI3' ; Environment='Admin tools'; Citrix='Windows-10-Persistent-Prod-IPT' }
  @{ Master='UKST1MICTXMI2' ; Environment='Admin tools'; Citrix='Windows-10-MS-Prod-Access' }
  @{ Master='UKST1MICTXMI2D'; Environment='Dev'        ; Citrix='Windows-10-MS-DEVL' }
  @{ Master='UKST1MICTXMI1D'; Environment='Dev'        ; Citrix='Windows-10-Persistent-DEVL' }
  @{ Master='UKST1MICTXMI1U'; Environment='CTE'        ; Citrix='Windows-10-MS-Test' }
  @{ Master='UKST1MICTXMI2U'; Environment='CTE'        ; Citrix='Windows-10-Persistent-CTE' }
) | ForEach-Object { [pscustomobject]$_ }  # coerce to PSCustomObject

# ---------------------- Citrix Machine Catalogues -> Master mapping ----------------------
$CatalogMap = @(
  # ---- Multi Session ----
  @{ Name='Windows 10 Multi Session UK South 3E Azure DR Prod' ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South AuditTrack Prod'  ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South DEVQA Prod'       ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South Ent Apps Prod'    ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South HR Systems Prod'  ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South iManage Prod'     ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South Standard Prod'    ;  Master='UKST1MICTXMI1' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South CTE'              ;  Master='UKST1MICTXMI1U'; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South CTE - UAT'        ;  Master='UKST1MICTXMI1U'; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South 3E Azure DR UAT'  ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South AuditTrack UAT'   ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South DEVQA UAT'        ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South Ent Apps UAT'     ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South HR Systems UAT'   ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South iManage UAT'      ;  Master='UKST1MICTXMI2' ; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South DEV'              ;  Master='UKST1MICTXMI2D'; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South DEV - UAT'        ;  Master='UKST1MICTXMI2D'; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South Standard Non-Prod';  Master='UKST1MICTXMI3U'; Type='Multi Session' }
  @{ Name='Windows 10 Multi Session UK South Standard Non-Prod IT UAT'; Master='UKST1MICTXMI1U'; Type='Multi Session' }

  # ---- Persistent / Single-User ----
  @{ Name='Windows 10 SU UK South Persistent CTE'                 ; Master='UKST1MICTXMI2U'; Type='Persistent' }
  @{ Name='Windows 10 SU UK South Persistent CTE - UAT'           ; Master='UKST1MICTXMI2U'; Type='Persistent' }
  @{ Name='Windows 10 SU UK South Persistent CTE 3rd Party Access'; Master='UKST1MICTXMI1D'; Type='Persistent' }
  @{ Name='Windows 10 SU UK South Persistent DEV'                 ; Master='UKST1MICTXMI1D'; Type='Persistent' }
  @{ Name='Windows 10 SU UK South Persistent DEV - UAT'           ; Master='UKST1MICTXMI1D'; Type='Persistent' }
  @{ Name='Windows 10 SU UK South Persistent DEV 3rd Party Access'; Master='UKST1MICTXMI1D'; Type='Persistent' }
  @{ Name='Windows 10 Single User UK South IPT UAT'               ; Master='UKST1MICTXMI3' ; Type='Persistent' }
  @{ Name='Windows 10 UK South IPT Prod'                          ; Master='UKST1MICTXMI3' ; Type='Persistent' }
) | ForEach-Object { [pscustomobject]$_ }

# ---------------------- Applications to check ----------------------
$AppsToCheck = @(
  @{ Name="Adobe Acrobat";           IsAdobeFull=$true;  LocalMatch=$null; EvergreenName=$null; WingetId=$null; ExpectEvergreen=$false }
  @{ Name="Microsoft Office 365"; LocalMatch="__OFFICE_C2R__"; EvergreenName="Microsoft365Apps"; PreferredChannel="MonthlyEnterprise"; WingetId="Microsoft.Office"; ExpectEvergreen=$true }
  @{ Name="Adobe Acrobat Reader";    LocalMatch="Adobe Acrobat Reader"; EvergreenName="AdobeAcrobatReaderDC"; WingetId="Adobe.Acrobat.Reader.64-bit"; ExpectEvergreen=$false }
  @{ Name="Google Chrome";           LocalMatch="Google Chrome";        EvergreenName="GoogleChrome";         WingetId="Google.Chrome";               ExpectEvergreen=$false }
  @{ Name="Microsoft Edge";          LocalMatch="Microsoft Edge";       EvergreenName="MicrosoftEdge";        WingetId="Microsoft.Edge";              ExpectEvergreen=$false }
  @{ Name="Visual Studio Code";      LocalMatch="Microsoft Visual Studio Code"; EvergreenName="MicrosoftVisualStudioCode"; WingetId="Microsoft.VisualStudioCode"; ExpectEvergreen=$false }
  @{ Name="Power BI Desktop";        LocalMatch="Microsoft Power BI Desktop";   EvergreenName="MicrosoftPowerBIDesktop";    WingetId="Microsoft.PowerBI";         ExpectEvergreen=$false }
  @{ Name="Microsoft Teams";         LocalMatch="Microsoft Teams";      EvergreenName=$null; WingetId="Microsoft.Teams";  ExpectEvergreen=$false }
  @{ Name="OneDrive";                LocalMatch="Microsoft OneDrive";   EvergreenName=$null; WingetId="Microsoft.OneDrive"; ExpectEvergreen=$false }
  @{ Name="Azure Data Studio";       LocalMatch="Azure Data Studio";    EvergreenName=$null; WingetId="Microsoft.AzureDataStudio"; ExpectEvergreen=$false }
)

# ---------------------- Utilities ----------------------
function Test-WingetReady  { try { $null = Get-Command winget -ErrorAction Stop; winget --version | Out-Null; $true } catch { $false } }
function Test-EvergreenReady { try { $m = Get-Module -ListAvailable -Name Evergreen; if ($m){Import-Module Evergreen -ErrorAction SilentlyContinue|Out-Null;$true}else{$false} } catch { $false } }

function Get-InstalledVersion {
  param([string]$DisplayNameMatch)

  if ($DisplayNameMatch -eq "Microsoft Teams") {
    $pkgs=@(); foreach($n in 'MSTeams','MicrosoftTeams'){ try{$pkgs+=Get-AppxPackage -AllUsers -Name $n -ErrorAction SilentlyContinue}catch{}; try{$pkgs+=Get-AppxPackage -Name $n -ErrorAction SilentlyContinue}catch{} }
    $pkg=$pkgs|Where-Object{$_}|Sort-Object Version -Descending|Select-Object -First 1
    if($pkg){ return $pkg.Version.ToString() }
  }

  if ($DisplayNameMatch -eq "__OFFICE_C2R__") {
    try{
      $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry64)
      $sub=$base.OpenSubKey('SOFTWARE\Microsoft\Office\ClickToRun\Configuration')
      $v=$sub.GetValue('VersionToReport'); if($v){return $v}
    }catch{}; return $null
  }

  if ($DisplayNameMatch -eq "Microsoft Visual Studio Code") {
    foreach($p in @("$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe","$env:ProgramFiles\Microsoft VS Code\Code.exe","${env:ProgramFiles(x86)}\Microsoft VS Code\Code.exe")){
      if(Test-Path $p){ try{ return (Get-Item $p).VersionInfo.ProductVersion }catch{} }
    }
  }

  $entries=@()
  $targets=@(
    @{Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='LocalMachine'; View='Registry32'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='CurrentUser' ; View='Default'   ; Path='Software\Microsoft\Windows\CurrentVersion\Uninstall'}
  )
  foreach($t in $targets){
    try{
      $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::$($t.Hive),[Microsoft.Win32.RegistryView]::$($t.View))
      $key=$base.OpenSubKey($t.Path); if(-not $key){continue}
      foreach($n in $key.GetSubKeyNames()){
        try{ $s=$key.OpenSubKey($n); $dn=$s.GetValue('DisplayName'); if(-not $dn){continue}; $dv=$s.GetValue('DisplayVersion'); $entries+=[pscustomobject]@{DisplayName=$dn;DisplayVersion=$dv} }catch{}
      }
    }catch{}
  }
  $nameExcludes='WebView','Runtime','WebView2','Updater','Update','AutoUpdate','Maintenance','Service','Helper','Crashpad','Stub','Machine-wide','User Installer','System Installer','Setup'
  $cands=@()
  foreach($e in $entries){
    if($e.DisplayName -notlike "*$DisplayNameMatch*"){continue}
    if($nameExcludes | Where-Object { $e.DisplayName -match [regex]::Escape($_) }){continue}
    $dv=$e.DisplayVersion
    $vObj=$null; if($dv){ try{$vObj=[version]($dv -replace '[^\d\.]','')}catch{} }
    $score= if($e.DisplayName -ieq $DisplayNameMatch){3}elseif($e.DisplayName -ilike "$DisplayNameMatch*"){2}else{1}
    $cands += [pscustomobject]@{DisplayName=$e.DisplayName;DisplayVersion=$dv;VersionObj=$vObj;Score=$score}
  }
  if($cands.Count -eq 0){return $null}
  ($cands | Sort-Object @{e='Score';Descending=$true}, @{e='VersionObj';Descending=$true} | Select-Object -First 1).DisplayVersion
}

# -------- Adobe Acrobat (Full/DC/Pro) robust detection --------
function Get-AdobeFullInstalledVersion {
  $paths = @(
    'HKLM:\SOFTWARE\Adobe\Adobe Acrobat',
    'HKLM:\SOFTWARE\WOW6432Node\Adobe\Adobe Acrobat',
    'HKLM:\SOFTWARE\Adobe\Acrobat',
    'HKLM:\SOFTWARE\WOW6432Node\Adobe\Acrobat'
  )
  foreach ($root in $paths) {
    if (Test-Path $root) {
      Get-ChildItem $root -EA SilentlyContinue | ForEach-Object {
        foreach ($leaf in 'Installer','CurrentVersion','\DC\Installer','\DC\CurrentVersion') {
          $k = ($_.PsPath + $leaf)
          if (Test-Path $k) {
            try {
              $p = Get-ItemProperty -Path $k -EA Stop
              foreach($name in 'ACPVersion','ProductVersion','Version','PV') {
                $v = $p.$name
                if ($v -and ($v -match '^\d{1,2}\.\d{3}\.\d{5}$')) { return $v }
              }
            } catch {}
          }
        }
      }
    }
  }
  $targets = @(
    @{Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='LocalMachine'; View='Registry32'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='CurrentUser' ; View='Default'   ; Path='Software\Microsoft\Windows\CurrentVersion\Uninstall'}
  )
  foreach($t in $targets){
    try{
      $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::$($t.Hive),[Microsoft.Win32.RegistryView]::$($t.View))
      $key=$base.OpenSubKey($t.Path); if(-not $key){continue}
      foreach($n in $key.GetSubKeyNames()){
        try{
          $s=$key.OpenSubKey($n)
          $dn=$s.GetValue('DisplayName'); if(-not $dn){continue}
          if($dn -match '^Adobe\s+Acrobat(?!.*Reader)'){
            foreach ($prop in 'DisplayVersion','BundleVersion') {
              $dv=$s.GetValue($prop); if($dv){ return $dv }
            }
          }
        }catch{}
      }
    }catch{}
  }
  $exe = Get-ChildItem "$env:ProgramFiles\Adobe","${env:ProgramFiles(x86)}\Adobe" -Recurse -Filter Acrobat.exe -EA SilentlyContinue | Select-Object -ExpandProperty FullName
  foreach($p in $exe){
    try{
      $fv = (Get-Item $p).VersionInfo.FileVersion
      if ($fv -and ($fv -match '^\d{1,2}\.\d{3}\.\d{5}$')) { return $fv }
    }catch{}
  }
  return $null
}
function Get-LatestFromAdobeAcrobatFull {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
  try {
    $v=(Invoke-RestMethod 'https://armmf.adobe.com/arm-manifests/win/AcrobatDC/acrobat/acrobat/current_version.txt' -UseBasicParsing -TimeoutSec 20).Trim()
    if ($v -match '^\d{1,2}\.\d{3}\.\d{5}$'){return $v}; $null
  } catch { $null }
}

function Get-LatestFromEvergreen {
  param([string]$EvergreenName,[string]$PreferredChannel)
  try {
    $data = Get-EvergreenApp -Name $EvergreenName -ErrorAction Stop
    if ($PreferredChannel -and ($data | Get-Member -Name Channel -EA SilentlyContinue)) {
      $pref = $data | Where-Object { $_.Channel -match $PreferredChannel }
      if ($pref) { $data = $pref }
    }
    if ($data -and ($data | Get-Member -Name Channel -EA SilentlyContinue)) {
      $stable = $data | Where-Object { $_.Channel -match 'Stable' }
      if ($stable) { $data = $stable }
    }
    $top = $data | Where-Object { $_.Version } |
           Sort-Object { try{ [version]($_.Version -replace '[^\d\.]','') }catch{ $_.Version } } -Descending |
           Select-Object -First 1
    if ($top) { [pscustomobject]@{ Version = $top.Version } }
  } catch { $null }
}
function Get-LatestFromWinget {
  param([string]$WingetId)
  try{
    $out = winget show --id $WingetId --exact --source winget --accept-source-agreements 2>$null
    if($out){
      $verLine = ($out -split "`r?`n") | Where-Object { $_ -match "^\s*Version\s*:" } | Select-Object -First 1
      $ver = if($verLine){ ($verLine -split ":\s*",2)[1].Trim() } else { $null }
      [pscustomobject]@{ Version=$ver }
    }
  }catch{}
}
function Compare-VersionSmart { param([string]$Installed,[string]$Latest)
  if(-not $Installed -or -not $Latest){ return $null }
  try{ $v1=[version]($Installed -replace '[^\d\.]',''); $v2=[version]($Latest -replace '[^\d\.]',''); [Math]::Sign($v1.CompareTo($v2)) }
  catch{ if($Installed -eq $Latest){0}else{-1} }
}
function Get-RecentOSUpdates {
  try {
    Get-HotFix | Sort-Object InstalledOn -Descending | ForEach-Object {
      [pscustomobject]@{
        Computer    = $env:COMPUTERNAME
        KB          = $_.HotFixID
        Description = $_.Description
        InstalledOn = $(if ($_.InstalledOn) { [datetime]$_.InstalledOn } else { $null })
      }
    }
  } catch { @() }
}

# ---------------------- Paths & dates ----------------------
$monthFolder = Join-Path $ShareRoot $Period
if (-not (Test-Path $monthFolder)) { New-Item -ItemType Directory -Path $monthFolder -Force | Out-Null }
$today     = Get-Date
$todayIso  = $today.ToString('yyyy-MM-dd')
$todayUK   = $today.ToString('dd/MM/yy')

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

  $computers = ($AppData.Computer + $OsData.Computer | Where-Object { $_ } | Select-Object -Unique | Sort-Object)

  $allMasters = ($Map | Select-Object -ExpandProperty Master -Unique | Sort-Object)
  $presentMasters = $computers | Where-Object { $allMasters -contains $_ } | Sort-Object
  $missingMasters = $allMasters | Where-Object { $presentMasters -notcontains $_ } | Sort-Object
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

    $osMonth = $OsData | Where-Object {
      $_.Computer -eq $comp -and $_.InstalledOn -ge $monthStart -and $_.InstalledOn -lt $monthEnd
    } | Sort-Object InstalledOn -Descending

    $osPrev3 = $OsData | Where-Object {
      $_.Computer -eq $comp -and $_.InstalledOn -lt $monthStart
    } | Sort-Object InstalledOn -Descending | Select-Object -First 3

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
      $rowCls = $a.Css
      $latestCls = if ($a.LatestCss) { $a.LatestCss } else { $rowCls }
      "<tr><td class='$rowCls'>$($a.Application)</td><td class='$rowCls'>$($a.Installed)</td><td class='$latestCls'>$($a.Latest)</td><td class='$rowCls'>$($a.Status)</td></tr>"
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
    <thead><tr><th>Application</th><th>Installed</th><th>Latest</th><th>Status</th></tr></thead>
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

# ---------------------- Channel versions (Evergreen + winget) ----------------------
function Get-ChannelVersions {
  param([string]$EvergreenName,[string]$WingetId,[string]$PreferredChannel)
  $ever = $null; $wing = $null
  if ($EvergreenName -and (Test-EvergreenReady)) { $ever = Get-LatestFromEvergreen -EvergreenName $EvergreenName -PreferredChannel $PreferredChannel }
  if ($WingetId -and (Test-WingetReady))         { $wing = Get-LatestFromWinget    -WingetId       $WingetId }
  [pscustomobject]@{
    EvergreenStable = if($ever  -and $ever.PSObject.Properties['Version']){ $ever.Version } else { $null }
    WingetLatest    = if($wing  -and $wing.PSObject.Properties['Version']){ $wing.Version } else { $null }
  }
}

# ---------------------- Per-machine run ----------------------
$machine  = $env:COMPUTERNAME
$wingetOk = Test-WingetReady
$everOk   = Test-EvergreenReady

Write-Host "Reporting for $machine  |  Period: $Period" -ForegroundColor Cyan
if (-not $everOk)  { Write-Warning "Evergreen module not available. Install with: Install-Module Evergreen -Scope AllUsers" }
if (-not $wingetOk){ Write-Warning "winget not available. Newer-than-Evergreen note may be missing." }

# --- NEW: capture OS info for this machine ---
try {
  $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
  $reg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
  $caption = $os.Caption
  $displayVer = $reg.DisplayVersion
  if ($displayVer) { $caption = "$caption $displayVer" }
  $build = $os.BuildNumber
  $ubr = $reg.UBR
  $buildFull = if ($ubr -ne $null) { "$build.$ubr" } else { $build }
  $osInfo = [pscustomobject]@{
    Computer  = $machine
    OSName    = $caption
    OSVersion = $os.Version
    OSBuild   = $buildFull
  }
} catch { $osInfo = [pscustomobject]@{ Computer=$machine; OSName='-'; OSVersion='-'; OSBuild='-' } }

# Save OS info CSV
$osInfoCsv = Join-Path $monthFolder ("{0} OSInfo {1}.csv" -f $machine,$todayIso)
$osInfo | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $osInfoCsv

$rows = @()

foreach ($app in $AppsToCheck) {

  if ($app.IsAdobeFull) {
    $installed = Get-AdobeFullInstalledVersion
    if (-not $installed) {
      $rows += [pscustomobject]@{ Computer=$machine; Application=$app.Name; Installed='-'; Latest='-'; Status='Application not installed'; Css=''; LatestCss='' }
      continue
    }
    $adobeLatest = Get-LatestFromAdobeAcrobatFull
    if ($adobeLatest) {
      $cmp = Compare-VersionSmart -Installed $installed -Latest $adobeLatest
      if ($cmp -eq 0)       { $status = "Latest version installed as of $todayUK"; $css='ok' }
      elseif ($cmp -lt 0)   { $status = "Update available"; $css='bad' }
      else                  { $status = "Ahead of catalog as of $todayUK"; $css='ok' }
      $rows += [pscustomobject]@{ Computer=$machine; Application=$app.Name; Installed=$installed; Latest=$adobeLatest; Status=$status; Css=$css; LatestCss=$css }
    } else {
      $rows += [pscustomobject]@{ Computer=$machine; Application=$app.Name; Installed=$installed; Latest='-'; Status='Unknown (could not query Adobe feed)'; Css=''; LatestCss='' }
    }
    continue
  }

  $installed = if ($app.LocalMatch) { Get-InstalledVersion -DisplayNameMatch $app.LocalMatch } else { $null }
  if (-not $installed) {
    $rows += [pscustomobject]@{ Computer=$machine; Application=$app.Name; Installed='-'; Latest='-'; Status='Application not installed'; Css=''; LatestCss='' }
    continue
  }

  $chan = Get-ChannelVersions -EvergreenName $app.EvergreenName -WingetId $app.WingetId -PreferredChannel $app.PreferredChannel

  $latestToShow = $null; $status = ''; $css = ''; $latestCss = ''

  if ($chan.EvergreenStable) {
    $ev = $chan.EvergreenStable
    $latestToShow = "$ev (Latest Evergreen stable version)"
    if ($chan.WingetLatest) {
      $cmpWingVsEver = Compare-VersionSmart -Installed $chan.WingetLatest -Latest $ev
      if ($cmpWingVsEver -gt 0) { $latestToShow = "$($chan.WingetLatest) (waiting on Evergreen release)"; $latestCss='info' }
    }
    $cmp = Compare-VersionSmart -Installed $installed -Latest $ev
    if     ($cmp -eq 0) { $status = "Latest Evergreen stable version installed as of $todayUK"; $css='ok' }
    elseif ($cmp -lt 0) { $status = "Update available (vs Evergreen Stable)"; $css='bad'; if(-not $latestCss){$latestCss='bad'} }
    else                { $status = "Ahead of Evergreen Stable as of $todayUK"; $css='ok' }
  }
  else {
    if ($chan.WingetLatest) {
      $cmp = Compare-VersionSmart -Installed $installed -Latest $chan.WingetLatest
      if     ($cmp -eq 0) { $latestToShow=$chan.WingetLatest; $status="Latest version installed as of $todayUK"; $css='ok' }
      elseif ($cmp -gt 0) { $latestToShow=$chan.WingetLatest; $status="Ahead of catalog as of $todayUK"; $css='ok' }
      else {
        if ($app.ExpectEvergreen) { $latestToShow="$($chan.WingetLatest) (Evergreen not available)"; $status="Evergreen not available"; $css='bad'; $latestCss='bad' }
        else { $latestToShow=$chan.WingetLatest; $status="Update available (catalog)"; $css='bad'; $latestCss='bad' }
      }
    } else { $latestToShow='-'; $status="Unknown (no latest info)"; $css='' }
  }

  if (-not $latestToShow) { $latestToShow='-' }
  $rows += [pscustomobject]@{
    Computer=$machine; Application=$app.Name; Installed=$installed; Latest=$latestToShow; Status=$status; Css=$css; LatestCss=$latestCss
  }
}

# Save per-machine CSV (apps)
$machineCsv = Join-Path $monthFolder ("{0} AppStatus {1}.csv" -f $machine,$todayIso)
$rows | Sort-Object Application | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $machineCsv
Write-Host "Per-machine report written: $machineCsv" -ForegroundColor Green

# Save OS updates (we filter later)
$osRowsAll = Get-RecentOSUpdates
$osCsv  = Join-Path $monthFolder ("{0} OSUpdates {1}.csv" -f $machine,$todayIso)
$osRowsAll | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $osCsv
Write-Host "OS updates saved: $osCsv" -ForegroundColor DarkCyan

# Ask to combine unless instructed otherwise
if (-not $NoCombine -and -not $Combine) {
  $ans = Read-Host "Create combined report for '$Period'? (Y/N)"
  if ($ans -match '^(y|yes)$') { $Combine = $true }
}

# ---------------------- Combine ----------------------
if ($Combine) {
  Write-Host "Building combined report for $Period..." -ForegroundColor Cyan

  $appFiles = Get-ChildItem -Path $monthFolder -Filter "* AppStatus *.csv" -ErrorAction SilentlyContinue
  if (-not $appFiles) { Write-Warning "No AppStatus CSVs found."; return }
  $allApps = @(); foreach($f in $appFiles){ $allApps += Import-CsvRobust -Path $f.FullName }

  $osFiles = Get-ChildItem -Path $monthFolder -Filter "* OSUpdates *.csv" -ErrorAction SilentlyContinue
  $allOs   = @(); foreach($f in ($osFiles | Where-Object { $_ })){ $allOs += Import-CsvRobust -Path $f.FullName }

  # NEW: OS Info files
  $osInfoFiles = Get-ChildItem -Path $monthFolder -Filter "* OSInfo *.csv" -ErrorAction SilentlyContinue
  $osInfoMap = @{}  # Computer -> info object
  foreach($f in $osInfoFiles){
    $info = Import-CsvRobust -Path $f.FullName
    foreach($row in $info){
      if (-not $osInfoMap.ContainsKey($row.Computer)) { $osInfoMap[$row.Computer] = $row }
    }
  }

  function Convert-ToDate {
    param($s)
    if ($s -is [datetime]) { return $s }
    foreach ($fmt in @('dd/MM/yyyy HH:mm:ss','dd/MM/yyyy','MM/dd/yyyy HH:mm:ss','MM/dd/yyyy','yyyy-MM-ddTHH:mm:ss','yyyy-MM-dd HH:mm:ss')) {
      try { return [datetime]::ParseExact("$s",$fmt,[System.Globalization.CultureInfo]::InvariantCulture) } catch {}
    }
    try { return [datetime]::Parse("$s",[System.Globalization.CultureInfo]::GetCultureInfo('en-GB')) } catch {}
    try { return (Get-Date -Date $s -ErrorAction Stop) } catch { return $null }
  }
  $allOs = $allOs | ForEach-Object { if ($_.InstalledOn) { $_.InstalledOn = Convert-ToDate $_.InstalledOn }; $_ }

  if (-not $allApps -or $allApps.Count -eq 0) { Write-Warning "No data rows to combine."; return }

  $baseName     = "Citrix Monthly patching report $todayIso"
  $combinedCsv  = Join-Path $monthFolder ($baseName + ".csv")
  $combinedHtml = Join-Path $monthFolder ($baseName + ".html")
  $combinedPdf  = Join-Path $monthFolder ($baseName + ".pdf")

  $allApps | Sort-Object Computer, Application | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $combinedCsv
  Write-Host "Combined CSV written: $combinedCsv" -ForegroundColor Green

  $null = New-CombinedHtml -AppData $allApps -OsData $allOs -OsInfoMap $osInfoMap `
           -Title "Citrix Monthly patching report ($Period)" `
           -OutPath $combinedHtml -Map $MasterMap -CatalogMap $CatalogMap -Period $Period
  Write-Host "Combined HTML written: $combinedHtml" -ForegroundColor Green

  Convert-HtmlToPdf -HtmlPath $combinedHtml -PdfPath $combinedPdf | Out-Null
}

