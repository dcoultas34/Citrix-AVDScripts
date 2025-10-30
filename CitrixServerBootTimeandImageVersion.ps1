<# 
Citrix multi-session image compliance report (catalog-only + per-catalog summary)
- Columns: Catalog, Machine, DNS, Registration, Power, Last Boot (derived), Image Status, Image Deployed, Image Version, Image Source, VDA Version
- Delivery Group removed from display
- Per-catalog summary (top): Image metadata + counts (Total / Up-to-date / Out-of-date)
- “Last Boot (derived)” = Broker.LastDeregistrationTime -> LastRegistrationTime -> LastConnectionTime -> (optional) Monitor OData fallback
- Image metadata (per catalog) = Get-ProvScheme: MasterImageVMDate (+ Version + Source)
#>

[CmdletBinding()]
param(
    [string]$OutputDir    = "\\transfer\transfer\Citrixapps\Citrix reports\Citrix reboots",
    [string]$ProfileName  = "",
    [switch]$ForceLogin,

    # Optional OData fallback (only used if all Broker fields are null)
    [string]$ClientId     = "",
    [string]$ClientSecret = "",
    [string]$CustomerId   = "",
    # Region base (Japan: https://api.citrixcloud.jp)
    [string]$ApiBase      = "https://api.cloud.com"
)

# ---------- Console cleanup ----------
$WarningPreference  = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

# ---------- Setup ----------
$null = New-Item -Path $OutputDir -ItemType Directory -Force -ErrorAction SilentlyContinue
$ts = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
$htmlPath = Join-Path $OutputDir "Citrix_Image_Compliance_$ts.html"
$pdfPath  = [System.IO.Path]::ChangeExtension($htmlPath, ".pdf")

# ---------- PDF helper ----------
function Convert-HtmlToPdf {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$HtmlFile,[Parameter(Mandatory=$true)][string]$PdfFile)
    $edgePaths = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe"
    ) | Where-Object { Test-Path $_ }
    if ($edgePaths.Count -gt 0) {
        $edge    = $edgePaths[0]
        $tmpHtml = Join-Path $env:TEMP ([IO.Path]::GetFileName($HtmlFile))
        Copy-Item -Path $HtmlFile -Destination $tmpHtml -Force
        $tmpPdf  = [System.IO.Path]::ChangeExtension($tmpHtml, ".pdf")
        $args    = @("--headless","--disable-gpu","--print-to-pdf=$tmpPdf","file:///$($tmpHtml -replace '\\','/')")
        $p       = Start-Process -FilePath $edge -ArgumentList $args -PassThru -WindowStyle Hidden
        $p.WaitForExit()
        if (Test-Path $tmpPdf) {
            Copy-Item -Path $tmpPdf -Destination $PdfFile -Force
            Remove-Item $tmpHtml,$tmpPdf -Force -ErrorAction SilentlyContinue
            return $true
        }
    }
    $wk = (Get-Command wkhtmltopdf.exe -ErrorAction SilentlyContinue)
    if ($wk) {
        & $wk.Source $HtmlFile $PdfFile | Out-Null
        if (Test-Path $PdfFile) { return $true }
    }
    return $false
}

# ---------- Citrix SDK check ----------
if (-not (Get-Command Get-XDAuthenticationEx -ErrorAction SilentlyContinue)) { throw "Citrix DaaS Remote PowerShell SDK not found." }
if (-not (Get-PSSnapin -Name Citrix.* -ErrorAction SilentlyContinue)) { asnp Citrix.* }

# ---------- Auth (reuse if valid) ----------
function Ensure-CitrixAuth {
    [CmdletBinding()]
    param([string]$ProfileName,[switch]$ForceLogin)
    if ($ForceLogin) {
        if ($ProfileName) { Get-XDAuthenticationEx -ProfileName $ProfileName | Out-Null }
        else              { Get-XDAuthenticationEx | Out-Null }
        return
    }
    try { Get-BrokerSite -ErrorAction Stop | Out-Null }
    catch {
        if ($ProfileName) { Get-XDAuthenticationEx -ProfileName $ProfileName | Out-Null }
        else              { Get-XDAuthenticationEx | Out-Null }
    }
}
Ensure-CitrixAuth -ProfileName $ProfileName -ForceLogin:$ForceLogin

# ---------- Optional OData (fallback only) ----------
function Get-BearerToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][string]$ClientSecret,
        [Parameter(Mandatory=$true)][string]$CustomerId,
        [Parameter(Mandatory=$true)][string]$ApiBase
    )
    $body = @{ grant_type='client_credentials'; client_id=$ClientId; client_secret=$ClientSecret }
    $url  = "$ApiBase/cctrustoauth2/$CustomerId/tokens/clients"
    $r    = Invoke-RestMethod -Uri $url -Method POST -Body $body
    "CWSAuth bearer=$($r.access_token)"
}
function UrlEnc([string]$s){ [System.Uri]::EscapeDataString($s) }
function Get-MonitorLastBoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$TokenHeader,
        [Parameter(Mandatory=$true)][string]$CustomerId,
        [Parameter(Mandatory=$true)][string]$ApiBase,
        [Parameter(Mandatory=$true)][string]$MachineName,
        [string]$DNSName = "",
        [string]$HostedMachineName = ""
    )
    $headers = @{
        "Authorization"     = $TokenHeader
        "Citrix-CustomerId" = $CustomerId
        "Accept"            = "application/json"
    }
    $short = $MachineName
    if ($MachineName -match "\\") { $short = $MachineName.Split("\")[1] }
    $dnsShort = if ($DNSName -and $DNSName.Contains(".")) { $DNSName.Split(".")[0] } else { $null }

    $cands = ($MachineName, $short, $HostedMachineName, $DNSName, $dnsShort) | Where-Object { $_ -and $_.Trim() -ne "" }
    $parts = @()
    foreach($cand in $cands){
        $lc = $cand.ToLower()
        $parts += @(
          "tolower(MachineName) eq '$lc'",
          "tolower(HostedMachineName) eq '$lc'",
          "tolower(DNSName) eq '$lc'",
          "tolower(FullyQualifiedDomainName) eq '$lc'",
          "endswith(tolower(DNSName),'$lc')",
          "endswith(tolower(FullyQualifiedDomainName),'$lc')"
        )
    }
    if (-not $parts) { return $null }
    $filter = ($parts | Select-Object -Unique) -join " or "
    $url = "$ApiBase/monitorodata/Machines?`$filter=$(UrlEnc $filter)&`$top=1"

    $priority = @('LastBootTime','OSLastBootTime','LastPowerOnTime','LastRegistrationTime','LastConnectionTime','LastDeregistrationTime')

    for($i=1;$i -le 3;$i++){
        try{
            $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method GET -ErrorAction Stop
            if ($resp.value -and $resp.value.Count -gt 0){
                $row = $resp.value[0]
                foreach($p in $priority){
                    if ($row.PSObject.Properties.Name -contains $p){
                        $v = $row.$p
                        if ($v){
                            try{
                                $dt = [datetime]$v
                                if ($dt.Year -gt 1901){ return $dt }
                            }catch{}
                        }
                    }
                }
            }
            return $null
        }catch{
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode.value__ -eq 429 -and $i -lt 3){
                Start-Sleep -Seconds (2*$i)
            }else{
                return $null
            }
        }
    }
    return $null
}

# ---------- Broker data ----------
try {
    $catalogs    = Get-BrokerCatalog -SessionSupport MultiSession -MaxRecordCount 100000
    $machinesRaw = foreach ($c in $catalogs) { Get-BrokerMachine -CatalogUid $c.Uid -MaxRecordCount 100000 }
} catch {
    Write-Error "Failed to retrieve Broker data: $_"
    exit 1
}

# Quick maps
$catNameByUid = @{}; foreach ($c in $catalogs) { $catNameByUid[$c.Uid] = $c.Name }

# ---------- Provisioning schemes (image metadata per catalog) ----------
$provAvailable = $false
$provById   = @{}
$provByName = @{}
try {
    if (Get-Command Get-ProvScheme -ErrorAction SilentlyContinue) {
        $provAvailable = $true
        $provSchemes = Get-ProvScheme -MaxRecordCount 100000
        foreach ($p in $provSchemes) {
            if ($p.PSObject.Properties.Name -contains 'Uid' -and $p.Uid) { $provById[[string]$p.Uid] = $p }
            if ($p.PSObject.Properties.Name -contains 'ProvisioningSchemeUid' -and $p.ProvisioningSchemeUid) { $provById[[string]$p.ProvisioningSchemeUid] = $p }
            if ($p.PSObject.Properties.Name -contains 'ProvisioningSchemeName' -and $p.ProvisioningSchemeName) {
                $provByName[$p.ProvisioningSchemeName] = $p
            }
        }
    }
} catch {
    Write-Warning "Could not query provisioning schemes (Get-ProvScheme): $_"
    $provAvailable = false
}

function Get-SchemeForCatalog {
    param([object]$Catalog)
    # Try by ProvisioningSchemeId (Guid) then by name
    if ($Catalog.PSObject.Properties.Name -contains 'ProvisioningSchemeId' -and $Catalog.ProvisioningSchemeId) {
        $id = [string]$Catalog.ProvisioningSchemeId
        if ($provById.ContainsKey($id)) { return $provById[$id] }
        $idU = $id.ToUpper(); $idL = $id.ToLower()
        if ($provById.ContainsKey($idU)) { return $provById[$idU] }
        if ($provById.ContainsKey($idL)) { return $provById[$idL] }
    }
    if ($Catalog.Name -and $provByName.ContainsKey($Catalog.Name)) { return $provByName[$Catalog.Name] }
    return $null
}

# ---------- OData token (only if provided) ----------
$odataEnabled = ($ClientId -and $ClientSecret -and $CustomerId)
$tokenHeader  = $null
if ($odataEnabled) {
    try { $tokenHeader = Get-BearerToken -ClientId $ClientId -ClientSecret $ClientSecret -CustomerId $CustomerId -ApiBase $ApiBase }
    catch { Write-Warning "Failed to obtain Monitor token: $_"; $odataEnabled = $false }
}

# ---------- Shape final data (catalog-only view) ----------
$machines = foreach ($m in $machinesRaw) {
    $catalogName = $catNameByUid[$m.CatalogUid]

    # Last Boot (derived): broker-first
    $lastBoot = $null
    if     ($m.LastDeregistrationTime) { $lastBoot = $m.LastDeregistrationTime }
    elseif ($m.LastRegistrationTime)   { $lastBoot = $m.LastRegistrationTime }
    elseif ($m.LastConnectionTime)     { $lastBoot = $m.LastConnectionTime }
    elseif (-not $lastBoot -and $odataEnabled -and $tokenHeader) {
        $lastBoot = Get-MonitorLastBoot -TokenHeader $tokenHeader -CustomerId $CustomerId -ApiBase $ApiBase `
                     -MachineName $m.MachineName -DNSName $m.DNSName -HostedMachineName $m.HostedMachineName
    }

    # Image metadata (per catalog)
    $imgDate    = $null; $imgVersion = $null; $imgSource = $null
    if ($provAvailable) {
        $catObj = ($catalogs | Where-Object Uid -EQ $m.CatalogUid)
        if ($catObj) {
            $scheme = Get-SchemeForCatalog -Catalog $catObj
            if ($scheme) {
                if ($scheme.PSObject.Properties.Name -contains 'MasterImageVMDate') { $imgDate = $scheme.MasterImageVMDate }
                if ($scheme.PSObject.Properties.Name -contains 'ProvisioningSchemeVersion') { $imgVersion = $scheme.ProvisioningSchemeVersion }
                if ($scheme.PSObject.Properties.Name -contains 'MasterImageVM') { $imgSource = $scheme.MasterImageVM }
            }
        }
    }

    [pscustomobject]@{
        CatalogName       = $catalogName
        MachineName       = $m.MachineName
        DNSName           = $m.DNSName
        RegistrationState = $m.RegistrationState
        PowerState        = $m.PowerState
        LastBootDerived   = $lastBoot
        ImageOutOfDate    = [bool]$m.ImageOutOfDate
        AgentVersion      = $m.AgentVersion
        ImageDeployed     = $imgDate
        ImageVersion      = $imgVersion
        ImageSource       = $imgSource
    }
}

$machines = $machines | Sort-Object CatalogName, MachineName

# ---------- Global counts ----------
$totalCount = $machines.Count
$oldCount   = ($machines | Where-Object { $_.ImageOutOfDate }).Count
$newCount   = $totalCount - $oldCount

# ---------- Build per-catalog summary ----------
$catalogSummaryRows = @()
$catalogGroups = $machines | Group-Object CatalogName
foreach ($g in $catalogGroups) {
    $name = $g.Name
    $grp  = $g.Group
    $ct   = $grp.Count
    $old  = ($grp | Where-Object { $_.ImageOutOfDate }).Count
    $new  = $ct - $old

    # use first row's image metadata for the catalog
    $first = $grp | Select-Object -First 1
    $imgDateStr = ""; if ($first.ImageDeployed) { $imgDateStr = (Get-Date $first.ImageDeployed).ToString("yyyy-MM-dd HH:mm") }
    $imgSourceShort = ""; if ($first.ImageSource) { $imgSourceShort = ($first.ImageSource -replace '^.+[\\\/]', '') }

    $catalogSummaryRows += @"
<tr>
  <td>$name</td>
  <td>$imgDateStr</td>
  <td>$($first.ImageVersion)</td>
  <td><span class="small">$imgSourceShort</span></td>
  <td style="text-align:right;">$ct</td>
  <td style="text-align:right;">$new</td>
  <td style="text-align:right;">$old</td>
</tr>
"@
}

# ---------- HTML ----------
$now = Get-Date
$style = @"
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;}
h1{margin-bottom:0;}
.meta{color:#555;margin-top:4px;}
.summarypill{margin-top:10px;margin-bottom:15px;}
.summarypill span{display:inline-block;margin-right:12px;padding:6px 12px;border-radius:999px;font-weight:600;}
.total{background:#e8e8e8;}
.ok{background:#e6ffed;color:#006400;}
.bad{background:#ffecec;color:#b30000;}
table{border-collapse:collapse;width:100%;margin-top:16px;}
th,td{border:1px solid #ddd;padding:8px;font-size:13px;}
th{background:#f2f2f2;position:sticky;top:0;}
tr:nth-child(even){background:#fafafa;}
tr.outdated{background:#ffecec !important;}
tr.outdated td{color:#b30000;font-weight:600;}
.footer{margin-top:24px;color:#777;font-size:12px;}
.small{font-size:12px;color:#666;}
.section-title{margin-top:22px;font-size:20px;}
</style>
"@

$header = @"
<h1>Citrix Image Compliance — Multi-Session</h1>
<div class="meta">Generated: $($now.ToString('yyyy-MM-dd HH:mm:ss'))</div>
<div class="summarypill">
  <span class="total">Total servers: $totalCount</span>
  <span class="ok">Up-to-date: $newCount</span>
  <span class="bad">Out-of-date: $oldCount</span>
</div>

<h2 class="section-title">Per-Catalog Summary</h2>
<table>
  <thead>
    <tr>
      <th>Catalog</th><th>Image Deployed</th><th>Image Version</th><th>Image Source</th>
      <th style="text-align:right;">Total</th><th style="text-align:right;">Up-to-date</th><th style="text-align:right;">Out-of-date</th>
    </tr>
  </thead>
  <tbody>
    $($catalogSummaryRows -join "`n")
  </tbody>
</table>
"@

$rows = foreach ($m in $machines) {
    $cls = if ($m.ImageOutOfDate) { "outdated" } else { "" }
    $status = if ($m.ImageOutOfDate) { '<span class="bad">Old image</span>' } else { '<span class="ok">Current</span>' }
    $bootStr = ""; if ($m.LastBootDerived) { $bootStr = (Get-Date $m.LastBootDerived).ToString("yyyy-MM-dd HH:mm") }
    $imgDateStr = ""; if ($m.ImageDeployed) { $imgDateStr = (Get-Date $m.ImageDeployed).ToString("yyyy-MM-dd HH:mm") }
    $imgSourceShort = ""; if ($m.ImageSource) { $imgSourceShort = ($m.ImageSource -replace '^.+[\\\/]', '') }
@"
<tr class="$cls">
  <td>$($m.CatalogName)</td>
  <td>$($m.MachineName)</td>
  <td>$($m.DNSName)</td>
  <td>$($m.RegistrationState)</td>
  <td>$($m.PowerState)</td>
  <td>$bootStr</td>
  <td>$status</td>
  <td>$imgDateStr</td>
  <td>$($m.ImageVersion)</td>
  <td><span class="small">$imgSourceShort</span></td>
  <td>$($m.AgentVersion)</td>
</tr>
"@
}

$table = @"
<h2 class="section-title">Machine Detail</h2>
<table>
  <thead>
    <tr>
      <th>Catalog</th><th>Machine</th><th>DNS</th>
      <th>Registration</th><th>Power</th><th>Last Boot (derived)</th>
      <th>Image Status</th><th>Image Deployed</th><th>Image Version</th><th>Image Source</th>
      <th>VDA Version</th>
    </tr>
  </thead>
  <tbody>
    $($rows -join "`n")
  </tbody>
</table>
"@

$footer = @"
<div class="footer">
Last Boot (derived) priority: Broker LastDeregistrationTime > LastRegistrationTime > LastConnectionTime > Monitor OData (fallback, if configured).<br/>
Image metadata from provisioning schemes (Get-ProvScheme): MasterImageVMDate (Image Deployed), ProvisioningSchemeVersion (Image Version), MasterImageVM (Image Source).<br/>
Note: For non-MCS catalogs (e.g., PVS), image fields may be blank.
</div>
"@

$html = "<html><head><meta charset='utf-8'/>$style</head><body>$header$table$footer</body></html>"
$html | Out-File -FilePath $htmlPath -Encoding UTF8 -Force
Write-Host "HTML report written to: $htmlPath"

# ---------- PDF (best effort) ----------
if (Convert-HtmlToPdf -HtmlFile $htmlPath -PdfFile $pdfPath) {
    Write-Host "PDF report written to: $pdfPath"
} else {
    Write-Warning "PDF export not available (Edge headless or wkhtmltopdf not found). HTML generated."
}

Write-Host "Done. ($totalCount servers total — $newCount up-to-date, $oldCount old image)"
