<#
.SYNOPSIS
  Citrix Cloud / CVAD Usage HTML Report for last month.

.DESCRIPTION
  - For the specified Delivery Groups:
      * Counts unique users
      * Sums total session duration
      * Counts app launches per application
  - Builds per-user/per-session detail including client info
  - Outputs to HTML
  - Each Delivery Group has its own section & tables
#>

# -------------------- OPTIONAL: LOAD CITRIX SNAPIN/MODULE --------------------
# Uncomment ONE of these if needed in your environment:

# Add-PSSnapin Citrix.Monitor.Admin.V2 -ErrorAction SilentlyContinue
# Import-Module Citrix.Monitor.Admin.V2 -ErrorAction SilentlyContinue

# -------------------- CONFIGURATION --------------------

# Delivery Groups you care about
$DeliveryGroups = @(
    "Windows 10 MS UK South CTE",
    "Windows 10 MS UK South DEV",
    "Windows 10 MS UK South Standard Non-Prod",
    "Windows 10 MS UK South 3E Cloud Prod",
    "Windows 10 MS UK South AuditTrack Prod",
    "Windows 10 MS UK South DEVQA Prod",
    "Windows 10 MS UK South Ent Apps Prod",
    "Windows 10 MS UK South HR Systems Prod",
    "Windows 10 MS UK South iManage Prod",
    "Windows 10 MS UK South Standard Prod"
)

# Time window – last 30 days
$StartTime = (Get-Date).AddDays(-30)
$EndTime   = Get-Date

# Output file
$OutputHtmlPath = "C:\Reports\Citrix-Usage-Report.html"

# -------------------- DATA COLLECTION --------------------

Write-Host "Querying Monitor Service from $StartTime to $EndTime ..." -ForegroundColor Cyan

try {
    $sessions = Get-XdMonitorData -DataType Sessions -StartTime $StartTime -EndTime $EndTime
    $appInstances = Get-XdMonitorData -DataType ApplicationInstances -StartTime $StartTime -EndTime $EndTime
}
catch {
    Write-Error "Failed to run Get-XdMonitorData. Make sure the Citrix Monitor SDK / snapin is loaded. Error: $($_.Exception.Message)"
    return
}

if (-not $sessions)  { Write-Warning "No session data returned for period."; }
if (-not $appInstances) { Write-Warning "No application instance data returned for period."; }

# -------------------- OVERALL SUMMARY --------------------

# Filter sessions to only the Delivery Groups we care about
$filteredSessions = $sessions | Where-Object { $DeliveryGroups -contains $_.DeliveryGroupName }

# Unique users overall
$overallUniqueUsers = ($filteredSessions.UserName | Sort-Object -Unique).Count

# Total session duration overall
# Adjust this field name if your environment uses something else,
# e.g. SessionDuration, TotalSessionDuration, etc.
$durationField = "ConnectedDuration"

$overallTotalDurationSeconds = ($filteredSessions | Measure-Object -Property $durationField -Sum).Sum
$overallTotalDurationHours = if ($overallTotalDurationSeconds) {
    [math]::Round(($overallTotalDurationSeconds / 3600), 2)
} else { 0 }

# -------------------- PER DELIVERY GROUP SUMMARY (DATA) --------------------

$dgSummary = foreach ($dg in $DeliveryGroups) {

    $dgSessions = $filteredSessions | Where-Object { $_.DeliveryGroupName -eq $dg }

    if (-not $dgSessions) {
        [PSCustomObject]@{
            DeliveryGroup               = $dg
            UniqueUsers                 = 0
            TotalSessions               = 0
            TotalDurationHours          = 0
            AvgSessionDurationMinutes   = 0
        }
        continue
    }

    $uniqueUsers = ($dgSessions.UserName | Sort-Object -Unique).Count
    $totalSessions = $dgSessions.Count
    $totalDurationSec = ($dgSessions | Measure-Object -Property $durationField -Sum).Sum
    $totalDurationHours = if ($totalDurationSec) {
        [math]::Round(($totalDurationSec / 3600), 2)
    } else { 0 }

    $avgDurationMin = if ($totalSessions -gt 0 -and $totalDurationSec) {
        [math]::Round((($totalDurationSec / $totalSessions) / 60), 2)
    } else { 0 }

    [PSCustomObject]@{
        DeliveryGroup               = $dg
        UniqueUsers                 = $uniqueUsers
        TotalSessions               = $totalSessions
        TotalDurationHours          = $totalDurationHours
        AvgSessionDurationMinutes   = $avgDurationMin
    }
}

# -------------------- APPLICATION USAGE SUMMARY (DATA) --------------------

# Join app instances to sessions via SessionKey (field name may vary: SessionKey, SessionId, etc.)
$appUsage = @()

if ($appInstances) {
    foreach ($app in $appInstances) {
        $sess = $filteredSessions | Where-Object { $_.SessionKey -eq $app.SessionKey } | Select-Object -First 1
        if ($null -eq $sess) { continue }

        $appUsage += [PSCustomObject]@{
            DeliveryGroup = $sess.DeliveryGroupName
            UserName      = $sess.UserName
            Application   = $app.ApplicationName
        }
    }
}

$appSummary = @()
if ($appUsage.Count -gt 0) {
    $appSummary = $appUsage |
        Group-Object DeliveryGroup, Application |
        ForEach-Object {
            $dgApp = $_.Name.Split(",")
            [PSCustomObject]@{
                DeliveryGroup = ($dgApp[0]).Trim()
                Application   = ($dgApp[1]).Trim()
                LaunchCount   = $_.Count
            }
        } |
        Sort-Object DeliveryGroup, -Property LaunchCount
}

# -------------------- PER-USER / PER-SESSION DETAIL (DATA) --------------------

$sessionDetail = $filteredSessions | ForEach-Object {

    $start = $_.StartDateTime
    $end   = $_.EndDateTime
    $durSec = $_.$durationField

    if (-not $durSec -and $start -and $end) {
        $durSec = (New-TimeSpan -Start $start -End $end).TotalSeconds
    }

    $durMin = if ($durSec) { [math]::Round(($durSec / 60), 2) } else { 0 }

    [PSCustomObject]@{
        UserName           = $_.UserName
        DeliveryGroup      = $_.DeliveryGroupName
        MachineName        = $_.MachineName
        SessionStart       = $start
        SessionEnd         = $end
        SessionDurationMin = $durMin
        ClientName         = $_.ClientName
        ClientIPAddress    = $_.ClientIPAddress
    }
}

# -------------------- HTML REPORT BUILD --------------------

Write-Host "Building HTML report..." -ForegroundColor Cyan

$style = @"
<style>
body { font-family: Arial, sans-serif; font-size: 13px; }
h1, h2, h3 { color: #333333; }
table { border-collapse: collapse; margin-bottom: 20px; width: 100%; }
th, td { border: 1px solid #cccccc; padding: 4px 6px; text-align: left; }
th { background-color: #f2f2f2; }
.summary { background-color: #e9f5ff; padding: 10px; margin-bottom: 20px; border: 1px solid #bcdffb; }
.section { margin-top: 30px; }
</style>
"@

# Overall header
$overallSummaryHtml = @"
<div class='summary'>
  <h1>Citrix Usage Summary (Last 30 Days)</h1>
  <p><strong>Report Period:</strong> $($StartTime.ToString("yyyy-MM-dd HH:mm")) to $($EndTime.ToString("yyyy-MM-dd HH:mm"))</p>
  <p><strong>Delivery Groups included:</strong> $(($DeliveryGroups -join ", "))</p>
  <p><strong>Total Unique Users (all DGs):</strong> $overallUniqueUsers</p>
  <p><strong>Total Session Duration (all DGs):</strong> $overallTotalDurationHours hours</p>
</div>
"@

# Global DG summary table
$dgSummaryHtml = $dgSummary | ConvertTo-Html -Fragment -PreContent "<h2>Overall Delivery Group Summary</h2>"

# Per-DG sections
$dgSectionsHtml = ""

foreach ($dg in $DeliveryGroups) {

    $dgSummaryRow   = $dgSummary   | Where-Object { $_.DeliveryGroup -eq $dg }
    $dgApps         = $appSummary  | Where-Object { $_.DeliveryGroup -eq $dg }
    $dgSessionsDet  = $sessionDetail | Where-Object { $_.DeliveryGroup -eq $dg }

    $sectionHtml = "<div class='section'><h2>Delivery Group: $dg</h2>"

    # Mini summary table (single row)
    if ($dgSummaryRow) {
        $sectionHtml += ($dgSummaryRow | ConvertTo-Html -Fragment -PreContent "<h3>Summary</h3>")
    }
    else {
        $sectionHtml += "<p><em>No sessions for this Delivery Group in the selected period.</em></p>"
    }

    # App usage table for this DG
    if ($dgApps -and $dgApps.Count -gt 0) {
        $sectionHtml += ($dgApps | Select-Object Application, LaunchCount |
                         ConvertTo-Html -Fragment -PreContent "<h3>Application Usage</h3>")
    }
    else {
        $sectionHtml += "<h3>Application Usage</h3><p><em>No application launches recorded for this Delivery Group.</em></p>"
    }

    # Session detail table for this DG
    if ($dgSessionsDet -and $dgSessionsDet.Count -gt 0) {
        $sectionHtml += ($dgSessionsDet |
                         Select-Object UserName, MachineName, SessionStart, SessionEnd, SessionDurationMin, ClientName, ClientIPAddress |
                         ConvertTo-Html -Fragment -PreContent "<h3>User / Session Detail</h3>")
    }
    else {
        $sectionHtml += "<h3>User / Session Detail</h3><p><em>No session detail available for this Delivery Group.</em></p>"
    }

    $sectionHtml += "</div>"

    $dgSectionsHtml += $sectionHtml + "`n"
}

# Assemble full HTML body
$bodyContent = @"
$overallSummaryHtml
$dgSummaryHtml
$dgSectionsHtml
"@

$fullHtml = ConvertTo-Html -Title "Citrix Usage Report" -Head $style -Body $bodyContent

# Ensure output folder exists
$directory = Split-Path $OutputHtmlPath -Parent
if (-not (Test-Path $directory)) {
    New-Item -ItemType Directory -Path $directory | Out-Null
}

$fullHtml | Out-File -FilePath $OutputHtmlPath -Encoding UTF8

Write-Host "Report generated: $OutputHtmlPath" -ForegroundColor Green
