# Get-LatestEvergreen.ps1
# Usage examples:
#   .\Get-LatestEvergreen.ps1 -Name AdobeAcrobatReaderDC -Architecture x64 -Language en-US
#   '7zip','GoogleChrome','AdobeAcrobatReaderDC' | .\Get-LatestEvergreen.ps1
#   .\Get-LatestEvergreen.ps1 -Name GoogleChrome -All | Format-Table -AutoSize

[CmdletBinding()]
param(
    # One or more Evergreen app names. Supports pipeline input.
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
    [string[]]$Name,

    # Optional filters; applied only if the app exposes the property.
    [ValidateSet('x64','x86','arm64','Any')]
    [string]$Architecture = 'Any',

    [string]$Language = $null,
    [string]$Channel = $null,   # e.g. Stable/Beta/Continuous (app-dependent)

    # Return all matching rows instead of only the newest record.
    [switch]$All
)

begin {
    $ErrorActionPreference = 'Stop'

    # Ensure NuGet provider & Evergreen module (install to CurrentUser if missing)
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -Force | Out-Null
    }
    if (-not (Get-Module -ListAvailable -Name Evergreen)) {
        Install-Module -Name Evergreen -Scope CurrentUser -Force -Repository PSGallery
    }
    Import-Module Evergreen -Force

    function Select-BestMatch {
        param(
            [Parameter(Mandatory=$true)][object[]]$InputObjects,
            [string]$Architecture,
            [string]$Language,
            [string]$Channel,
            [switch]$All
        )

        $rows = $InputObjects

        # Helper: does property exist on the objects?
        function Has-Prop($prop) { ($rows | Select-Object -First 1).PSObject.Properties.Name -contains $prop }

        if ($Architecture -and $Architecture -ne 'Any' -and (Has-Prop 'Architecture')) {
            $rows = $rows | Where-Object { $_.Architecture -eq $Architecture }
        }
        if ($Language -and (Has-Prop 'Language')) {
            $rows = $rows | Where-Object { $_.Language -eq $Language }
        }
        if ($Channel -and (Has-Prop 'Channel')) {
            $rows = $rows | Where-Object { $_.Channel -eq $Channel }
        }

        # Sort by Version (desc) if present; fall back to Date/Release if available
        if ($rows -and (Has-Prop 'Version')) {
            $rows = $rows | Sort-Object {
                try { [version]($_.Version) } catch { [version]'0.0.0.0' }
            } -Descending
        } elseif ($rows -and (Has-Prop 'Date')) {
            $rows = $rows | Sort-Object Date -Descending
        } elseif ($rows -and (Has-Prop 'Release')) {
            $rows = $rows | Sort-Object Release -Descending
        }

        if ($All) { return $rows }
        return ($rows | Select-Object -First 1)
    }

    function Format-Output {
        param([object]$row, [string]$AppName)

        if (-not $row) {
            return [pscustomobject]@{
                AppName      = $AppName
                Version      = $null
                Architecture = $null
                Language     = $null
                Channel      = $null
                Type         = $null
                SizeMB       = $null
                SHA256       = $null
                Uri          = $null
                Note         = "No results (check spelling or filters)."
            }
        }

        $props = $row.PSObject.Properties.Name
        $sizeMB = $null
        if ($props -contains 'Size' -and $row.Size) { $sizeMB = [math]::Round(($row.Size / 1MB), 1) }

        [pscustomobject]@{
            AppName      = $AppName
            Version      = $(if ($props -contains 'Version')      { $row.Version }      else { $null })
            Architecture = $(if ($props -contains 'Architecture') { $row.Architecture } else { $null })
            Language     = $(if ($props -contains 'Language')     { $row.Language }     else { $null })
            Channel      = $(if ($props -contains 'Channel')      { $row.Channel }      else { $null })
            Type         = $(if ($props -contains 'Type')         { $row.Type }         else { $null })
            SizeMB       = $sizeMB
            SHA256       = $(if ($props -contains 'SHA256')       { $row.SHA256 }       else { $null })
            Uri          = $(if ($props -contains 'Uri')          { $row.Uri }          else { $null })
            Note         = $null
        }
    }
}

process {
    foreach ($app in $Name) {
        try {
            $raw = Get-EvergreenApp -Name $app -ErrorAction Stop
        } catch {
            Write-Warning "Evergreen could not find app '$app' ($($_.Exception.Message))"
            $raw = $null
        }

        if (-not $raw) {
            Format-Output -row $null -AppName $app
            continue
        }

        $selected = Select-BestMatch -InputObjects $raw -Architecture $Architecture -Language $Language -Channel $Channel -All:$All

        if ($All) {
            foreach ($row in $selected) { Format-Output -row $row -AppName $app }
        } else {
            Format-Output -row $selected -AppName $app
        }
    }
}
