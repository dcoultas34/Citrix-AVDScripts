<# 
.SYNOPSIS
  Look up the latest version(s) of apps via the Evergreen module.

.DESCRIPTION
  - Accepts one or many app names supported by Evergreen (e.g., 7zip, GoogleChrome, AdobeAcrobatReaderDC).
  - Lets you filter by Architecture, Language, and Channel if an app exposes those fields.
  - Prints a clean, single “best match” record per app (newest version), or all matches with -All.

.EXAMPLES
  .\Get-LatestEvergreen.ps1 -Name AdobeAcrobatReaderDC -Architecture x64 -Language en-US
  '7zip','GoogleChrome','AdobeAcrobatReaderDC' | .\Get-LatestEvergreen.ps1
  .\Get-LatestEvergreen.ps1 -Name GoogleChrome -All | Format-Table

.NOTES
  - Requires PowerShell 5.1+ (or PowerShell 7+).
  - Installs Evergreen for the current user if missing.
#>

[CmdletBinding()]
param(
    # One or more Evergreen app names. Supports pipeline input.
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
    [string[]]$Name,

    # Optional filters (only applied if the returned objects have these properties).
    [ValidateSet('x64','x86','arm64','Any')]
    [string]$Architecture = 'Any',

    [string]$Language = $null,

    [string]$Channel = $null,  # e.g., "Stable", "Beta", "Continuous", etc. (app-dependent)

    # Return *all* matching rows instead of the single newest record.
    [switch]$All
)

begin {
    $ErrorActionPreference = 'Stop'

    # Ensure NuGet provider & Evergreen module
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

        # Apply Architecture filter if present and if the objects expose that property
        if ($Architecture -and $Architecture -ne 'Any' -and ($rows | Get-Member -Name Architecture -MemberType NoteProperty,Property)) {
            $rows = $rows | Where-Object { $_.Architecture -eq $Architecture }
        }

        # Apply Language filter if property exists
        if ($Language -and ($rows | Get-Member -Name Language -MemberType NoteProperty,Property)) {
            $rows = $rows | Where-Object { $_.Language -eq $Language }
        }

        # Apply Channel filter if property exists
        if ($Channel -and ($rows | Get-Member -Name Channel -MemberType NoteProperty,Property)) {
            $rows = $rows | Where-Object { $_.Channel -eq $Channel }
        }

        # Sort by Version (desc) if present, otherwise by Date or Release if present, else leave as-is
        if ($rows -and ($rows | Get-Member -Name Version -MemberType NoteProperty,Property)) {
            $rows = $rows | Sort-Object { [version]($_.Version -as [string]) } -Descending, Version -Descending
        } elseif ($rows -and ($rows | Get-Member -Name Date -MemberType NoteProperty,Property)) {
            $rows = $rows | Sort-Object Date -Descending
        } elseif ($rows -and ($rows | Get-Member -Name Release -MemberType NoteProperty,Property)) {
            $rows = $rows | Sort-Object Release -Descending
        }

        if ($All) { return $rows }

        # Return just the single “best” record
        return ($rows | Select-Object -First 1)
    }

    function Format-Output {
        param([object]$row, [string]$AppName)

        if (-not $row) {
            [pscustomobject]@{
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
            return
        }

        # Handle Size if present
        $sizeMB = $null
        if ($row.PSObject.Properties.Name -contains 'Size' -and $row.Size) {
            $sizeMB = [math]::Round(($row.Size / 1MB), 1)
        }

        [pscustomobject]@{
            AppName      = $AppName
            Version      = $row.Version
            Architecture = $row.PSObject.Properties['Architecture']?.Value
            Language     = $row.PSObject.Properties['Language']?.Value
            Channel      = $row.PSObject.Properties['Channel']?.Value
            Type         = $row.PSObject.Properties['Type']?.Value
            SizeMB       = $sizeMB
            SHA256       = $row.PSObject.Properties['SHA256']?.Value
            Uri          = $row.PSObject.Properties['Uri']?.Value
            Note         = $null
        }
    }
}

process {
    foreach ($app in $Name) {
        try {
            $raw = Get-EvergreenApp -Name $app -ErrorAction Stop
        } catch {
            # If a single bad name was given among many, keep going
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
