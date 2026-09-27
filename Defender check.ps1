$Result = "" | Select-Object `
    @{Name='ComputerName';Expression={$env:COMPUTERNAME}},
    @{Name='ScanDate';Expression={Get-Date -Format "yyyy-MM-dd HH:mm:ss"}},
    @{Name='DefenderAVEnabled';Expression={
        if ($Defender) {
            if ($Defender.AntivirusEnabled -eq $true) { "Enabled" }
            elseif ($Defender.AntivirusEnabled -eq $false) { "Disabled" }
            else { "Unknown" }
        }
        else { "Unknown" }
    }},
    @{Name='DefenderRunningMode';Expression={
        if ($Defender) { $Defender.AMRunningMode }
        else { "Unknown" }
    }},
    @{Name='RealTimeProtection';Expression={
        if ($Defender) {
            if ($Defender.RealTimeProtectionEnabled -eq $true) { "Enabled" }
            elseif ($Defender.RealTimeProtectionEnabled -eq $false) { "Disabled" }
            else { "Unknown" }
        }
        else { "Unknown" }
    }},
    @{Name='BehaviorMonitoring';Expression={
        if ($Defender) {
            if ($Defender.BehaviorMonitorEnabled -eq $true) { "Enabled" }
            elseif ($Defender.BehaviorMonitorEnabled -eq $false) { "Disabled" }
            else { "Unknown" }
        }
        else { "Unknown" }
    }},
    @{Name='IOAVProtection';Expression={
        if ($Defender) {
            if ($Defender.IoavProtectionEnabled -eq $true) { "Enabled" }
            elseif ($Defender.IoavProtectionEnabled -eq $false) { "Disabled" }
            else { "Unknown" }
        }
        else { "Unknown" }
    }},
    @{Name='NetworkInspection';Expression={
        if ($Defender) {
            if ($Defender.NISEnabled -eq $true) { "Enabled" }
            elseif ($Defender.NISEnabled -eq $false) { "Disabled" }
            else { "Unknown" }
        }
        else { "Unknown" }
    }},
    @{Name='SignatureVersion';Expression={
        if ($Defender) { $Defender.AntivirusSignatureVersion }
        else { "Unknown" }
    }},
    @{Name='SignatureLastUpdated';Expression={
        if ($Defender) { $Defender.AntivirusSignatureLastUpdated }
        else { "Unknown" }
    }},
    @{Name='SignaturesOutOfDate';Expression={
        if ($Defender) {
            if ($Defender.DefenderSignaturesOutOfDate -eq $true) { "Yes" }
            elseif ($Defender.DefenderSignaturesOutOfDate -eq $false) { "No" }
            else { "Unknown" }
        }
        else { "Unknown" }
    }},
    @{Name='TamperProtection';Expression={
        if ($Defender) {
            if ($Defender.IsTamperProtected -eq $true) { "Enabled" }
            elseif ($Defender.IsTamperProtected -eq $false) { "Disabled" }
            else { "Unknown" }
        }
        else { "Unknown" }
    }},
    @{Name='MDEOnboardingState';Expression={$MDEStatus}},
    @{Name='MDEOnboardingValue';Expression={
        if ($MDE) { $MDE.OnboardingState }
        else { "Unknown" }
    }},
    @{Name='MDESenseService';Expression={
        if ($SenseService) { $SenseService.Status }
        else { "Not Found" }
    }},
    @{Name='DefenderService';Expression={
        if ($DefenderService) { $DefenderService.Status }
        else { "Not Found" }
    }},
    @{Name='FirewallDomain';Expression={
        if ($DomainFirewall -eq $true) { "Enabled" }
        elseif ($DomainFirewall -eq $false) { "Disabled" }
        else { "Unknown" }
    }},
    @{Name='FirewallPrivate';Expression={
        if ($PrivateFirewall -eq $true) { "Enabled" }
        elseif ($PrivateFirewall -eq $false) { "Disabled" }
        else { "Unknown" }
    }},
    @{Name='FirewallPublic';Expression={
        if ($PublicFirewall -eq $true) { "Enabled" }
        elseif ($PublicFirewall -eq $false) { "Disabled" }
        else { "Unknown" }
    }},
    @{Name='SecuritySoftware';Expression={$SecuritySoftware}},
    @{Name='ProtectionStatus';Expression={$ProtectionStatus}},
    @{Name='ReviewNotes';Expression={$ReviewNotes}}

return $Result
