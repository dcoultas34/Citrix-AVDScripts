$Computer = $env:COMPUTERNAME

$Defender = Get-MpComputerStatus -ErrorAction SilentlyContinue

$MDE = Get-ItemProperty `
    'HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status' `
    -ErrorAction SilentlyContinue

$Sense = Get-Service Sense -ErrorAction SilentlyContinue

$SecuritySoftware = Get-ItemProperty `
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' `
    -ErrorAction SilentlyContinue |
    Where-Object {
        $_.DisplayName -match 'Defender|CrowdStrike|Sentinel|Sophos|McAfee|Trellix|Symantec|Carbon Black|Trend|Cylance|FireEye|Rapid7|Qualys|Nessus|Zscaler'
    } |
    Select-Object -ExpandProperty DisplayName -Unique

[PSCustomObject]@{

    ComputerName                 = $Computer

    DefenderAVEnabled            = $Defender.AntivirusEnabled
    DefenderRunningMode          = $Defender.AMRunningMode
    RealTimeProtection           = $Defender.RealTimeProtectionEnabled
    BehaviorMonitoring           = $Defender.BehaviorMonitorEnabled
    NetworkInspection            = $Defender.NISEnabled

    SignatureVersion             = $Defender.AntivirusSignatureVersion
    SignatureLastUpdated         = $Defender.AntivirusSignatureLastUpdated
    SignaturesOutOfDate          = $Defender.DefenderSignaturesOutOfDate

    TamperProtection             = $Defender.IsTamperProtected

    MDEOnboardingState           = $MDE.OnboardingState
    MDESenseService              = $Sense.Status

    SecuritySoftware             = ($SecuritySoftware -join '; ')
} | Format-List