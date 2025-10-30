<# 
.SYNOPSIS
  Cleans Windows Update caches and repairs component store on Windows 10/11.
  Tested on Win10 22H2 (incl. AVD session hosts).

.NOTES
  Run from an elevated PowerShell session.
  A reboot is recommended after completion.
#>

# --- Helper: Require elevation ---
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator. Right-click PowerShell and choose 'Run as administrator'."
    exit 1
}

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$logPath   = "$env:SystemDrive\WindowsUpdateRepair-$timestamp.log"
Start-Transcript -Path $logPath -Append | Out-Null

Write-Host "=== Windows Update repair starting @ $timestamp ==="

# --- Services to manage ---
$svcNames = @(
  'wuauserv',   # Windows Update
  'bits',       # Background Intelligent Transfer Service
  'cryptsvc',   # Cryptographic Services
  'msiserver'   # Windows Installer (often helps when resetting)
)

function Stop-Services {
  foreach ($s in $svcNames) {
    try {
      Write-Host "Stopping service: $s"
      Stop-Service -Name $s -Force -ErrorAction SilentlyContinue
      # Wait a moment for service to stop
      (Get-Service $s -ErrorAction SilentlyContinue) | Where-Object {$_.Status -ne 'Stopped'} | ForEach-Object {
        $_.WaitForStatus('Stopped','00:00:15')
      }
    } catch {
      Write-Warning "Couldn't stop $s: $($_.Exception.Message)"
    }
  }
}

function Start-Services {
  foreach ($s in $svcNames) {
    try {
      Write-Host "Starting service: $s"
      Start-Service -Name $s -ErrorAction SilentlyContinue
    } catch {
      Write-Warning "Couldn't start $s: $($_.Exception.Message)"
    }
  }
}

# --- Stop services first ---
Stop-Services

# --- Reset BITS transfer queue (safe) ---
Write-Host "Resetting BITS job queue (if any)…"
try {
  bitsadmin /reset | Out-Null
} catch {
  Write-Warning "BITS reset reported: $($_.Exception.Message)"
}

# --- Rename (clear) SoftwareDistribution & Catroot2 ---
$sdPath  = Join-Path $env:windir 'SoftwareDistribution'
$cr2Path = Join-Path $env:windir 'System32\catroot2'

foreach ($p in @($sdPath,$cr2Path)) {
  try {
    if (Test-Path $p) {
      $backup = "$p.$timestamp.bak"
      Write-Host "Renaming $p -> $backup"
      Rename-Item -Path $p -NewName (Split-Path $backup -Leaf) -ErrorAction Stop
    } else {
      Write-Host "$p not found—skipping rename."
    }
  } catch {
    Write-Warning "Failed to rename $p: $($_.Exception.Message)"
  }
}

# --- Reset WinSock (helps when update fails with network/cert issues) ---
Write-Host "Resetting WinSock…"
try {
  netsh winsock reset | Out-Null
} catch {
  Write-Warning "WinSock reset reported: $($_.Exception.Message)"
}

# --- Start services back up before DISM/SFC (DISM can run either way, but this is fine) ---
Start-Services

# --- Component store repair ---
Write-Host "Running DISM /Online /Cleanup-Image /RestoreHealth (this can take a while)…"
$dism = Start-Process -FilePath dism.exe `
  -ArgumentList "/Online","/Cleanup-Image","/RestoreHealth" `
  -NoNewWindow -PassThru -Wait
if ($dism.ExitCode -ne 0) {
  Write-Warning "DISM returned exit code $($dism.ExitCode). Check $logPath for details."
}

# --- System file scan ---
Write-Host "Running SFC /Scannow…"
$sfc = Start-Process -FilePath sfc.exe -ArgumentList "/scannow" -NoNewWindow -PassThru -Wait
if ($sfc.ExitCode -ne 0) {
  Write-Warning "SFC returned exit code $($sfc.ExitCode). (0 is good; nonzero may indicate repairs)."
}

# --- Kick off a fresh update scan/download/install ---
Write-Host "Restarting update services…"
Start-Services

Write-Host "Triggering a fresh Windows Update scan…"
# UsoClient is present on Win10/11. Ignore errors if restricted.
$usoCmds = @("StartScan","StartDownload","StartInstall")
foreach ($cmd in $usoCmds) {
  try {
    Start-Process -FilePath "$env:windir\system32\UsoClient.exe" -ArgumentList $cmd -NoNewWindow -Wait
  } catch {
    Write-Warning "UsoClient $cmd failed/blocked: $($_.Exception.Message)"
  }
}

Write-Host "=== Repair complete. Log: $logPath ==="
Write-Host "A reboot is recommended. Press Y to reboot now, or any other key to skip."
$choice = Read-Host "[Y/N]"
if ($choice -match '^[Yy]$') {
  Restart-Computer -Force
}

Stop-Transcript | Out-Null
