Param(
  [string]$PkgFolder = (Get-Location).Path
)

# enable sideload
New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" -Name "AllowAllTrustedApps" -Type DWord -Value 1

# install appx/msix in deterministic order: appx (vclibs) then msix (ui.xaml) then bundle
$apps = Get-ChildItem -Path $PkgFolder -File | Sort-Object Extension

foreach($f in $apps) {
  Write-Host "Installing $($f.Name)..."
  try {
    Add-AppxPackage -Path $f.FullName -ForceApplicationShutdown -ErrorAction Stop
    Write-Host "Installed $($f.Name)"
  } catch {
    Write-Warning "Failed to install $($f.Name): $($_.Exception.Message)"
  }
}

Write-Host "Verify App Installer / winget..."
Get-AppxPackage -AllUsers Microsoft.DesktopAppInstaller | Format-Table Name, Version, PackageFullName -AutoSize
try { winget --version } catch { Write-Warning "winget not found after install." }
