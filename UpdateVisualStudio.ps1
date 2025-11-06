$vsInstaller = "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vs_installer.exe"

if (Test-Path $vsInstaller) {
    & $vsInstaller update --all --quiet --norestart
    Write-Host "Visual Studio update started (quiet mode, no restart)."
} else {
    Write-Host "Visual Studio Installer not found!"
}
