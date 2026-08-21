# Define the network share path and drive letter
#Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

#$sharePath = "\\SAAVDUKWG2V.File.core.windows.net\fslogix\Profiles_W11"
$sharePath = "\\saavduksg2v.file.core.windows.net\fslogix\Profiles"
$driveLetter = "N"

# Check if the drive is already mapped
$existingDrive = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue

# If the drive exists, unmount it first
if ($existingDrive) {
    Write-Host "Drive $driveLetter is already mapped. Unmounting first..." -ForegroundColor Yellow
    Remove-PSDrive -Name $driveLetter -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2  # Give some time before remounting
}

# Map the network share
Write-Host "Mapping network share to $driveLetter..." -ForegroundColor Cyan
New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $sharePath -Persist -ErrorAction Stop

# Wait for mapping to complete
Start-Sleep -Seconds 2

# Retrieve disk information
$diskInfo = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue

if ($diskInfo) {
    # Calculate total size, free space, and free space percentage
    $totalSizeGB = [math]::Round(($diskInfo.Used + $diskInfo.Free) / 1GB, 2)
    $freeSpaceGB = [math]::Round($diskInfo.Free / 1GB, 2)
    $freePercentage = [math]::Round(($diskInfo.Free / ($diskInfo.Used + $diskInfo.Free)) * 100, 2)

    # Display the results
    Write-Host "`n==== Network Share Drive Space Information ====" -ForegroundColor Green
    Write-Host ("{0,-10} {1,-12} {2,13} {3,-10}" -f "Drive", "Total GB", "Free GB", "% Free") -ForegroundColor White
    Write-Host ("{0,-10} {1,12} {2,13} {3,-10}" -f $driveLetter, $totalSizeGB, $freeSpaceGB, $freePercentage) -ForegroundColor White

    # Display a warning if free space is low
    if ($freePercentage -lt 10) {
        Write-Host "`n⚠️ WARNING: Low disk space! Only $freePercentage% free!" -ForegroundColor Red
    } else {
        Write-Host "`n✅ Sufficient space available: $freePercentage% free" -ForegroundColor Green
    }
} else {
    Write-Host "❌ Failed to retrieve disk information. Drive may not be mapped correctly." -ForegroundColor Red
}

# Unmount the drive
Write-Host "`nUnmounting network share..." -ForegroundColor Yellow
Remove-PSDrive -Name $driveLetter -Force -ErrorAction SilentlyContinue
Write-Host "Network share unmounted." -ForegroundColor Cyan
