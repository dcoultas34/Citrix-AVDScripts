##-------------------------------------------------------------------------##
# Running the following commands to log into Azure:
#
#Connect-AzAccount 
#set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#Install-module -Name AZ -AllowClobber -Force -Scope currentuser
##-------------------------------------------------------------------------##





[CmdletBinding()]
param (
    [string]$VmFilter = "*MasterIm*",
    [int]$SysPrepTimeoutMinutes = 60,
    [int]$ImageBuildTimeoutMinutes = 90
)

function Prompt-Continue {
    Write-Host
    $response = Read-Host "Do you wish to proceed? (yes/no)"
    if ($response -notin @("yes","y")) {
        Write-Host
        Write-Host "Exiting script."
        exit 0
    }
}

function Get-VMPowerstate {
    param(
        [string]$VMName,
        [string]$VMRG
    )
    $vmStatus = Get-AzVM -Name $VMName -ResourceGroupName $VMRG -Status
    ($vmStatus.Statuses | Where-Object { $_.Code -like "PowerState*" }).DisplayStatus
}

function Wait-VMToStop {
    param(
        [string]$VMName,
        [string]$VMRG,
        [int]$TimeoutMinutes = 60
    )
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    do {
        Start-Sleep -Seconds 10
        $state = Get-VMPowerstate -VMName $VMName -VMRG $VMRG
        Write-Host "." -NoNewline -ForegroundColor DarkYellow
        if ($state -in @("VM stopped","VM deallocated")) { Write-Host; return $true }
    } while ((Get-Date) -lt $deadline)

    Write-Host
    Write-Host "Timed out waiting for VM '$VMName' to stop/deallocate. Last state: $state" -ForegroundColor Red
    return $false
}

function Wait-VMExtensionCompleted {
    param(
        [string]$ResourceGroupName,
        [string]$VMName,
        [string]$ExtensionName,
        [int]$TimeoutMinutes = 60
    )
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)

    do {
        Start-Sleep -Seconds 10
        $ext = Get-AzVMExtension -ResourceGroupName $ResourceGroupName -VMName $VMName -Name $ExtensionName -Status -ErrorAction SilentlyContinue

        Write-Host "." -NoNewline -ForegroundColor DarkYellow

        if (-not $ext) { continue }

        $prov = $ext.ProvisioningState
        if ($prov -eq "Succeeded") { Write-Host; return $true }

        if ($prov -eq "Failed") {
            Write-Host
            Write-Host "Extension '$ExtensionName' FAILED." -ForegroundColor Red
            ($ext.Statuses | ForEach-Object { $_.Message } | Where-Object { $_ } | Select-Object -First 1) | ForEach-Object {
                Write-Host "Message: $_" -ForegroundColor Red
            }
            return $false
        }

    } while ((Get-Date) -lt $deadline)

    Write-Host
    Write-Host "Timed out waiting for extension '$ExtensionName' on '$VMName'." -ForegroundColor Red
    return $false
}

function Wait-GalleryImageVersion {
    param(
        [string]$GalleryName,
        [string]$GalleryImageDefinitionName,
        [string]$ResourceGroupName,
        [string]$Version,
        [int]$TimeoutMinutes = 90
    )
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)

    do {
        Start-Sleep -Seconds 30
        $status = Get-AzGalleryImageVersion `
            -ResourceGroupName $ResourceGroupName `
            -GalleryName $GalleryName `
            -GalleryImageDefinitionName $GalleryImageDefinitionName `
            -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq $Version }

        Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
        Write-Host " SIG version provisioning state: " -NoNewLine

        if (-not $status) {
            Write-Host "Not yet visible" -ForegroundColor DarkCyan
            continue
        }

        Write-Host $status.ProvisioningState -ForegroundColor DarkCyan

        if ($status.ProvisioningState -eq "Succeeded") { return $true }
        if ($status.ProvisioningState -eq "Failed")    { return $false }

    } while ((Get-Date) -lt $deadline)

    Write-Host "Timed out waiting for SIG version '$Version'." -ForegroundColor Red
    return $false
}

##-------------------------------------------------------------------##
## Select Master Image VM
##-------------------------------------------------------------------##

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Select from one of the available Master Images: "

$vms = Get-AzVM
$filteredVMs = $vms | Where-Object { $_.Name -like $VmFilter }

if (-not $filteredVMs) {
    Write-Host "No VMs found matching filter '$VmFilter'." -ForegroundColor Red
    exit 1
}

for ($i=0; $i -lt $filteredVMs.Count; $i++) {
    Write-Host ("      {0}. {1}" -f ($i+1), $filteredVMs[$i].Name)
}

$selection = Read-Host "Select a VM by number"
if (-not ($selection -as [int]) -or $selection -lt 1 -or $selection -gt $filteredVMs.Count) {
    Write-Host "Invalid selection." -ForegroundColor Red
    exit 1
}
$selectedVM = $filteredVMs[$selection - 1]

# NIC & vNet info
$vmnic = ($selectedVM.NetworkProfile.NetworkInterfaces.id).Split('/')[-1]
$vmnicinfo = Get-AzNetworkInterface -Name $vmnic

$vnetList = Get-AzVirtualNetwork | Where-Object { $_.Location -eq $selectedVM.Location }
$vnet = $vnetList | Where-Object { $_.Name -eq ((($vmnicinfo.IpConfigurations.subnet.id).Split('/'))[-3]) }

$ResourceGroupName     = $selectedVM.ResourceGroupName
$Location              = $selectedVM.Location
$vNetResourceGroupName = $vnet.ResourceGroupName
$vNetName              = ((($vmnicinfo.IpConfigurations.subnet.id).Split('/'))[-3])
$SubnetName            = ((($vmnicinfo.IpConfigurations.subnet.id).Split('/'))[-1])

$AVDMasterVM  = $selectedVM.Name
$PrepVmName   = "${AVDMasterVM}_Prep"
$snapshotName = "${AVDMasterVM}_Pre-Sysprep"

$securitytype = "TrustedLaunch"
$secureboot   = $true
$vtpm         = $true

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " You selected: " -NoNewline
Write-Host "$($selectedVM.Name)" -ForegroundColor Cyan
Write-Host "      Resource Group: $ResourceGroupName"
Write-Host "      Location: $Location"
Write-Host "      vNet Resource Group: $vNetResourceGroupName"
Write-Host "      Virtual Network: $vNetName"
Write-Host "      Subnet: $SubnetName"

##-------------------------------------------------------------------##
## Select Azure Compute Gallery + Definition
##-------------------------------------------------------------------##

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Select from one of the available galleries: "

$galleries = Get-AzGallery
if (-not $galleries) {
    Write-Host "No galleries found in subscription." -ForegroundColor Red
    exit 1
}

for ($i=0; $i -lt $galleries.Count; $i++) {
    Write-Host ("      {0}. {1} (RG: {2})" -f ($i+1), $galleries[$i].Name, $galleries[$i].ResourceGroupName)
}

$gallerySelection = Read-Host "Select a gallery by number"
if (-not ($gallerySelection -as [int]) -or $gallerySelection -lt 1 -or $gallerySelection -gt $galleries.Count) {
    Write-Host "Invalid selection." -ForegroundColor Red
    exit 1
}

$selectedGallery    = $galleries[$gallerySelection - 1]
$galleryName        = $selectedGallery.Name
$ImageResourceGroup = $selectedGallery.ResourceGroupName   # IMPORTANT FIX

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Select from one of the available image definitions: "

$imageDefinitions = Get-AzGalleryImageDefinition -ResourceGroupName $ImageResourceGroup -GalleryName $galleryName
if (-not $imageDefinitions) {
    Write-Host "No image definitions found in gallery '$galleryName'." -ForegroundColor Red
    exit 1
}

for ($i=0; $i -lt $imageDefinitions.Count; $i++) {
    Write-Host ("      {0}. {1}" -f ($i+1), $imageDefinitions[$i].Name)
}

$imageSelection = Read-Host "Select an image definition by number"
if (-not ($imageSelection -as [int]) -or $imageSelection -lt 1 -or $imageSelection -gt $imageDefinitions.Count) {
    Write-Host "Invalid selection." -ForegroundColor Red
    exit 1
}

$selectedImageDefinition    = $imageDefinitions[$imageSelection - 1]
$galleryImageDefinitionName = $selectedImageDefinition.Name

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Confirm the following details are correct: "
Write-Host "    Master Image: " -NoNewline ; Write-Host $AVDMasterVM -ForegroundColor Cyan
Write-Host "    Gallery: "      -NoNewline ; Write-Host $galleryName -ForegroundColor Cyan
Write-Host "    Definition: "   -NoNewline ; Write-Host $galleryImageDefinitionName -ForegroundColor Cyan
Write-Host "    Gallery RG: "   -NoNewline ; Write-Host $ImageResourceGroup -ForegroundColor Cyan

Prompt-Continue

##-------------------------------------------------------------------##
## Begin process
##-------------------------------------------------------------------##

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Stopping and deallocating master VM " -NoNewline
Write-Host $AVDMasterVM -ForegroundColor Cyan
Stop-AzVM -ResourceGroupName $ResourceGroupName -Name $AVDMasterVM -Force | Out-Null

$vm = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $AVDMasterVM

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Creating snapshot " -NoNewline
Write-Host $snapshotName -ForegroundColor Cyan

$snapshotConfig = New-AzSnapshotConfig -SourceUri $vm.StorageProfile.OsDisk.ManagedDisk.Id -Location $Location -CreateOption Copy
New-AzSnapshot -Snapshot $snapshotConfig -SnapshotName $snapshotName -ResourceGroupName $ResourceGroupName | Out-Null
$snapshot = Get-AzSnapshot -ResourceGroupName $ResourceGroupName -SnapshotName $snapshotName

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Creating prep VM " -NoNewline
Write-Host $PrepVmName -ForegroundColor Cyan

$diskconfig = New-AzDiskConfig -Location $Location -SourceResourceId $snapshot.Id -CreateOption Copy
$diskName   = "${PrepVmName}_OSdisk"
$newdisk    = New-AzDisk -Disk $diskconfig -ResourceGroupName $ResourceGroupName -DiskName $diskName

$vmconfig = New-AzVMConfig -VMName $PrepVmName -VMSize "Standard_D4ads_v5"
$vmconfig = Set-AzVMOSDisk -VM $vmconfig -ManagedDiskId $newdisk.Id -CreateOption Attach -Windows
$vmconfig = Set-AzVMSecurityProfile -VM $vmconfig -SecurityType $securitytype
$vmconfig = Set-AzVmUefi -VM $vmconfig -EnableVtpm $vtpm -EnableSecureBoot $secureboot

$vnet   = Get-AzVirtualNetwork -Name $vNetName -ResourceGroupName $vNetResourceGroupName
$subnet = Get-AzVirtualNetworkSubnetConfig -Name $SubnetName -VirtualNetwork $vnet
$nicName = "${PrepVmName}-nic"
$vmnic   = New-AzNetworkInterface -Name $nicName -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $subnet.Id

$vmconfig = Add-AzVMNetworkInterface -VM $vmconfig -Id $vmnic.Id
$vmconfig = Set-AzVMBootDiagnostic -VM $vmconfig -Disable
New-AzVM -VM $vmconfig -ResourceGroupName $ResourceGroupName -Location $Location | Out-Null

# Remove BGInfo if present
try {
    Get-AzVMExtension -ResourceGroupName $ResourceGroupName -VMName $PrepVmName -Name "BGInfo" -ErrorAction Stop | Out-Null
    Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
    Write-Host " Removing BGInfo extension" -ForegroundColor DarkYellow
    Remove-AzVMExtension -ResourceGroupName $ResourceGroupName -VMName $PrepVmName -Name "BGInfo" -Force | Out-Null
} catch { }

## Run sealing script via CustomScriptExtension (FIXED quoting with cmd /c)
Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Running sealing script via CustomScriptExtension on " -NoNewLine
Write-Host $PrepVmName -ForegroundColor Cyan

$sealScript = "C:\Apps\Tools\Scripts\Stage2\Stage2_SysPrep.ps1"
$command = "cmd /c powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$sealScript`""

Set-AzVMExtension `
    -ExtensionName "AVD_BIS-F" `
    -Location $Location `
    -ResourceGroupName $ResourceGroupName `
    -VMName $PrepVmName `
    -Publisher "Microsoft.Compute" `
    -ExtensionType "CustomScriptExtension" `
    -TypeHandlerVersion "1.10" `
    -SettingString (@{ commandToExecute = $command } | ConvertTo-Json -Compress) | Out-Null

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Waiting for extension to complete" -ForegroundColor DarkYellow -NoNewline
$ok = Wait-VMExtensionCompleted -ResourceGroupName $ResourceGroupName -VMName $PrepVmName -ExtensionName "AVD_BIS-F" -TimeoutMinutes $SysPrepTimeoutMinutes

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Removing Custom Script Extension" -ForegroundColor DarkYellow
Remove-AzVMExtension -ResourceGroupName $ResourceGroupName -VMName $PrepVmName -Name "AVD_BIS-F" -Force | Out-Null

if (-not $ok) {
    throw "Sealing script failed (CustomScriptExtension provisioning state Failed). Check C:\WindowsAzure\Logs\Plugins\Microsoft.Compute.CustomScriptExtension on $PrepVmName."
}

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Waiting for prep VM to shut down" -ForegroundColor DarkYellow -NoNewline
$stopped = Wait-VMToStop -VMName $PrepVmName -VMRG $ResourceGroupName -TimeoutMinutes $SysPrepTimeoutMinutes
if (-not $stopped) { throw "Prep VM did not stop/deallocate in time." }

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Prep VM shutdown complete" -ForegroundColor DarkYellow

## Generalize + Build Image Version
Stop-AzVM -ResourceGroupName $ResourceGroupName -Name $PrepVmName -Force | Out-Null
Set-AzVm -ResourceGroupName $ResourceGroupName -Name $PrepVmName -Generalized | Out-Null
$PrepVM = Get-AzVM -Name $PrepVmName -ResourceGroupName $ResourceGroupName

$version = Get-Date -Format "yyyyMM.dd.HHmm"
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Creating SIG version: " -NoNewLine
Write-Host $version -ForegroundColor DarkYellow

$targetRegions = @(
    @{ Name = 'uk south'; ReplicaCount = 1 },
    @{ Name = 'uk west';  ReplicaCount = 1 }
)

New-AzGalleryImageVersion `
    -GalleryImageDefinitionName $galleryImageDefinitionName `
    -GalleryImageVersionName $version `
    -GalleryName $galleryName `
    -ResourceGroupName $ImageResourceGroup `
    -Location $Location `
    -TargetRegion $targetRegions `
    -SourceImageVMId $PrepVM.Id `
    -StorageAccountType "Standard_LRS" `
    -PublishingProfileEndOfLifeDate $((Get-Date).ToUniversalTime().AddMonths(12).ToString('yyyy-MM-dd')) `
    -AsJob | Out-Null

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Waiting for SIG version to finish..." -ForegroundColor DarkYellow

$done = Wait-GalleryImageVersion `
    -GalleryName $galleryName `
    -GalleryImageDefinitionName $galleryImageDefinitionName `
    -ResourceGroupName $ImageResourceGroup `
    -Version $version `
    -TimeoutMinutes $ImageBuildTimeoutMinutes

if (-not $done) {
    throw "SIG image version provisioning did not succeed. Check the gallery image version '$version' in the portal."
}

Write-Host
Write-Host "SIG image version '$version' created successfully." -ForegroundColor Cyan

## Cleanup
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Cleaning up prep VM + snapshot" -ForegroundColor DarkYellow

$prepVmObj = Get-AzVM -Name $PrepVmName -ResourceGroupName $ResourceGroupName
$prepNic = Get-AzNetworkInterface -ResourceId $prepVmObj.NetworkProfile.NetworkInterfaces.Id

Remove-AzVM -ResourceGroupName $ResourceGroupName -Name $PrepVmName -Force | Out-Null
Remove-AzDisk -ResourceGroupName $ResourceGroupName -DiskName $prepVmObj.StorageProfile.OsDisk.Name -Force | Out-Null
Remove-AzNetworkInterface -ResourceGroupName $ResourceGroupName -Name $prepNic.Name -Force | Out-Null
Remove-AzSnapshot -SnapshotName $snapshotName -ResourceGroupName $ResourceGroupName -Force | Out-Null

Write-Host "Succeeded" -ForegroundColor Cyan
