[CmdletBinding()]
param (
    [string]$VmFilter = "*MasterImage*",
    [int]$SysPrepTimeoutMinutes = 50   # <- change this if you want a different timeout
)

function Get-VMPowerstate {
    param (
        [string]$VMName,
        [string]$VMRG
    )
    $VMStatus = Get-AzVM -Name $VMName -ResourceGroupName $VMRG -Status
    foreach ($Status in $VMStatus.Statuses) {
        if ($Status.Code -like "PowerState*") {
            return $Status.DisplayStatus
        }
    }
    return "Unknown Power State"
}

function Prompt-Continue {
    Write-Host
    $response = Read-Host "Do you wish to proceed? (yes/no)"
    if ($response -notin @("yes", "y")) {
        Write-Host
        Write-Host "Exiting script."
        exit
    }
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

$index = 1
$filteredVMs | ForEach-Object {
    Write-Host ("      {0}. {1}" -f $index, $_.Name)
    $index++
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

$vnetList = Get-AzVirtualNetwork | Where-Object { $_.Location -eq $($selectedVM.Location) }
$vnet = $vnetList | Where-Object { $_.Name -eq $((($vmnicinfo.IpConfigurations.subnet.id).Split('/'))[-3]) }

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " You selected: " -NoNewline
Write-Host "$($selectedVM.Name)" -ForegroundColor Cyan
Write-Host "      Resource Group: $($selectedVM.ResourceGroupName)"
Write-Host "      Location: $($selectedVM.Location)"
Write-Host "      vNet Resource Group: $($vnet.ResourceGroupName)"
Write-Host "      Virtual Network: $((($vmnicinfo.IpConfigurations.subnet.id).Split('/'))[-3])"
Write-Host "      Subnet: $((($vmnicinfo.IpConfigurations.subnet.id).Split('/'))[-1])"

## Set variables

$ResourceGroupName   = $selectedVM.ResourceGroupName
$ImageResourceGroup  = $selectedVM.ResourceGroupName  # will be overwritten with gallery RG below
$Location            = $selectedVM.Location
$vNetResourceGroupName = $vnet.ResourceGroupName
$vNetName            = $((($vmnicinfo.IpConfigurations.subnet.id).Split('/'))[-3])
$SubnetName          = $((($vmnicinfo.IpConfigurations.subnet.id).Split('/'))[-1])

$AVDMasterVM         = $selectedVM.Name
$PrepVmName          = "${AVDMasterVM}_Prep"
$Now                 = Get-Date -UFormat "%Y%m%d_%H%M"
$snapshotName        = "${AVDMasterVM}_Pre-Sysprep"
$securitytype        = "TrustedLaunch"
$secureboot          = $true
$vtpm                = $true

##-------------------------------------------------------------------##
## Select Azure Compute Gallery
##-------------------------------------------------------------------##

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Select from one of the available galleries: "

$galleries = Get-AzGallery

if (-not $galleries) {
    Write-Host "No Azure Compute Galleries found in subscription." -ForegroundColor Red
    exit 1
}

$index = 1
$galleries | ForEach-Object {
    Write-Host ("      {0}. {1} (RG: {2})" -f $index, $_.Name, $_.ResourceGroupName)
    $index++
}

$gallerySelection = Read-Host "Select a gallery by number"
if (-not ($gallerySelection -as [int]) -or $gallerySelection -lt 1 -or $gallerySelection -gt $galleries.Count) {
    Write-Host "Invalid selection." -ForegroundColor Red
    exit 1
}

$selectedGallery   = $galleries[$gallerySelection - 1]
$galleryName       = $selectedGallery.Name
$ImageResourceGroup = $selectedGallery.ResourceGroupName   # <- fix: use gallery RG

##-------------------------------------------------------------------##
## Select Image Definition
##-------------------------------------------------------------------##

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Select from one of the available image definitions: "

$imageDefinitions = Get-AzGalleryImageDefinition -ResourceGroupName $selectedGallery.ResourceGroupName -GalleryName $selectedGallery.Name

if (-not $imageDefinitions) {
    Write-Host "No image definitions found in gallery '$($selectedGallery.Name)'." -ForegroundColor Red
    exit 1
}

$index = 1
$imageDefinitions | ForEach-Object {
    Write-Host ("      {0}. {1}" -f $index, $_.Name)
    $index++
}

$imageSelection = Read-Host "Select an image definition by number"
if (-not ($imageSelection -as [int]) -or $imageSelection -lt 1 -or $imageSelection -gt $imageDefinitions.Count) {
    Write-Host "Invalid selection." -ForegroundColor Red
    exit 1
}

$selectedImageDefinition   = $imageDefinitions[$imageSelection - 1]
$galleryImageDefinitionName = $selectedImageDefinition.Name

##-------------------------------------------------------------------##
## Confirm
##-------------------------------------------------------------------##

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Confirm the following details are correct: "
Write-Host "    Master Image: " -NoNewline ; Write-Host "$($selectedVM.Name)" -ForegroundColor Cyan
Write-Host "    Azure Compute Gallery: " -NoNewline ; Write-Host "$($selectedGallery.Name)" -ForegroundColor Cyan
Write-Host "    Image Definition: " -NoNewline ; Write-Host "$($selectedImageDefinition.Name)" -ForegroundColor Cyan

Prompt-Continue

##-------------------------------------------------------------------##
## Begin version creation process
##-------------------------------------------------------------------##

# Stop Master VM
Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Stopping and Deallocating AVD Master " -NoNewline
Write-Host $AVDMasterVM -ForegroundColor Cyan
Stop-AzVM -ResourceGroupName $ResourceGroupName -Name $AVDMasterVM -Force | Out-Null

# Get VM Info
$vm = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $AVDMasterVM

# Create Snapshot
$snapshotConfig = New-AzSnapshotConfig `
    -SourceUri $vm.StorageProfile.OsDisk.ManagedDisk.Id `
    -Location  $Location `
    -CreateOption Copy

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Taking Snapshot Backup of AVD Master called " -NoNewline
Write-Host $snapshotName -ForegroundColor Cyan

New-AzSnapshot -Snapshot $snapshotConfig -SnapshotName $snapshotName -ResourceGroupName $ResourceGroupName | Out-Null
$snapshot = Get-AzSnapshot -ResourceGroupName $ResourceGroupName -SnapshotName $snapshotName

# Create Prep VM from Snapshot
Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Creating Preparation VM " -NoNewline
Write-Host $PrepVmName -ForegroundColor Cyan

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Creating OS Disk from Snapshot" -ForegroundColor DarkYellow

$diskconfig = New-AzDiskConfig -Location $Location -SourceResourceId $snapshot.Id -CreateOption Copy
$DiskName   = "${PrepVmName}_OSdisk"
$newdisk    = New-AzDisk -Disk $diskconfig -ResourceGroupName $ResourceGroupName -DiskName $DiskName

$vmconfig = New-AzVMConfig -VMName $PrepVmName -VMSize Standard_D4ads_v5
$vmconfig = Set-AzVMOSDisk -VM $vmconfig -ManagedDiskId $newdisk.Id -CreateOption Attach -Windows
$vmconfig = Set-AzVMSecurityProfile -VM $vmconfig -SecurityType $securityType
$vmconfig = Set-AzVmUefi -VM $vmconfig -EnableVtpm $vtpm -EnableSecureBoot $secureboot

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Creating NIC" -ForegroundColor DarkYellow

$vnet   = Get-AzVirtualNetwork -Name $vNetName -ResourceGroupName $vNetResourceGroupName
$subnet = Get-AzVirtualNetworkSubnetConfig -Name $SubnetName -VirtualNetwork $vnet

$nicName = "${PrepVmName}-nic"
$vmnic   = New-AzNetworkInterface -Name $nicName -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $subnet.Id
$vmconfig = Add-AzVMNetworkInterface -VM $vmconfig -Id $vmnic.Id

$vmconfig = Set-AzVMBootDiagnostic -VM $vmconfig -Disable

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Deploying Preparation VM" -ForegroundColor Yellow
New-AzVM -VM $vmconfig -ResourceGroupName $ResourceGroupName -Location $Location | Out-Null

# Remove BGInfo Extension (if installed)
try {
    $VMExtensions = Get-AzVMExtension -ResourceGroupName $ResourceGroupName -VMName $PrepVmName -Name "BGInfo" -ErrorAction Stop
    foreach ($ext in $VMExtensions) {
        Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
        Write-Host "   Removing the BGInfo Extension" -ForegroundColor DarkYellow
        Remove-AzVMExtension -ResourceGroupName $ResourceGroupName -VMName $PrepVmName -Name "BGInfo" -Force | Out-Null
    }
}
catch {
    # BGInfo not present – ignore
}

# Run Sealing script
Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Running Sealing script on " -NoNewline
Write-Host "$PrepVMName" -ForeGroundColor Cyan

Set-AzVMExtension -ExtensionName "AVD_BIS-F" `
    -Location $Location `
    -ResourceGroupName $ResourceGroupName `
    -VMName $PrepVMName `
    -Publisher Microsoft.Compute `
    -ExtensionType CustomScriptExtension `
    -TypeHandlerVersion 1.8 `
    -SettingString '{"commandToExecute":"powershell C:\\Apps\\Tools\\Scripts\\Stage2\\Stage2_SysPrep.ps1"}' | Out-Null

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Removing Custom Script Extension" -ForegroundColor DarkYellow
Remove-AzVMExtension -ResourceGroupName $ResourceGroupName -VMName $PrepVMName -Name "AVD_BIS-F" -Force | Out-Null

# Wait for SysPrep to shut down the VM
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Waiting for BIS-F/SysPrep to Shut Down Preparation VM" -ForegroundColor DarkYellow
Write-Host "            " -NoNewline

# Initial grace period (same as your original: 15 * 15s ≈ 4 min)
for ($i = 1; $i -lt 16; $i++) {
    Start-Sleep -Seconds 15
    Write-Host "." -ForegroundColor DarkYellow -NoNewline
}

# Poll until VM stopped or timeout
$x = 0
$maxLoops = [int](($SysPrepTimeoutMinutes * 60) / 10)   # 10-second polls

do {
    Start-Sleep -Seconds 10
    $VMStatus = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $PrepVMName -Status
    Write-Host "." -ForegroundColor DarkYellow -NoNewline

    $powerState = ($VMStatus.Statuses | Where-Object { $_.Code -like "PowerState*" }).DisplayStatus
    $x++

    if ($x -gt $maxLoops) {
        Write-Host
        Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
        Write-Host "   ERROR:" -ForegroundColor Red -NoNewLine
        Write-Host " SysPrep has been running for too long (>${SysPrepTimeoutMinutes} minutes)"
        Write-Host "               Check the SysPrep logs on " -NoNewline
        Write-Host "'$PrepVMName'" -ForeGroundColor Cyan -NoNewline
        Write-Host " in C:\Windows\System32\Sysprep\Panther"
        exit 1
    }
}
while ($powerState -ne "VM stopped")

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Shut Down complete" -ForegroundColor DarkYellow

## Creating the Disk Image

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Stopping AVD Prep Image VM " -NoNewline
Write-Host $PrepVMName -ForegroundColor Cyan
Stop-AzVM -ResourceGroupName $ResourceGroupName -Name $PrepVMName -Force | Out-Null

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Generalise Preparation VM" -ForegroundColor DarkYellow
Set-AzVm -ResourceGroupName $ResourceGroupName -Name $PrepVMName -Generalized | Out-Null

$PrepVM = Get-AzVM -Name $PrepVMName -ResourceGroupName $ResourceGroupName

## Shared Image Gallery

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Getting Details of Shared Image Gallery " -NoNewline
Write-Host $galleryName -ForegroundColor Green
$GetSharedImageGallery = Get-AzGallery -Name $galleryName -ResourceGroupName $ImageResourceGroup -ErrorAction Stop

Write-Host
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host " Getting Details of Shared Image Gallery Definition " -NoNewline
Write-Host $galleryImageDefinitionName -ForegroundColor Green
$GetSharedImageGalleryDefinition = Get-AzGalleryImageDefinition `
    -ResourceGroupName $ImageResourceGroup `
    -GalleryName $galleryName `
    -Name $galleryImageDefinitionName `
    -ErrorAction Stop

## Add Image to Shared Image Gallery

$Version = Get-Date -Format "yyyyMM.dd.HHmm"
Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Version Number is: " -NoNewLine
Write-Host $Version -ForegroundColor DarkYellow

# Replication regions (kept same as your original)
$region1 = @{ Name = 'uk south'; ReplicaCount = 1 }
$region2 = @{ Name = 'uk west';  ReplicaCount = 1 }
$targetRegions = @($region1, $region2)

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Adding Image to Shared Image Gallery " -NoNewline
Write-Host "(via AsJob)" -ForegroundColor DarkYellow

$job = New-AzGalleryImageVersion `
    -GalleryImageDefinitionName $GetSharedImageGalleryDefinition.Name `
    -GalleryImageVersionName $Version `
    -GalleryName $galleryName `
    -StorageAccountType "Standard_LRS" `
    -ResourceGroupName $ImageResourceGroup `
    -Location $Location `
    -TargetRegion $targetRegions `
    -SourceImageVMId $PrepVM.Id `
    -PublishingProfileEndOfLifeDate $((Get-Date).ToUniversalTime().AddMonths(12).ToString('yyyy-MM-dd')) `
    -AsJob

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Check Status of Deployment in Azure Portal," -NoNewline
Write-Host " This can take up to 30 minutes to complete" -ForegroundColor DarkYellow

Start-Sleep -Seconds 30

do {
    $Status = Get-AzGalleryImageVersion `
        -ResourceGroupName $ImageResourceGroup `
        -GalleryImageDefinitionName $GalleryImageDefinitionName `
        -GalleryName $galleryName `
        | Where-Object { $_.Name -eq $Version }

    Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
    Write-Host "   Current Status is: " -NoNewline

    if (-not $Status) {
        Write-Host "Not yet visible" -ForegroundColor DarkCyan
        Start-Sleep -Seconds 60
        continue
    }

    if ($Status.ProvisioningState -ne 'Succeeded') {
        Write-Host $Status.ProvisioningState -ForegroundColor DarkCyan
        Start-Sleep -Seconds 60
    }
}
while ($Status.ProvisioningState -ne 'Succeeded')

## Cleanup Prep VM and snapshot

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Remove Preparation VM from Azure" -ForegroundColor DarkYellow

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "     Remove VM" -ForegroundColor DarkYellow
$PrepVMNameResource = Get-AzVM -Name $PrepVmName -ResourceGroupName $ResourceGroupName
$prepNic = Get-AzNetworkInterface -ResourceID $PrepVMNameResource.NetworkProfile.NetworkInterfaces.Id
Remove-AzVM -ResourceGroupName $ResourceGroupName -Name $PrepVMName -Force | Out-Null

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "     Remove OS Disk" -ForegroundColor DarkYellow
Remove-AzDisk -ResourceGroupName $ResourceGroupName -DiskName $PrepVMNameResource.StorageProfile.OsDisk.Name -Force | Out-Null

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "     Remove NIC" -ForegroundColor DarkYellow
Remove-AzNetworkInterface -ResourceGroupName $ResourceGroupName -Name $prepNic.Name -Force | Out-Null

Write-Host (Get-Date -Format HH:mm:ss:) -ForegroundColor Gray -NoNewLine
Write-Host "   Remove Snapshot" -ForegroundColor DarkYellow
Remove-AzSnapshot -SnapshotName $snapshotName -ResourceGroupName $ResourceGroupName -Force | Out-Null

Write-Host 'Succeeded' -ForegroundColor Cyan
