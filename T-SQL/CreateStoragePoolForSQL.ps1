$dataStoragePoolName = 'SQL Data disks pool'
$dataVirtualDiskName = 'SQL Data disk'
$dataDiskLabelName = 'Data'
$logStoragePoolName = 'SQL Log disks pool'
$logVirtualDiskName = 'SQL Log disk'
$logDiskLabelName = 'Log'
$numberOfDataDisks = 8
$numberOfLogDisks = 0

$ConfirmPreference = 'None' # Confirm:$false is not working for Format-Volume. Setting this global variable to 'none' will fix this

function CreateNewStoragePool {
    param(
        [Object[]] $PhysicalDisks,
        [string] $StoragePoolName,
        [string] $VirtualDiskName,
        [string] $DiskLabelName
    )
    
    # Disable autodetect hardware notifications. Stops showing confirm messagebox for formatting volume
    Stop-Service -Name ShellHWDetection

    New-StoragePool -FriendlyName $StoragePoolName -StorageSubsystemFriendlyName "Windows Storage*" -PhysicalDisks $PhysicalDisks -ResiliencySettingNameDefault Mirror -ProvisioningTypeDefault Fixed -Verbose `
        |New-VirtualDisk -FriendlyName $VirtualDiskName -ResiliencySettingName Simple -ProvisioningType Fixed -UseMaximumSize `
        |Initialize-Disk -PassThru `
        |New-Partition -AssignDriveLetter -UseMaximumSize `
        |Format-Volume -FileSystem NTFS -AllocationUnitSize 65536 -NewFileSystemLabel $DiskLabelName -Confirm:$false

    # Enable autodetect hardware notifications
    Start-Service -Name ShellHWDetection
}

# Create data disk pool
$dataDisks = Get-PhysicalDisk -CanPool $true|Sort-Object DeviceId|Select-Object -First $numberOfDataDisks
if ($dataDisks -ne $null) {
    CreateNewStoragePool -PhysicalDisks $dataDisks -StoragePoolName $dataStoragePoolName -VirtualDiskName $dataVirtualDiskName -DiskLabelName $dataDiskLabelName
}

# Create log disk pool
$logDisks = Get-PhysicalDisk -CanPool $true|Sort-Object DeviceId|Select-Object -First $numberOfLogDisks
if ($logDisks -ne $null) {
    CreateNewStoragePool -PhysicalDisks $logDisks -StoragePoolName $logStoragePoolName -VirtualDiskName $logVirtualDiskName -DiskLabelName $logDiskLabelName
}

$dataDisks = $null
$logDisks = $null
