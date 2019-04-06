$VmResourceGroupName = ''
$VmName = ''
$KeyVaultResourceGroupName = ''
$KeyVaultName = ''
$KeyVaultKeyName = ''
$AADApplicationID = ''
$AADApplicationSecret = ''

$KeyVault = Get-AzureRmKeyVault -VaultName $KeyVaultName -ResourceGroupName $KeyVaultResourceGroupName

$DiskEncryptionKeyVaultURI  = $KeyVault.VaultUri
$DiskKeyVaultResourceId = $KeyVault.ResourceId
$DiskkeyEncryptionKeyUrl = (Get-AzureKeyVaultKey -VaultName $KeyVaultName -Name $KeyVaultKeyName).Key.kid;

Set-AzureRmVMDiskEncryptionExtension -ResourceGroupName $VmResourceGroupName -VMName $VmName `
                                    -AadClientID $AADApplicationID `
                                    -AadClientSecret $AADApplicationSecret `
                                    -DiskEncryptionKeyVaultUrl $DiskEncryptionKeyVaultURI `
                                    -DiskEncryptionKeyVaultId $DiskKeyVaultResourceId `
                                    -KeyEncryptionKeyUrl $DiskkeyEncryptionKeyUrl `
                                    -KeyEncryptionKeyVaultId $DiskKeyVaultResourceId `
                                    -VolumeType All
