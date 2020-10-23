<###################################################################################
# NAME:   CreateCodeSigningCertificate.ps1
# AUTHOR: Niels van de Coevering
# DATE:   4 October 2020
#
# COMMENTS: This script will create a local code signing certificate on a Windows
# computer. Use this if you are using a Powershell script on a single machine.
# For using a signed Powershell script on multiple machines, a public code signing
# certificate is recommended.
###################################################################################>
#Requires -Version 5.1

# Check for code signing certificate for local machine saved in Trusted Root Certification Authorities
#Get-ChildItem -Path Cert:\LocalMachine\Root -CodeSigningCert

### Configure variables ###
$pfxExportPath = "C:\TEMP\Certificate\"
$pfxFileName = "powershellcert.pfx"
$pfxPassword = "12345" # Password is temporary, because the pfx file is deleted.
$replaceCert = $true
###########################

# Leave this as-is, if you don't know what you are doing
$certSubject = "Localhost code signing"
$certMyPath = "Cert:\LocalMachine\My"

# Run script code
$validUntil = (Get-Date).AddYears(5)
if ($replaceCert) {
    # When specifief, remove existing certificate
    Get-ChildItem -Path "Cert:\LocalMachine\Root" -CodeSigningCert | Where-Object { $_.Subject -eq "CN=$certSubject" } | Remove-Item
    Get-ChildItem -Path "Cert:\LocalMachine\TrustedPublisher" -CodeSigningCert | Where-Object { $_.Subject -eq "CN=$certSubject" } | Remove-Item
}
$cert = New-SelfSignedCertificate -CertStoreLocation $certMyPath -Type CodeSigningCert -NotAfter $validUntil -Subject $certSubject
if ($null -ne $cert) {
    # Create a secure password string
    $passwd = ConvertTo-SecureString -String $pfxPassword -Force -AsPlainText
    # Create full path to certificate in the personal certificate folder
    $path = Join-Path -Path $certMyPath -ChildPath $cert.Thumbprint
    # When not exist, create the export folder
    if (-not(Test-Path -Path $pfxExportPath)) { New-Item -ItemType Directory -Path $pfxExportPath -Force }
    # Create full path for certificate file export
    $pfxPath = Join-Path -Path $pfxExportPath -ChildPath $pfxFileName

    # Export certificate from the personal certificate folder
    Export-PfxCertificate -Cert $path -FilePath $pfxPath -Password $passwd | Out-Null
    # Import certificate into the Trusted Root Certification Authorities folder
    Import-PfxCertificate -FilePath $pfxPath -CertStoreLocation "Cert:\LocalMachine\Root" -Password $passwd | Out-Null
    # Import certificate into the Trusted Publishers folder
    Import-PfxCertificate -FilePath $pfxPath -CertStoreLocation "Cert:\LocalMachine\TrustedPublisher" -Password $passwd | Out-Null

    # Remove the certificate from the personal certificate folder and remove the pfx file
    Get-ChildItem -Path $certMyPath -CodeSigningCert | Where-Object { $_.Subject -eq "CN=$certSubject" } | Remove-Item
    Get-ChildItem $pfxPath | Remove-Item
}
