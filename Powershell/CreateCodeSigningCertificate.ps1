###################################################################################
# NAME:   CreateCodeSigningCertificate.ps1
# AUTHOR: Niels van de Coevering
# DATE:   4 October 2020
#
# COMMENTS: This script will create a local code signing certificate on a Windows
# computer. Use this if you are using a Powershell script on a single machine.
# For using a signed Powershell script on multiple machines, a public code signing
# certificate is recommended.
###################################################################################

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
$certPath = "Cert:\CurrentUser\My"

# Run script code
$validUntil = (Get-Date).AddYears(5)
if ($replaceCert) {
    Get-ChildItem -Path "Cert:\LocalMachine\Root" -CodeSigningCert | Where-Object { $_.Subject -eq "CN=$certSubject" } | Remove-Item
}
$cert = New-SelfSignedCertificate -CertStoreLocation $certPath -Type CodeSigningCert -NotAfter $validUntil -Subject $certSubject
if ($null -ne $cert) {
    $passwd = ConvertTo-SecureString -String $pfxPassword -Force -AsPlainText
    $path = Join-Path -Path $certPath -ChildPath $cert.Thumbprint
    if (-not(Test-Path -Path $pfxExportPath)) { New-Item -ItemType Directory -Path $pfxExportPath -Force }
    $pfxPath = Join-Path -Path $pfxExportPath -ChildPath $pfxFileName
    Export-PfxCertificate -Cert $path -FilePath $pfxPath -Password $passwd | Out-Null
    Get-ChildItem -Path $certPath -CodeSigningCert | Where-Object { $_.Subject -eq "CN=$certSubject" } | Remove-Item
    Import-PfxCertificate -FilePath $pfxPath -CertStoreLocation "Cert:\LocalMachine\Root" -Password $passwd | Out-Null
    Get-ChildItem $pfxPath | Remove-Item
}
