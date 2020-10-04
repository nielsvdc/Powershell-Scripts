<###################################################################################
# NAME:   SignPowershellScript.ps1
# AUTHOR: Niels van de Coevering
# DATE:   4 October 2020
#
# COMMENTS: This script will use a code signing certificate found in on a Windows
# computer to sign a specified Powershell script.
###################################################################################>

# Sign a Powershell script
$scriptPath = "C:\TEMP\"
$scriptName = "MyScript.ps1"

# Leave this as-is, if you don't know what you are doing
$certPath = "Cert:\LocalMachine\Root"

$filePath = Join-Path -Path $scriptPath -ChildPath $scriptName
$cert = Get-ChildItem -Path $certPath -CodeSigningCert
if (-not($cert)) {
    Write-Host "Self-signed code signing certificate not found in: $certPath.`nLocation should be 'Cert:\LocalMachine\Root'. Please create a self-signed certificate first." -BackgroundColor Red
}
else {
    $result = Set-AuthenticodeSignature -FilePath $filePath -Certificate $cert
    $result
    if ($result.Status -ne 'Valid') {
        Write-Host "Failed signing Powershell script: $filePath" -BackgroundColor Red
    }
}
