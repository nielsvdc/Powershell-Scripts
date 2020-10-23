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

# Create variable for full path to Powershell script
$filePath = Join-Path -Path $scriptPath -ChildPath $scriptName
# If exist, get a code signing certificate
$cert = Get-ChildItem -Path $certPath -CodeSigningCert

# Check of a code signing certificate exists
if (-not($cert)) {
    # Show error message when code signing certificate does not exist
    Write-Host "Self-signed code signing certificate not found in: $certPath.`nLocation should be 'Cert:\LocalMachine\Root'. Please create a self-signed certificate first." -BackgroundColor Red
}
else {
    # Sign the Powershell script using the code signing certificate
    $result = Set-AuthenticodeSignature -FilePath $filePath -Certificate $cert

    # Show result
    $result
    if ($result.Status -ne 'Valid') {
        Write-Host "Failed signing Powershell script: $filePath" -BackgroundColor Red
    }
}
