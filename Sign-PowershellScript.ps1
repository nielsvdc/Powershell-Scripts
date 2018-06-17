## sign-script.ps1
## Sign a powershell script with a Thawte certificate and 
## timestamp the signature
##
## usage: ./sign-script.ps1 c:\foo.ps1 

param([string] $File=$(throw "Please specify a script filepath.")) 

$certFriendlyName = "Thawte Code Signing"
$cert = Get-ChildItem cert:\CurrentUser\My -codesigning #| where -Filter {$_.FriendlyName -eq $certFriendlyName}

# https://www.thawte.com/ssl-digital-certificates/technical-  
#   support/code/msauth.html#timestampau
# We thank VeriSign for allowing public use of their timestamping server.
# Add the following to the signcode command line: 
# -t http://timestamp.verisign.com/scripts/timstamp.dll 
$timeStampURL = "http://timestamp.verisign.com/scripts/timstamp.dll"

if($cert) {
    Set-AuthenticodeSignature -filepath $file -cert $cert -IncludeChain All -TimeStampServer $timeStampURL
}
else {
    throw "Did not find certificate with friendly name of `"$certFriendlyName`""
}
