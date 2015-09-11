###################################################################################
# NAME:   Update-ADUser-HomeDir.ps1
# AUTHOR: Niels van de Coevering
# DATE:   12 September 2015
#
# COMMENTS: This script will get all enabled AD user accounts and the HomeDirectory
# property and change the current path to a new path. Optionaly the script can be
# used to also move the old content to the new directory.
###################################################################################

$newHomeDir = "\\fs01\Users\";
$homeDrive = "J:";

###################################################################################

Get-ADUser -Filter {Enabled -eq $true} -Properties ScriptPath, HomeDrive, HomeDirectory | Select SamAccountName, HomeDirectory | 
    foreach {
        $samAccountName = $_.SamAccountName;
        $oldDir = $_.HomeDirectory;
        $newDir = $newHomeDir+$samAccountName;
        
        if ($oldDir -ne $newDir)
        {
            try {
                # Move old home dir content to new location
                #[System.IO.Directory]::Move($oldDir, $newDir);
                # Change profile home dir in AD
                Set-ADUser $_.SamAccountName -HomeDirectory $newDir -HomeDrive $homeDrive;
                
                Write-Host "Finished "$samAccountName;
            }
            catch {
                Write-Host "Failed to change "$samAccountName;
            }
        }
    }
