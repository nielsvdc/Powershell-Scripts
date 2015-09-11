###################################################################################
# NAME:   Get-MailboxStatistics.ps1
# AUTHOR: Niels van de Coevering
# DATE:   12 September 2015
#
# COMMENTS: Get a CSV statistics report of Exchange Server mailbox size and usage.
###################################################################################

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue

$outputfile = "D:\MailboxStats.csv" # Change this file for output report

###################################################################################

function ConvertToDateNL 
{
	Param (
		[String]$datetime
	)

    if ($datetime -ne "")
    {
        # Get position of first space, which splits date and time
        $posfirstspace = $datetime.IndexOf(" ")
        # Get date string
        $tempdate = $datetime.SubString(0, $posfirstspace)
        # Split date string in mm dd yyyy
        $d = $tempdate.Split("/")
        # Rearrange date parts in dd-mm-yyyy format
        $date = $d[1] + "-" + $d[0] + "-" + $d[2].SubString(0,4)
    
        # Get time string. No need to change
        $time = $datetime.SubString($posfirstspace).Trim()
        
        # Return new date and time strings
        Write-Output $date $time
    }
}

# Send message to command box
Write-Host "Please wait while creating mailbox statistics report $outputfile..."

# Get mailbox information and output it to a temp file
$Mailboxes = Get-Mailbox -ResultSize Unlimited | select DisplayName, Alias, PrimarySmtpAddress,
@{name='IssuewarningQuota';expression={if ($_.IssueWarningQuota -match "UNLIMITED") {"-1"} else {$_.IssueWarningQuota.value.ToMB() }}},
@{name='ProhibitSendQuota';expression={if ($_.ProhibitSendQuota -match "UNLIMITED") {"-1"} else {$_.ProhibitSendQuota.value.ToMB() }}},
@{name='ProhibitSendReceiveQuota';expression={if ($_.ProhibitSendReceiveQuota -match "UNLIMITED") {"-1"} else {$_.ProhibitSendReceiveQuota.value.ToMB() }}},
WhenCreated

# Output columnheader line for CSV file
"DisplayName,Alias,PrimarySmtpAddress,DatabaseName,MailboxSizeMB,ItemCount,IssueWarningQuotaMB,ProhibitSendQuotaMB,ProhibitSendReceiveQuotaMB,CreationDate,LastLogonTime,LastLogoffTime,isActive" | out-file $outputfile

# Get mailbox statistics for each mailbox
foreach($Mailbox in $Mailboxes)
{
    # Get mailbox statistics
    $MailboxStats =  Get-MailboxStatistics $Mailbox.Alias | select DatabaseName,TotalItemSize,Itemcount,LastLogoffTime,LastLogonTime
    
    #Convert TotalItemSize to MB value
    $L = "{0:N0}" -f $MailboxStats.TotalItemSize.value.ToMB()
    # Remove comma as thousands seperator from size value
    $Size = $L.Replace(",", "")

    # Convert dates to format "dd-mm-yyyy hh:mm:ss" for Excel
    $LastLogoffTime = ConvertToDateNL $MailboxStats.LastLogoffTime
    $LastLogonTime = ConvertToDateNL $MailboxStats.LastLogonTime
    $WhenCreated = ConvertToDateNL $Mailbox.WhenCreated
    
    # Check in AD if mailbox account is enabled
    $temp = $Mailbox.PrimarySmtpAddress
    $adobjroot = "[adsi]"
    $objdisabsearcher = New-Object System.DirectoryServices.DirectorySearcher($adobjroot)
    $objdisabsearcher.Filter = "(&(objectCategory=Person)(objectClass=user)(mail=$temp)(userAccountControl:1.2.840.113556.1.4.803:=2))"
    $resultdisabaccn = $objdisabsearcher.FindOne() | select path

    if ($resultdisabaccn.Path) { $actStatus = "Disabled" }
    else { $actStatus = "Active" }
 
    # Create comma seperated output line
    $out = $Mailbox.Displayname  + "," + 
           $Mailbox.Alias  + "," + 
           $Mailbox.PrimarySmtpAddress + "," + 
           $MailboxStats.DatabaseName + "," + 
           $Size + "," + 
           $MailboxStats.ItemCount + "," + 
           $Mailbox.IssuewarningQuota + "," + 
           $Mailbox.ProhibitSendQuota + "," + 
           $Mailbox.ProhibitSendReceiveQuota + "," + 
           $WhenCreated + "," + 
           $LastLogonTime + "," + 
           $LastLogoffTime + "," + 
           $actStatus
    
    # Output line to file
    $out | Out-File $outputfile -Append
}
