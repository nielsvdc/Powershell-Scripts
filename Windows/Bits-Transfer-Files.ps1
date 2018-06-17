<#
  Use the Bits Transfer service in Windows to transfer files from a source to a destination
  when the network connection is unstable. The Bits Transfer service will continue the transfer
  even of the network connection disconnected for a moment.
  This script will show a progressbar in Powershell.
#>

$source = ''
$destinationFolder = ''


$start = Get-Date
$job = Start-BitsTransfer -Source $source -Destination $destinationFolder -DisplayName 'Transfer bak file' -Asynchronous 
$destination = Join-Path $destinationFolder ($source | Split-Path -Leaf)

while (($job.JobState -eq 'Transferring') -or ($job.JobState -eq 'Connecting')){ 
        filter Get-FileSize {
	    "{0:N2} {1}" -f $(
	    if ($_ -lt 1kb) { $_, 'Bytes' }
	    elseif ($_ -lt 1mb) { ($_/1kb), 'KB' }
	    elseif ($_ -lt 1gb) { ($_/1mb), 'MB' }
	    elseif ($_ -lt 1tb) { ($_/1gb), 'GB' }
	    elseif ($_ -lt 1pb) { ($_/1tb), 'TB' }
	    else { ($_/1pb), 'PB' }
	    )
    }
    $elapsed = ((Get-Date) - $start)
    #calculate average speed in Mbps
    $averageSpeed = ($job.BytesTransferred * 8 / 1MB) / $elapsed.TotalSeconds
    $elapsed = $elapsed.ToString('hh\:mm\:ss')
    #calculate remaining time considering average speed
    $remainingSeconds = ($job.BytesTotal - $job.BytesTransferred) * 8 / 1MB / $averageSpeed
    $receivedSize = $job.BytesTransferred | Get-FileSize
    $totalSize = $job.BytesTotal | Get-FileSize 
    $progressPercentage = [int]($job.BytesTransferred / $job.BytesTotal * 100)
    $retries = $job.TransientErrorCount
    if ($remainingSeconds -as [int]){
        Write-Progress -Activity (" $source, speed: {0:N0} Mbps" -f $averageSpeed)  -Status ("{0} of {1} ({2}% in {3}), retries: {4}" -f $receivedSize, $totalSize, $progressPercentage, $elapsed, $retries) -SecondsRemaining $remainingSeconds -PercentComplete $progressPercentage
    }
} 
if ($includeStats.IsPresent){
    ([PSCustomObject]@{Name=$MyInvocation.MyCommand;TotalSize=$totalSize;Time=$elapsed}) | Out-Host
}

Write-Progress -Activity (" $source {0:N2} Mbps" -f $averageSpeed) -Status 'Done' -Completed
Switch($job.JobState){
	'Transferred' {
        Complete-BitsTransfer -BitsJob $job
        Get-Item $destination | Unblock-File
    }
	'Error' {
        Write-Warning "Download of $source failed" 
        $job | Format-List
    } 
}
