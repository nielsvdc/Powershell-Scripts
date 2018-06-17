<#
    Use this script for Azure VM's with SQL Server to be able to use the temporary storage
    to hold the SQL Server tempdb files.
    Set the SQL Server services to Manual and create a scheduler task to run this script
    at the startup of the VM.
    Change the location of the tempdb files in SQL Server.
    Now the tempdb can make use of the SSD speed of the temporary storage.
#>

function Wait-ForServiceStatus($searchString, $status)
{
    # Get all services where DisplayName matches $searchString and loop through each of them.
    foreach($service in (Get-Service $searchString))
    {
        # Wait for the service to reach the $status or a maximum of 30 minutes
        $service.WaitForStatus($status, '00:30:00')
    }
}

try {
    $SQLService = 'MSSQLSERVER'
    $SQLAgentService = 'SQLSERVERAGENT'
    $SQLSisService = 'MsDtsServer130'

    $tempfolder='D:\SQLTEMP'
    if (!(Test-Path -path $tempfolder)) {
        New-Item -ItemType Directory -Path $tempfolder
    }
    Start-Service $SQLService
    Wait-ForServiceStatus $SQLService 'Running'

    Start-Service $SQLAgentService
    #Start-Service $SQLSisService
}
catch {
    Write-EventLog –LogName Application –Source 'SQL-Startup Script' –EntryType Error –EventID 1 –Message $_.Exception
}


<#
# Create scheduler task
$action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument '-File C:\SQL-Startup.ps1'
$trigger = New-ScheduledTaskTrigger -AtStartup -RandomDelay '00:01:00'
$principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName "SQL Startup" -Description "Start SQL Server services with use of the temporary storage."

# Move SQL Server tempdb files
alter database tempdb modify file (name= tempdev, filename='D:\SQLTEMP\tempdb.mdf')
go
alter database tempdb modify file (name= templog, filename='D:\SQLTEMP\templog.ldf')
go
#>
