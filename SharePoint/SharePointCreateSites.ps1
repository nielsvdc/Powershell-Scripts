$csvFilePath = "D:\TEMP\SP\sites.csv" # CSV file containing 2 columns: SiteUrlName,SiteTitle
$logFilePath = "D:\TEMP\SP\output.log" # Log output file
$tenantName = "demotenant" # Name of the SharePoint tenant
$adminUser = "admin@$tenantName.onmicrosoft.com" # Login name of the administrator user that can create sites
$siteMembersGroupName = "SG - Test Members" # Leave empty when not used
$siteVisitorsGroupName = "" # Leave empty when not used
$siteOwnersGroupName = "" # Leave empty when not used
$storageQuota = 1000

$hubUrl = "https://$tenantName.sharepoint.com"
$template = "STS#3" # STS#3 = Team site (no Office 365 group). You can check the available templates using the Get-SPOWebTemplate command

#region Functions
#####################################################################################################
function Import-KrSpoModule {
    $moduleName = "Microsoft.Online.SharePoint.PowerShell"
    if (-not(Get-module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue)) {
        Write-Error "It seems that the PowerShell module for SharePoint is not installed. Goto https://www.microsoft.com/en-us/download/details.aspx?id=35588 to download and install this module."
        break
    }
    if (-not(Get-Module -Name $moduleName)) {
        try {
            Import-Module -Name $moduleName -WarningAction Ignore
        }
        catch {
            Write-Error $ErrorMessage = $_.Exception.Message
            break
        }
    }
}

function Connect-KrSharePoint {
    # Let the user fill in their password in the PowerShell window
    Write-Host "#############################################################" -ForegroundColor Green
    $password = Read-Host "Please enter the password for $adminUser" -AsSecureString
 
    # Set credentials
    $credentials = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $adminUser, $password
    $adminUrl = "https://$tenantName-admin.sharepoint.com"
    
    # Connect to to Office 365
    try {
        Connect-SPOService -Url $adminUrl -Credential $credentials
        Write-Host "Info: Connected succesfully to Office 365" -foregroundcolor green
    }
    catch {
        Write-Host "Error: Could not connect to Office 365" -foregroundcolor red
        break
    }
}

function New-KrSpoSite {
    param(
        [parameter(Mandatory = $true)]
        [String]
        $SiteName,

        [parameter(Mandatory = $false)]
        [String]
        $SiteTitle
    )

    $url = "https://$tenantName.sharepoint.com/sites/$siteName"
 
    # Verify if site already exists in SharePoint Online
    $siteExists = Get-SPOSite -Filter "Url -eq $url"
 
    # Verify if site already exists in the recycle bin
    $siteExistsInRecycleBin = Get-SPODeletedSite | Where-Object "$_.url -eq $url"
 
    # Create site if it doesn't exists
    if (($null -eq $siteExists) -and ($null -eq $siteExistsInRecycleBin)) {
        Write-Host "Info: Start creating $($SiteTitle)..." -ForegroundColor Green
        New-SPOSite -Url $url -title $SiteTitle -Owner $adminUser -StorageQuota $storageQuota -Template $template -NoWait
    }
    elseif ($null -eq $siteExists) {
        Write-Host "Info: $($url) already exists" -ForegroundColor Red
    }
    else {
        Write-Host "Info: $($url) still exists in the recyclebin" -ForegroundColor Red
    }

    return $url
}

function Add-KrHubSiteAssociation {
    param(
        [parameter(Mandatory = $true)]
        [String]
        $SiteUrl,

        [parameter(Mandatory = $true)]
        [String]
        $HubUrl
    )
    #Write-Host "Info: Associating site with sitehub for $SiteUrl" -ForegroundColor Green
    try {
        Add-SPOHubSiteAssociation -Site $SiteUrl -HubSite $HubUrl
    }
    catch {
        throw $_
    }
}

function Add-KrSiteGroupUsers {
    param(
        [parameter(Mandatory = $true)]
        [String]
        $SiteUrl,

        [parameter(Mandatory = $false)]
        [String]
        $SiteTitle
    )

    if ($siteMembersGroupName -ne "") {
        #Write-Host "Info: Assigning site group members for $SiteTitle" -ForegroundColor Green
        $membersGroupName = "$SiteTitle Members"
        try {
            Add-SPOUser -Group $membersGroupName -LoginName $siteMembersGroupName -Site $SiteUrl
        }
        catch {
            throw $_
        }
    }

    if ($siteVisitorsGroupName -ne "") {
        #Write-Host "Info: Assigning site group visitors for $SiteTitle" -ForegroundColor Green
        $visitorsGroupName = "$SiteTitle Visitors"
        try {
            Add-SPOUser -Group $visitorsGroupName -LoginName $siteVisitorsGroupName -Site $SiteUrl   
        }
        catch {
            throw $_
        }
    }

    if ($siteOwnersGroupName -ne "") {
        #Write-Host "Info: Assigning site group owners for $SiteTitle" -ForegroundColor Green
        $ownersGroupName = "$SiteTitle Owners"
        try {
            Add-SPOUser -Group $ownersGroupName -LoginName $siteOwnersGroupName -Site $SiteUrl    
        }
        catch {
            throw $_
        }
    }
}

function Test-KrLogFile {
    if (Test-Path -Path $logFilePath) {
        do {
            Write-Host "#############################################################" -ForegroundColor Yellow
            Write-Host "The log file '$logFilePath' already exists. Do you want to overwrite it? (y/n)" -ForegroundColor Yellow -NoNewLine
            $confirm = Read-Host 
        } while (-not(('y', 'n').Contains($confirm)))
        $global:replaceLogFile = $confirm -eq 'y'
    
        if (-not($global:replaceLogFile)) {
            $folder = Split-Path $logFilePath -Resolve
            $fileName = (Split-Path $logFilePath -Leaf).Split('.')[0]
            $extension = (Split-Path $logFilePath -Leaf).Split('.')[1]
            $cnt = (Get-ChildItem -Path $folder -Filter $fileName*).Count
            $cnt++
            $newName = "$fileName`_$cnt.$extension"
            $logFilePath = Join-Path -Path $folder -ChildPath $newName
            Write-Host "New log file is: $logFilePath" -ForegroundColor Yellow
        }
        Write-Host "#############################################################" -ForegroundColor Yellow
    }
}

function Write-KrCreateSiteLog {
    param (
        [Parameter(Mandatory = $true)]
        [String] $FileName,
        [Parameter(Mandatory = $true)]
        [String] $SiteName,
        [String] $SiteTitle,
        [String] $Siteurl,
        [String] $Status,
        [String] $Duration
    )

    if ($replaceLogFile) {
        Remove-Item -Path $logFilePath -Force
        $global:replaceLogFile = $false # Reset variable
    }

    if (-not(Test-Path -Path $FileName)) {
        $header = "SiteName`tSiteTitle`tSiteUrl`tStatus`tDuration"
        Add-Content -Path $FileName -Value $header
    }
    
    # Create new log line for site or replace an existing log line
    $content = Get-Content -Path $FileName
    $siteLine = $content | Select-String $SiteName | Select-Object -ExpandProperty Line
    $newLine = "$SiteName`t$SiteTitle`t$SiteUrl`t$Status`t$Duration"
    if ($null -eq $siteLine) {
        Add-Content -Path $FileName -Value $newLine
    }
    else {
        if ($siteLine -ne $newLine) {
            $newContent = $content | ForEach-Object { $_ -replace $siteLine, $newLine }
            $newContent | Set-Content -Path $FileName
        }
    }
}

#####################################################################################################
#endregion Functions

$replaceLogFile = $false
Clear-Host

# Check if the SharePoint PowerShell modules are installed on the computer
Import-KrSpoModule

# Test if CSV file exists
if (-not(Test-Path -Path $csvFilePath)) {
    Write-Host "Error: CSV file '$csvFilePath' does not exists." -ForegroundColor Red
    break
}

# Test if log file exists
Test-KrLogFile

# Connect to customer SharePoint Online
Connect-KrSharePoint

# Read site list from CSV file
$sites = Import-Csv -Path $csvFilePath
$sites | Add-Member -MemberType NoteProperty -Name "SiteUrl" -Value $null
$sites | Add-Member -MemberType NoteProperty -Name "Status" -Value $null
$sites | Add-Member -MemberType NoteProperty -Name "StartTime" -Value $null
$sites | Add-Member -MemberType NoteProperty -Name "Duration" -Value $null

# For each site line in CSV, create a SharePoint site
foreach ($site in $sites) {
    if ($site.SiteTitle -eq "") {
        $site.SiteTitle = $site.SiteUrlName
    }

    $site.StartTime = (Get-Date)
    $siteUrl = New-KrSpoSite -SiteName $site.SiteUrlName -SiteTitle $site.SiteTitle
    $site.SiteUrl = $siteUrl
    $site.Status = (Get-SPOSite -Filter { Url -eq $siteUrl }).Status
    Write-KrCreateSiteLog -FileName $logFilePath -SiteName $site.SiteUrlName -SiteTitle $site.SiteTitle -Siteurl $siteUrl -Status $site.Status
}
#Write-KrCreateSiteLog -FileName $logFilePath -SiteName "TestSite11" -SiteTitle "Test site 11" -Siteurl "https://test" -Status "Retrying" -Duration 10

# Wait for site to get active
$waitSeconds = 90
Write-Host "`nInfo: Hang on for $waitSeconds seconds, while SharePoint creates the sites for us...`n" -ForegroundColor Green
Start-Sleep -Seconds $waitSeconds

Write-Host "`nInfo: Let's start associating sites to a hub and assigning site group users..." -ForegroundColor Green
do {
    # Reset check
    $allSitesFinished = $true

    # For each site line in CSV, create a SharePoint site
    foreach ($site in $sites) {
        if ($site.Status -ne "Finished") {
            $siteUrl = $site.SiteUrl
            $siteStatus = (Get-SPOSite -Filter "Url -eq $siteUrl").Status
            if ($siteStatus -eq "Active") {
                try {
                    Add-KrHubSiteAssociation -SiteUrl $siteUrl -HubUrl $hubUrl
                    Add-KrSiteGroupUsers -SiteUrl $siteUrl -SiteTitle $site.SiteTitle
                    $site.Status = "Finished"

                    $site.Duration = [Math]::Round((New-TimeSpan -Start $site.StartTime -End (Get-Date)).TotalSeconds)
                }
                catch {
                    $allSitesFinished = $false
                    $site.Status = "Retry"
                }

                Write-KrCreateSiteLog -FileName $logFilePath -SiteName $site.SiteUrlName -SiteTitle $site.SiteTitle -Siteurl $siteUrl -Status $site.Status -Duration $site.Duration
            }
            else {
                $allSitesFinished = $false
            }
        }
    }

    # Wait a moment with every retry loop
    Start-Sleep -Seconds 10

} while (-not($allSitesFinished))

Write-Host "`nInfo: All sites have been created." -ForegroundColor Green
Write-Host "`nCheck the log file '$logFilePath' for information." -ForegroundColor Green
Write-Host "#############################################################" -ForegroundColor Green
Import-Csv -Path $logFilePath -Delimiter "`t" | Format-Table
