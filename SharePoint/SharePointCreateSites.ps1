<##################################################################################
# NAME:   SharePointCreateSites.ps1
# AUTHOR: Niels van de Coevering
# DATE:   4 October 2020
#
# COMMENTS: This script will creates new SharePoint sites specified in a CSV, 
# associate these sites with a hub and assign Office users or groups to 
# site groups.
# The CSV file should contain 2 columns: SiteUrlName,SiteTitle. When the site title
# it not specified, the site name is used as site title.
#
# CSV EXAMPLE:
# SiteUrlName,SiteTitle
# TestSite1,Test site 1
# TestSite2,Test site 2
# TestSite3,
# TestSite4,
# TestSite5,Another site title
##################################################################################>

$csvFilePath = "C:\TEMP\SP\sites.csv" # CSV file containing 2 columns: SiteUrlName,SiteTitle
$logFilePath = "C:\TEMP\SP\output.log" # Log output file
$tenantName = "demotenant" # Name of the SharePoint tenant
$adminUser = "admin@$tenantName.onmicrosoft.com" # Login name of the administrator user that can create sites
$hubUrl = "https://$tenantName.sharepoint.com" # URL for the hub site
$siteMembersGroupName = "SG - Test Members" # Leave empty when not used
$siteVisitorsGroupName = "" # Leave empty when not used
$siteOwnersGroupName = "" # Leave empty when not used
$documentLibraryToAdd = "Archive" # Add a single new document library to the new sites
$quickLaunchItemsToDelete = @("Site contents", "Notebook", "Pages")

$template = "STS#3" # STS#3 = Team site (no Office 365 group). You can check the available templates using the Get-SPOWebTemplate command
$storageQuota = 1000 # 

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
    $global:password = Read-Host "Please enter the password for $adminUser" -AsSecureString
 
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

function Test-KrHubSite {
    if ($null -eq (Get-SPOSite -Identity $hubUrl)) {
        Write-Host "Error: The defnied hub site '$hubUrl' does not exist in SharePoint." -ForegroundColor Red
        break
    }
    if ($null -eq (Get-SPOHubSite -Identity $hubUrl)) {
        Write-Host "Error: The defnied hub site '$hubUrl' is not marked as a hub site in SharePoint." -ForegroundColor Red
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

function New-KrSpoDocumentLib {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $SiteUrl,

        [Parameter(Mandatory = $true)]
        [String]
        $Title,

        [String]
        $Description
    )
    $listTemplate = 101 # Document library

    # set SharePoint Online credentials
    $SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($global:adminUrl, $global:password)
         
    # Creating client context object
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $context.credentials = $SPOCredentials
     
    #create list using ListCreationInformation object (lci)
    $lci = New-Object Microsoft.SharePoint.Client.ListCreationInformation
    $lci.Title = $Title
    $lci.Description = $Description
    $lci.TemplateType = $listTemplate
    $list = $context.Web.Lists.Add($lci)
    $context.Load($list)

    #send the request containing all operations to the server
    try {
        $context.ExecuteQuery()
        Write-Host "Info: Created $($listTitle)" -ForegroundColor Green
    }
    catch {
        Write-Host "Info: $($_.Exception.Message)" -ForegroundColor Red
    }  
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

function New-DocumentLibrary {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $SiteUrl,

        [Parameter(Mandatory = $true)]
        [String]
        $Admin,

        [Parameter(Mandatory = $true)]
        [System.Security.SecureString]
        $Password,

        [Parameter(Mandatory = $true)]
        [String]
        $Title,

        [String]
        $Description
    )

    # Creating client context object
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Admin, $Password)
    $context.credentials = $SPOCredentials
    $site = $context.Web
    $context.Load($site)

    # Check if document library already exists
    $listExists = $true
    try {
        $myList = $context.Web.Lists.GetByTitle($Title)
        $context.Load($myList)
        $context.ExecuteQuery()
    }
    catch {
        $listExists = $false
    }

    if (-not($listExists)) {
        # Create new archive document library
        $myList = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $myList.Title = $Title
        $myList.Description = $Description
        $myList.TemplateType = 101 # Document library
        $newList = $context.Web.Lists.Add($myList)
        $context.Load($newList)
        $context.ExecuteQuery()

        # Update archive library and show in quick launch menu
        $myList = $context.Web.Lists.GetByTitle($Title)
        $context.Load($myList)
        $myList.OnQuickLaunch = $true
        $myList.Update()
        $context.Load($myList)
        $context.ExecuteQuery()
    }
}

function Remove-SiteQuickLaunchItems {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $SiteUrl,

        [Parameter(Mandatory = $true)]
        [String]
        $Admin,

        [Parameter(Mandatory = $true)]
        [System.Security.SecureString]
        $Password,

        [Parameter(Mandatory = $true)]
        [String[]]
        $ItemsToDelete
    )

    # Creating client context object
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Admin, $global:password)
    $context.credentials = $SPOCredentials
    $site = $context.Web
    $context.Load($site)
    $context.ExecuteQuery()

    # Remove 
    $quickLaunchNodes = $site.Navigation.QuickLaunch
    $context.Load($quickLaunchNodes)
    $context.ExecuteQuery()

    for ($i = $quickLaunchNodes.Count - 1; $i -ge 0; $i--) {
        switch ($quickLaunchNodes[$i].Title) {
            { $ItemsToDelete -contains $_ } {
                $quickLaunchNodes[$i].DeleteObject()
                $context.ExecuteQuery()
            }
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

function Get-KrSiteArray {
    # Check for delimiter comma or semicolon in csv file
    $csvContent = Get-Content -Path $csvFilePath
    $commaCnt = ($csvContent.ToCharArray() | Where-Object { $_ -eq ',' } | Measure-Object).Count
    $semiColonCnt = ($csvContent.ToCharArray() | Where-Object { $_ -eq ';' } | Measure-Object).Count
    $delimiter = if ($semiColonCnt -gt $commaCnt) { ';' } else { ',' }

    # Read site list from CSV file
    $siteArray = Import-Csv -Path $csvFilePath -Delimiter $delimiter
    $siteArray | Add-Member -MemberType NoteProperty -Name "SiteUrl" -Value $null
    $siteArray | Add-Member -MemberType NoteProperty -Name "Status" -Value $null
    $siteArray | Add-Member -MemberType NoteProperty -Name "StartTime" -Value $null
    $siteArray | Add-Member -MemberType NoteProperty -Name "Duration" -Value $null

    return $siteArray
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

# Read CSV file into array
$sites = Get-KrSiteArray

# Connect to customer SharePoint Online
Connect-KrSharePoint

# Check if defined hub site is a hub site
Test-KrHubSite

# TODO: Test if user groups exists

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

# Wait for site to get active
$waitSeconds = 0
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
                    New-DocumentLibrary -SiteUrl $siteUrl -Admin $adminUser -Password $password -Title $documentLibraryToAdd
                    Remove-SiteQuickLaunchItems -SiteUrl $siteUrl -Admin $adminUser -Password $password -ItemsToDelete $quickLaunchItemsToDelete
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
