<#
.SYNOPSIS
Tests the health of a Lync Server 2013 environment
Current Version: 1.00

.DESCRIPTION
This script will test various components of an Lync Server 2013
environment including services started, Test cmdlets, SQL Server availability,
file share availability, and CMS replication status.

.EXAMPLE
Test-LyncHealth
This command will run checks against the entire Lync environment

.NOTES
Written by Jeff Brown
Jeff@UpStartTech.com
@JeffWBrown
www.upstarttech.com

Any and all technical advice, scripts, and documentation are provided as is with no guarantee.
Always review any code and steps before applying to a production system to understand their full impact.

Version Notes
V1.00 - 7/3/2015 - Initial Version
#>

#*************************************************************************************
#************************       Parameters    ****************************************
#*************************************************************************************

[CmdletBinding()]
param()

#*************************************************************************************
#**************************     Variables     ****************************************
#*************************************************************************************

[array]$AllUserPools    = @(Get-CsService -UserServer | Select-Object -ExpandProperty PoolFqdn)
[array]$AllMedPools		= @(Get-CsService -MediationServer | Select-Object -ExpandProperty PoolFqdn)
[array]$AllDirPools		= @(Get-CsService -Director | Select-Object -ExpandProperty PoolFqdn)
[array]$AllPChatPools	= @(Get-CsService -PersistentChatServer | Select-Object -ExpandProperty PoolFqdn)
[array]$AllEdgePools    = @(Get-CsService -EdgeServer | Select-Object -ExpandProperty PoolFqdn)
[array]$AllFileStores	= @(Get-CsService -FileStore | Select-Object -ExpandProperty UncPath)
[array]$AllSQLServers   = @(Get-CsService -ApplicationDatabase | Select-Object -ExpandProperty PoolFqdn)

[array]$AllLyncServers  = @()
[array]$AllEdgeServers  = @()

# Dynamically builds list of all Lync Servers based on pools
foreach ($pool in $AllUserPools)
{
    $AllLyncServers += Get-CsComputer -Pool $pool | Select-Object -ExpandProperty Fqdn
}

foreach ($pool in $AllDirPools)
{
    $AllLyncServers += Get-CsComputer -Pool $pool | Select-Object -ExpandProperty Fqdn
}

foreach ($pool in $AllMedPools)
{
    $AllLyncServers += Get-CsComputer -Pool $pool | Select-Object -ExpandProperty Fqdn
}

foreach ($pool in $AllPChatPools)
{
    $AllLyncServers += Get-CsComputer -Pool $pool | Select-Object -ExpandProperty Fqdn
}

# Sorts $AllLyncServers by Name
$AllLyncServers = $AllLyncServers | Sort-Object

# Dynamically builds list of Edge servers based on Edge pools
foreach ($pool in $AllEdgePools)
{
    $AllEdgeServers += Get-CsComputer -Pool $pool | Select-Object -ExpandProperty Fqdn
}

# Sorts $AllEdgeServers by Name
$AllEdgeServers = $AllEdgeServers | Sort-Object

#*************************************************************************************
#**************************     Main Code     ****************************************
#*************************************************************************************

Write-Host "`nRunning Invoke-CsManagementStoreReplication to verify replication status at end of script" -ForegroundColor Yellow
Invoke-CsManagementStoreReplication

Start-Sleep -Seconds 5

Write-Host "`nChecking Services`n" -ForegroundColor Yellow
foreach ($server in $AllLyncServers)
{
    Write-Host "$server : " -NoNewline    
    if ((Test-Connection -ComputerName $server -Count 2 -Quiet) -eq $true)
    {
        [array]$services = Get-CsWindowsService -ComputerName $server | select DisplayName, Status
	    [bool]$servicesGood = $true
        [array]$badServices = @()
        	    
	    foreach ($service in $services)
	    {
		    if ($service.Status -ne "Running")
		    {
			    $servicesGood = $false
                $badServices += $service.DisplayName
		    }
	    }

        if ($servicesGood -eq $true)
	    {
		    Write-Host "All Services Running" -ForegroundColor green
	    }
	    else
	    {
		    Write-Host "Services Not Running" -ForegroundColor White -BackgroundColor DarkRed
		    Write-Host "Following Services are Not Running:"
		    foreach ($badService in $badServices)
		    {
			    Write-Host $badService
		    }
	    }
    }
    else
    {
        Write-Host "Server Unavailable" -ForegroundColor White -BackgroundColor DarkRed
    } #End of if ((Test-Connection -ComputerName $server -Count 2 -Quiet) -eq $true)
} # End of Checking Services

# Runs Test-Cs cmdlets against each user pool
# Some Test cmdlets only seem to work when running locally on a Lync server
foreach ($pool in $AllUserPools)
{
    Write-Host "`nChecking $pool`n" -ForegroundColor Yellow
    
    Write-Host "Test-CsAddressBookService (Address Book Web Service) : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsAddressBookService -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "Test-CsAddressBookWebQuery (Address Book Search) : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsAddressBookWebQuery -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "Test-CsASConference (Application Sharing Conference) : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsASConference -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "Test-CsAVConference : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsAVConference -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "Test-CsAVEdgeConnectivity (Edge Server Connectivity) : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsAVEdgeConnectivity -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "Test-CsDataConference : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsDataConference -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "Test-CsGroupIM (Conference IM) : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsGroupIM -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "Test-CsIM (Peer-to-Peer IM) : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsIM -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "Test-CsPresence (Publish Presence) : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsPresence -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "Test-CsRegistration (User Log On) : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsRegistration -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "Test-CsUcwaConference (Lync Web App) : " -NoNewLine
    $test = Invoke-Expression -Command "Test-CsUcwaConference -TargetFqdn $pool"
    if ($test.Result -eq "Success") {Write-Host $test.Result -ForegroundColor Green} else {Write-Host $test.Result -ForegroundColor White -BackgroundColor DarkRed}

    Write-Host "`nChecking $pool Database Status`n" -ForegroundColor Yellow
    $databaseStatus = Get-CsDatabaseMirrorState -PoolFqdn $pool
    foreach ($status in $databaseStatus)
    {
        Write-Host $status.DatabaseName": " -NoNewline
        if ($status.MirroringStatusOnPrimary -ine "synchronized")
        {
            Write-Host "Not Synchronized" -ForegroundColor White -BackgroundColor DarkRed
        }
        elseif ($status.StateOnPrimary -ine "Principal")
        {
            Write-Host "Not on Principal" -ForegroundColor White -BackgroundColor DarkRed
        }
        else
        {
            Write-Host "Synchronized" -ForegroundColor Green
        }
    }
} # End of Tests for Each Pool

# Testing if Edge Servers respond to connectivity test
Write-Host "`nPinging Edge Servers`n" -ForegroundColor Yellow
foreach ($edge in $AllEdgeServers)
{
    Write-Host $edge": " -NoNewline
    if ((Test-Connection -ComputerName $edge -Count 2 -Quiet) -eq $true)
    {
        Write-Host "Available" -ForegroundColor Green
    }
    else
    {
        Write-Host "Not Available" -ForegroundColor White -BackgroundColor DarkRed
    }
} # End of testing Edge server connectivity

# Testing if SQL Servers respond to connectivity test
Write-Host "`nPinging SQL Servers`n" -ForegroundColor Yellow
foreach ($sql in $AllSQLServers)
{
    Write-Host $sql": " -NoNewline
    if ((Test-Connection -ComputerName $sql -Count 2 -Quiet) -eq $true)
    {
        Write-Host "Available" -ForegroundColor Green
    }
    else
    {
        Write-Host "Not Available" -ForegroundColor White -BackgroundColor DarkRed
    }
} # End of testing SQL server connectivity

# Checks availability of each pool's file share
Write-Host "`nTesting File Share Availability`n" -ForegroundColor Yellow
foreach ($store in $AllFileStores)
{
    Write-Host "File Share $store : " -NoNewLine
	#$test = Test-Path -Path $store
	if ((Test-Path -Path $store) -eq $true)
    {
        Write-Host "Success" -ForegroundColor Green
    }
    else
    {
        Write-Host "Failure" -ForegroundColor White -BackgroundColor DarkRed
    }
} # End of testing file shares

# Checks all servers for up to topology replication
# Invoke-CsManagementStoreReplication was ran at beginning of script to check UpToDate field
Write-Host "`nChecking Replication Status" -ForegroundColor Yellow
Write-Host "`nServer`t`t`t`tUp To Date`t`tLast Status Report" -ForegroundColor yellow
$replServers = Get-CsManagementStoreReplicationStatus | Sort-Object ReplicaFqdn
foreach ($replServer in $replServers)
{	
	Write-Host $replServer.ReplicaFqdn `t`t -NoNewLine
	if ($replServer.UpToDate -eq "True")
	{
		Write-Host $replServer.UpToDate `t`t`t -ForegroundColor Green -NoNewLine
		Write-Host $replServer.LastStatusReport
	}
	else
	{
		Write-Host $replServer.UpToDate `t`t`t`t`t`t -ForegroundColor Red
	}
} # End of Replication Test