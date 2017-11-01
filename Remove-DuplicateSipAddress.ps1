<#
.SYNOPSIS
Disables user accounts and removes account resource from servers.

.DESCRIPTION
This script removes Lync or Skype for Business user accounts that have
been removed but still exist in the RTCLOCAL instance on Standard Edition,
Front End, and Director servers. The script will first check if the SIP Address
is attached to an active account. If so, it will prompt to disable the account
before running the clean up actions against SQL Express databases.

.PARAMETER SipAddress
This is the SIP Address of the user account to search for and to remove from
the SQL databases.

.EXAMPLE
.\Remove-DuplicateSipAddress.ps1 -SipAddress john@contoso.com
Example 1 will see if john@contoso.com is attached to an enabled user account.
If it is, it will prompt to disable the user account, then run clean up actions
on each server.

.NOTES
Written by Jeff Brown
Jeff@JeffBrown.tech
@JeffWBrown
www.jeffbrown.tech

Any and all technical advice, scripts, and documentation are provided as is with no guarantee.
Always review any code and steps before applying to a production system to understand their full impact.

Version Notes
V1.0 - 10/31/2017 - Initial Version
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = "Enter the SIP address to search for")]
    [string[]]$SipAddress
)

BEGIN
{
    # Check either server module or remote PS session exists
    # We also temporarily change the verbosity as the Get-Module commands produce a lot of extraneous output if using the -Verbose parameter
    $scriptVerboseLevel = $VerbosePreference
    $VerbosePreference = "SilentlyContinue"
    $availableModules = Get-Module -ListAvailable
    $VerbosePreference = $scriptVerboseLevel
    [bool]$foundModule = $false

    if ($availableModules.Name -contains "Lync" -or $availableModules.Name -contains "SkypeForBusiness")
    {        
        Write-Verbose -Message "PowerShell Management Shell module is installed."
        $foundModule = $true
    }
    
    if ($availableModules.Description -like "*/ocspowershell")
    {
        Write-Verbose -Message "Remote PowerShell Session is created."
        $foundModule = $true
    }
    
    if ($foundModule -eq $false)
    {
        Write-Warning -Message "No Lync or Skype for Business Management Shell modules or remote PowerShell Sessions found."
        Write-Warning -Message "These are required for this script as it relies on native Lync/Skype for Business cmdlets."
        Write-Warning -Message "Please verify these are installed or create before running script again."
        Write-Warning -Message "It is recommended to run this from a Lync or Skype for Business Server."
        EXIT
    }
    
    # Get list of all Directors and Standard Edition/Enterprise Pools
    Write-Verbose -Message "Gathering all Front End and Director Pools in environment."
    [array]$allFrontEndPools = @((Get-CsService -UserServer).PoolFqdn)
    [array]$allDirectorPools = @((Get-CsService -Director).PoolFqdn)
    
    # Find individual servers in each pool and save into single array
    [array]$allCsServers = @()
    
    # Check to see if the first element in the array is empty; if not, find all servers in the pool    
    if ($null -ne $allFrontEndPools[0])
    {
        foreach ($fePool in $allFrontEndPools)
        {
            Write-Verbose -Message "Gathering all servers in $fePool."
            $allCsServers += (Get-CsComputer -Pool $fePool).Fqdn
        }
    }

    if ($null -ne $allDirectorPools[0])
    {
        foreach ($dirPool in $allDirectorPools)
        {
            Write-Verbose -Message "Gathering all servers in $dirPool."
            $allCsServers += (Get-CsComputer -Pool $dirPool).Fqdn
        }
    }
} # End of BEGIN

PROCESS
{
    foreach ($sip in $SipAddress)
    {
        [bool]$continue = $false

        # Check if $sip is attached to an active user account
        try
        {
            $sipUserInfo = Get-CsUser -Identity $sip -ErrorAction STOP
            
            $title = "Existing User Account Found"
            $message = "A user account currently exists for $sip on pool $($sipUserInfo.RegistrarPool). Do you wish to disable and remove this account (this will result in data loss)?"
            $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Disables and removes the user account."
            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Skips disabling account and SQL cleanup."
            $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
            $answer = $host.UI.PromptForChoice($title,$message,$options,0)

            switch ($answer)
            {
                # Yes answer
                0 {
                    $continue = $true
                    
                    # Attempt to disable/remove user account
                    try
                    {
                        Disable-CsUser -Identity $sip -ErrorAction STOP
                    }
                    catch
                    {
                        Write-Warning -Message $_
                        CONTINUE
                    }

                    # Verifies user is disabled successfully
                    do
			        {
				        Write-Verbose -Message "Waiting for $sip to be disabled"
				        $userEnabledCheck = Get-CsUser -Identity $sip -ErrorAction SilentlyContinue
				        Start-Sleep -Seconds 3
			        } until ($null -eq $userEnabledCheck)
                } # End of switch 0
                
                # No answer
                1 {$continue = $false}
            }
        }
        catch # $sip is not attached to an account but may still have leftover entries in the SQL Express RTCLOCAL database
        {
            $title = "No Active User Account Found"
            $message = "An active user account was not found for $sip. Do you wish to continue searching for a leftover account attached to this address?"
            $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Searches and removes any leftover accounts on the servers through SQL."
            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Skips SQL cleanup for this account."
            $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
            $answer = $host.UI.PromptForChoice($title,$message,$options,0)

            switch ($answer)
            {
                # Yes answer
                0 {$continue = $true}
                
                # No answer
                1 {$continue = $false}
            }
        }
        
        if ($continue -eq $true)
        {
            foreach ($csServer in $allCsServers)
            {
                # Attempt to create SQL Connection
                try
                {
                    $sqlConn = New-Object System.Data.SqlClient.SqlConnection -ErrorAction STOP
                    $sqlConn.ConnectionString = "Server=$csServer\rtclocal;Integrated Security=true;Initial Catalog=master"
                    $sqlConn.Open()
                }
                catch
                {
                    $connErrorMsg = $_.Exception.Message
                
                    if ($connErrorMsg -like "*The server was not found or was not accessible*")
                    {
                        Write-Warning -Message "$csServer was not found or was not accessible. Verify firewall rules allow UDP 1434 and SQLServr.exe as exceptions."
                    }
                    else
                    {
                        Write-Warning -Message $connErrorMsg
                    }

                    CONTINUE
                }

                # Create SQL Query and get results
                $sqlSearchUserCmd = New-Object System.Data.SqlClient.SqlCommand
                $sqlSearchUserCmd.Connection = $sqlConn
                $searchQuery = "SELECT * FROM [rtc].[dbo].[Resource] WHERE [UserAtHost]='$sip'"
                $sqlSearchUserCmd.CommandText = $searchQuery

                $sqlSearchUserAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlSearchUserCmd
                $searchResults = New-Object System.Data.DataSet
                $sqlSearchUserAdapter.Fill($searchResults) | Out-Null

                if ($null -ne $searchResults.Tables.UserAtHost)
                {
                    Write-Verbose -Message "Found $sip on $csServer, attempting removal"
                    
                    # Remove user
                    try
                    {
                        $sqlRemoveUserCmd = New-Object System.Data.SqlClient.SqlCommand
                        $sqlRemoveUserCmd.Connection = $sqlConn
                        $sqlRemoveUserCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
                        $sqlRemoveUserCmd.CommandText = "rtc.dbo.RtcDeleteResource"
                        $sqlRemoveUserCmd.Parameters.AddWithValue("@_UserAtHost",[string]$sip) | Out-Null # Prevents some wacky output to the screen
                        $sqlRemoveUserAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlRemoveUserCmd
                        $removeUserResults = New-Object System.Data.DataSet
                        $nullResults = $sqlRemoveUserAdapter.Fill($removeUserResults) # Hides output to the screen by saving to a variable

                        $outputObj = [PSCustomObject][ordered]@{
                            RTCLocal= $csServer
                            SipAddress = $sip
                            Result = "SIP Address record was removed"
                        }
                    }
                    catch
                    {
                        #Write-Warning $_
                        $outputObj = [PSCustomObject][ordered]@{
                            RTCLocal= $csServer
                            SipAddress = $sip
                            Result = "ERROR: Something happened when removing SIP Address record: $($_)"
                        }
                    }
                    finally
                    {
                        Write-Output $outputObj
                        $sqlConn.Close()
                    }
                }
                else
                {
                    $outputObj = [PSCustomObject][ordered]@{
                        RTCLocal= $csServer
                        SipAddress = $sip
                        Result = "No SIP Address record found"
                    }
                    Write-Output $outputObj
                    $sqlConn.Close()
                } # End of if ($null -ne $searchResults.Tables.UserAtHost)
            } # End of foreach ($csServer in $allCsServers)
        } # End of if ($continue -eq $true)
    } # End of foreach ($sip in $SipAddress)
} # End of PROCESS