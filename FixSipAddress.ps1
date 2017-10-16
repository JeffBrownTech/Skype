

[CmdletBinding()]
param(
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = "Enter the SIP address to search for")]
    [string[]]$SipAddress
)

BEGIN
{
    # Check either server module or remote PS session exists
    $availableModules = Get-Module -ListAvailable
    $loadedModules = Get-Module
    
    if ($availableModules.Name -notcontains "Lync")
    {        
        Write-Verbose -Message "Lync Management Shell not found"
    }
    elseif ($availableModules.Name -notcontains "SkypeForBusiness")
    {
        Write-Verbose -Message "Skype for Business Management Shell not found"
    }
    elseif ($availableModules.Description -notlike "*/ocspowershell")
    {
        Write-Verbose -Message "No Remote PowerShell Session found to a Lync or Skype for Business Server"            
    }

    # Get list of all Directors and Front-Ends, STD Edition Servers
    #[System.Collections.ArrayList]$csServers = @()

    [array]$allFrontEndPools = @((Get-CsService -UserServer).PoolFqdn)
    [array]$allDirectorPools = @((Get-CsService -Director).PoolFqdn)
    [array]$allCsServers = @()

    if ($null -ne $allFrontEndPools[0])
    {
        foreach ($fePool in $allFrontEndPools)
        {
            $allCsServers += (Get-CsComputer -Pool $fePool).Fqdn
        }
    }

    if ($null -ne $allDirectorPools)
    {
        foreach ($dirPool in $allDirectorPools)
        {
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
            $message = "A user account currently exists for $sip. Do you wish to disable and remove this account (this will result in data loss)?"
            $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Disables and removes the user account."
            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Skips disabling account and SQL cleanup."
            $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
            $answer = $host.UI.PromptForChoice($title,$message,$options,0)

            switch ($answer)
            {
                0 {
                    $continue = $true
                    
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
                
                1 {$continue = $false}
            }
        }
        catch # $sip is not attached to an account but may still have leftover entries in the SQL Express RTCLOCAL database
        {
            $title = "No User Account Found"
            $message = "A user account was not found for $sip. Do you wish to continue searching for a leftover account attached to this address?"
            $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Searches and removes any leftover accounts on the servers through SQL."
            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Skips SQL cleanup for this account."
            $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
            $answer = $host.UI.PromptForChoice($title,$message,$options,0)

            switch ($answer)
            {
                0 {$continue = $true} #Disable user
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
                        Write-Warning -Message "$csServer was not found or is not accessible. Verify firewall rules allow UDP 1434 and SQLServr.exe as exceptions."
                    }
                    else
                    {
                        Write-Warning -Message $connErrorMsg
                    }

                    CONTINUE
                }

                # Create SQL Query
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
                    $sqlRemoveUserCmd = New-Object System.Data.SqlClient.SqlCommand
                    $sqlRemoveUserCmd.Connection = $sqlConn
                    $sqlRemoveUserCmd.CommandType = [System.Data.CommandType]'StoredProcedure'
                    $sqlRemoveUserCmd.CommandText = "rtc.dbo.RtcDeleteResource"
                    $sqlRemoveUserCmd.Parameters.AddWithValue("@_UserAtHost",[string]$sip)
                    $sqlRemoveUserAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlRemoveUserCmd
                    $removeUserResults = New-Object System.Data.DataSet
                    $sqlRemoveUserAdapter.Fill($removeUserResults)
                }
                else
                {
                    Write-Verbose -Message "$sip was not found on $csServer"
                }

                # Close SQL Connection
                $sqlConn.Close()
            } # End of foreach ($csServer in $allCsServers)
        } # End of if ($continue -eq $true
    } # End of foreach ($sip in $SipAddress)
} # End of PROCESS