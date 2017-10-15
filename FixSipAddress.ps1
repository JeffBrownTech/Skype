

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
            $sqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $sqlCmd.Connection = $sqlConn
            $searchQuery = "SELECT * FROM [rtc].[dbo].[Resource] WHERE [UserAtHost]='$sip'"
            $sqlCmd.CommandText = $searchQuery

            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCmd
            $searchResults = New-Object System.Data.DataSet
            $adapter.Fill($searchResults) | Out-Null

            if ($null -ne $searchResults.Tables.UserAtHost)
            {
                # Found user
                # Remove user
            }
            else
            {
                # Some message saying user not found
            }

            # Close SQL Connection
            $sqlConn.Close()
        } # End of foreach ($csServer in $allCsServers)
    } # End of foreach ($sip in $SipAddress)
} # End of PROCESS