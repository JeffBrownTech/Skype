

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
    
} # End of PROCESS