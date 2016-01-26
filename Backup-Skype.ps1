<#
.SYNOPSIS
Creates backup of various components of a Skype for Business environment
Current Version: 1.00

.DESCRIPTION
This script will back up various components of a Skype for Business environment
including user data, CMS, voice routing information, and response groups settings.

.EXAMPLE
Backup-Skype
This command will back up the Skype environment using to the predefined backup location.

.NOTES
Written by Jeff Brown
Jeff@JeffBrown.tech
@JeffWBrown
www.jeffbrown.tech

Any and all technical advice, scripts, and documentation are provided as is with no guarantee.
Always review any code and steps before applying to a production system to understand their full impact.

Version Notes
V1.00 - 1/26/2016 - Initial Version
#>

#*************************************************************************************
#************************       Parameters    ****************************************
#*************************************************************************************

[CmdletBinding()]
Param()

#*************************************************************************************
#**************************     Variables     ****************************************
#*************************************************************************************

# Array of network paths to back up to one or multiple locations
# Change UNC Path to Existing File Share
[array]$FileShares = @("\\SERVER\FILESHARE")

# Get various date fields
$RemoveDate = ((Get-Date).AddDays(-30))
$Now = Get-Date -UFormat %m-%d-%Y_%H%M%S

# Create array of file shares if you want to back up to multiple paths
[array]$BackupPaths = @("$FileShare\$Now")

# Gets all Front End User Pools in the Environment
[array]$AllUserPools = @((Get-CsService -UserServer).PoolFqdn)

#*************************************************************************************
#**************************     Main Code     ****************************************
#*************************************************************************************

# Remove back ups older than 30 days
# Adjust the variable $RemoveDate to your preference
foreach ($Share in $FileShares)
{
    Write-Verbose -Message "Deleting Old Backups in $Share"
    Get-ChildItem -Path $Shared | Where-Object {$_.LastWriteTime -lt $RemoveDate} | Remove-Item -Recurse
}

foreach ($Path in $BackupPaths)
{    
    # Creates backup directory for current backup if it does not exist
    if ((Test-Path -Path $Path) -eq $false)
    {
        Write-Verbose "Creating backup directory $Path"
        New-Item -Path $Path -ItemType directory | Out-Null
    }

    # Export Common Items to all Pools
    Write-Verbose "Exporting Topology"
    Export-CsConfiguration -FileName $Path\TopologyBackup.zip
    (Get-CsTopology -AsXml).ToString() > $Path\TopologyBackup.xml

    Write-Verbose "Exporting Location Information Server"
    Export-CsLisConfiguration -FileName $Path\LISConfig.zip

    Write-Verbose "Exporting Dial Plans"
    Get-CsDialPlan | Export-Clixml -Path $Path\DialPlan.xml

    Write-Verbose "Exporting Voice Policies"
    Get-CsVoicePolicy | Export-Clixml -Path $Path\VoicePolicy.xml

    Write-Verbose "Exporting Voice Routes"
    Get-CsVoiceRoute | Export-Clixml -Path $Path\VoiceRoute.xml

    Write-Verbose "Exporting PSTN Usage"
    Get-CsPstnUsage | Export-Clixml -Path $Path\PSTNUsage.xml

    Write-Verbose "Exporting Voice Configuration"
    Get-CsVoiceConfiguration | Export-Clixml -Path $Path\VoiceConfiguration.xml

    Write-Verbose "Exporting Trunk Configuration"
    Get-CsTrunkConfiguration | Export-Clixml -Path $Path\TrunkConfiguration.xml
    
    # Export Pool Specific Settings and Applications
    foreach ($Pool in $AllUserPools)
    {
        Write-Verbose "Exporting $Pool Response Groups"
        Export-CsRgsConfiguration -Source “ApplicationServer:$Pool” -FileName $Path\$Pool-RgsConfig.zip

        Write-Verbose "Exporting $Pool User Data"
        Export-CsUserData -PoolFqdn $Pool -FileName $Path\$Pool-UserData.zip
    } # End Of foreach ($Pool in $AllUserPools)
} # End of foreach ($Path in $BackupPaths)