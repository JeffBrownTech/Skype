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
[array]$FileShares = @("\\SERVER\FILESHARE","\\SERVER2\FILESHARE")

# Get various date fields
$RemoveDate = ((Get-Date).AddDays(-30))
$Now = Get-Date -UFormat %m-%d-%Y_%H%M%S

# Gets all Front End User Pools in the Environment
[array]$AllUserPools = @((Get-CsService -UserServer).PoolFqdn)

#*************************************************************************************
#**************************     Main Code     ****************************************
#*************************************************************************************

# Remove back ups older than 30 days
# Adjust the variable $RemoveDate to your preference
foreach ($Share in $FileShares)
{
    Write-Verbose -Message "Deleting Backups in $Share older than $RemoveDate"
    Get-ChildItem -Path $Share | Where-Object {$_.LastWriteTime -lt $RemoveDate} | Remove-Item -Recurse
}

foreach ($Share in $FileShares)
{
    # Create UNC path for current backup
    [string]$BackupPath = "$Share\$Now"

    # Creates backup directory for current backup if it does not exist
    if ((Test-Path -Path $BackupPath) -eq $false)
    {
        Write-Verbose -Message "Creating backup directory $BackupPath"
        New-Item -Path $BackupPath -ItemType Directory | Out-Null
    }

    # Export Common Items to all Pools
    Write-Verbose -Message "Exporting Topology"
    Export-CsConfiguration -FileName "$BackupPath\TopologyBackup.zip"
    (Get-CsTopology -AsXml).ToString() > "$BackupPath\TopologyBackup.xml"

    Write-Verbose -Message "Exporting Location Information Server"
    Export-CsLisConfiguration -FileName "$BackupPath\LISConfig.zip"

    Write-Verbose -Message "Exporting Dial Plans"
    Get-CsDialPlan | Export-Clixml -Path "$BackupPath\DialPlan.xml"

    Write-Verbose -Message "Exporting Voice Policies"
    Get-CsVoicePolicy | Export-Clixml -Path "$BackupPath\VoicePolicy.xml"

    Write-Verbose -Message "Exporting Voice Routes"
    Get-CsVoiceRoute | Export-Clixml -Path "$BackupPath\VoiceRoute.xml"

    Write-Verbose -Message "Exporting PSTN Usage"
    Get-CsPstnUsage | Export-Clixml -Path "$BackupPath\PSTNUsage.xml"

    Write-Verbose -Message "Exporting Voice Configuration"
    Get-CsVoiceConfiguration | Export-Clixml -Path "$BackupPath\VoiceConfiguration.xml"

    Write-Verbose -Message "Exporting Trunk Configuration"
    Get-CsTrunkConfiguration | Export-Clixml -Path "$BackupPath\TrunkConfiguration.xml"
    
    # Export Pool Specific Response Group Settings and User Data
    foreach ($Pool in $AllUserPools)
    {
        Write-Verbose -Message "Exporting $Pool Response Groups"
        Export-CsRgsConfiguration -Source “ApplicationServer:$Pool” -FileName "$BackupPath\$Pool-RgsConfig.zip"

        Write-Verbose -Message "Exporting $Pool User Data"
        Export-CsUserData -PoolFqdn $Pool -FileName "$BackupPath\$Pool-UserData.zip"
    } # End of foreach ($Pool in $AllUserPools)

} # End of foreach ($Share in $FileShares)