<#
.SYNOPSIS
Creates backup of various components of a Skype for Business environment

.DESCRIPTION
This script will back up various components of a Skype for Business environment
including user data, topology, voice routing information, and response groups settings.

.PARAMETER FileShare
Enter a UNC path to a file share to save the backup data.
Example:  \\SERVER\FILESHARE

.EXAMPLE
Backup-SkypeServer.ps1 -FileShare "\\SERVER\FILESHARE"
This command will back up the Skype environment to \\SERVER\FILESHARE.

.NOTES
Written by Jeff Brown
Jeff@JeffBrown.tech
@JeffWBrown
www.jeffbrown.tech

Any and all technical advice, scripts, and documentation are provided as is with no guarantee.
Always review any code and steps before applying to a production system to understand their full impact.

Version Notes
V1.0 - 2/03/2016 - Initial Version
V1.1 - 3/04/2016 - Removed erroneous $FileShare settings in Variables section
V2.0 - 10/4/2017 - Updated name to Backup-SkypeServer
#>

#*************************************************************************************
#************************       Parameters    ****************************************
#*************************************************************************************

[CmdletBinding()]
Param(
    [Parameter(Position=0, Mandatory=$true, HelpMessage = "Enter a network/UNC path to save the backup files.")]
    [ValidateScript({
        if ((Test-Path -Path $_) -eq $true) {return $true} else {Throw "$_ is not accessible or a valid network path."}
    })]
    [string]$FileShare
)

#*************************************************************************************
#**************************     Variables     ****************************************
#*************************************************************************************

# Get various date fields
[string]$RemoveDate = ((Get-Date).AddDays(-30))
[string]$Now = Get-Date -UFormat %m-%d-%Y_%H%M%S
[string]$BackupPath = "$FileShare\$Now"

# Gets all Front End User Pools in the Environment
[array]$AllUserPools = @((Get-CsService -UserServer).PoolFqdn)

#*************************************************************************************
#**************************     Main Code     ****************************************
#*************************************************************************************

# Remove back ups older than 30 days
# Adjust the variable $RemoveDate to your preference
Write-Verbose -Message "Deleting Backups in $FileShare older than $RemoveDate"
Get-ChildItem -Path $FileShare | Where-Object {$_.LastWriteTime -lt $RemoveDate} | Remove-Item -Recurse

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