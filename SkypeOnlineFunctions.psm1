#Requires -Version 3.0

<#
Written by Jeff Brown
Jeff@JeffBrown.tech
@JeffWBrown
www.jeffbrown.tech
https://github.com/JeffBrownTech

To use the functions in this module, use the Import-Module command followed by the path to this file. For example:
Import-Module C:\files\SkypeOnlineFunctions.psm1
You can also place the .psm1 file in one of the locations PowerShell searches for available modules to import.
These paths can be found in the $env:PSModulePath variable.A common path is C:\Users\<UserID>\Documents\WindowsPowerShell\Modules

Any and all technical advice, scripts, and documentation are provided as is with no guarantee.
Always review any code and steps before applying to a production system to understand their full impact.

Version Notes
V1.3 - 3/11/2018 - Update new license names; added Common Area Phone to license-related commands
V1.2 - 11/11/2017 - Added Remove-SkypeOnlineNormalizationRule cmdlet
V1.1 - 10/8/2017 - Add non-exported helper functions for AzureAD and Skype connections; fixed a few errors
V1.0 - 10/2/2017 - Initial Version
#>

# *** Exported Functions ***

function Add-SkypeOnlineUserLicense {
    <#
.SYNOPSIS
Adds one or more Skype related licenses to a user account.

.DESCRIPTION
Skype for Business Online services are available through assignment of different types of licenses.
This command allows assigning one or more Skype related Office 365 licenses to a user account to enable
the different services, such as E1/E3/E5, Phone System, Calling Plans, and Audio Conferencing.

.PARAMETER Identity
The sign-in address or User Principal Name of the user account to modify.

.PARAMETER AddSkypeStandalone
Adds a Skype for Business Plan 2 license to the user account.

.PARAMETER AddE1
Adds an E1 license to the user account.

.PARAMETER AddE3
Adds an E3 license to the user account.

.PARAMETER AddE5
Adds an E5 license to the user account.

.PARAMETER AddE5NoAudioConferencing
Adds an E5 without Audio Conferencing license to the user account.

.PARAMETER AddAudioConferencing
Adds a Audio Conferencing add-on license to the user account.

.PARAMETER AddPhoneSystem
Adds a Phone System add-on license to the user account.

.PARAMETER AddDomesticCallingPlan
Adds a Domestic Calling Plan add-on license to the user account.

.PARAMETER AddInternationalCallingPlan
Adds an International Calling Plan add-on license to the user account.

.PARAMETER AddCommunicationsCredit
Adds an Communications Credit add-on license to the user account.

.PARAMETER AddCommonAreaPhone
Adds a Common Area Phone license to the user account.

.EXAMPLE
Add-SkypeOnlineUserLicense -Identity Joe@contoso.com -AddE3 -AddPhoneSystem
Example 1 will add the an E3 and Cloud PBX add-on license to Joe@contoso.com

.EXAMPLE
Add-SkypeOnlineUserLicense -Identity Joe@contoso.com -AddE5 -AddDomesticCallingPlan
Example 2 will add the an E5 and Domestic Calling Plan add-on license to Joe@contoso.com

.EXAMPLE
Add-SkypeOnlineUserLicense -Identity Joe@contoso.com -AddSkypeStandalone
Example 3 will add the a Skype for Business Plan 2 license to Joe@contoso.com

.NOTES
The command will test to see if the license exists in the tenant as well as if the user already
has the licensed assigned. It does not keep track or take into account the number of licenses
available before attempting to assign the license.
#>
    [CmdletBinding(DefaultParameterSetName = 'AddDomestic')]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'AddDomestic')]
        [Parameter(ParameterSetName = 'AddInternational')]
        [Alias("UPN", "UserPrincipalName", "Username")]
        [string[]]$Identity,

        [Parameter(ParameterSetName = 'AddDomestic')]
        [Parameter(ParameterSetName = 'AddInternational')]
        [switch]$AddSkypeStandalone,

        [Parameter(ParameterSetName = 'AddDomestic')]
        [Parameter(ParameterSetName = 'AddInternational')]
        [switch]$AddE1,

        [Parameter(ParameterSetName = 'AddDomestic')]
        [Parameter(ParameterSetName = 'AddInternational')]
        [switch]$AddE3,

        [Parameter(ParameterSetName = 'AddDomestic')]
        [Parameter(ParameterSetName = 'AddInternational')]
        [switch]$AddE5,

        [Parameter(ParameterSetName = 'AddDomestic')]
        [Parameter(ParameterSetName = 'AddInternational')]
        [switch]$AddE5NoAudioConferencing,

        [Parameter(ParameterSetName = 'AddDomestic')]
        [Parameter(ParameterSetName = 'AddInternational')]
        [switch]$AddAudioConferencing,

        [Parameter(ParameterSetName = 'AddDomestic')]
        [Parameter(ParameterSetName = 'AddInternational')]
        [switch]$AddPhoneSystem,

        [Parameter(ParameterSetName = 'AddDomestic')]
        [switch]$AddDomesticCallingPlan,

        [Parameter(ParameterSetName = 'AddInternational')]
        [switch]$AddInternationalCallingPlan,

        [Parameter(ParameterSetName = 'AddDomestic')]
        [Parameter(ParameterSetName = 'AddInternational')]
        [switch]$AddCommunicationsCredit,    

        [Parameter(ParameterSetName = 'AddDomestic')]
        [Parameter(ParameterSetName = 'AddInternational')]
        [switch]$AddCommonAreaPhone
    )

    BEGIN {
        <#
        # Verify Azure Active Directory PowerShell Module (AzureAD) is installed on the PC
        if ((Get-Module -ListAvailable).Name -notcontains "AzureAD")
        {
            Write-Warning -Message "Azure Active Directory PowerShell module is not installed. Please install and run command again."
            RETURN
        }
        
        # Attempt to get tenant Account SKUs
        # If error is found, attempt connection to Azure AD using Connect-AzureAD
        $tenantSKUs = Get-AzureADSubscribedSku -ErrorAction SilentlyContinue -ErrorVariable getAzureADTenantDetail
        if ($getAzureADTenantDetail)
        {
            Write-Warning -Message "You must connect to Azure AD before continuing"
            Connect-AzureAD -Credential (Get-Credential -Message "Enter the sign-in name and password for an O365 Global Admin")
            $tenantSKUs = Get-AzureADSubscribedSku -ErrorAction STOP
        }
        #>
                
        if ((TestAzureADModule) -eq $false) {RETURN}

        if ((TestAzureADConnection) -eq $false) {
            try {
                Connect-AzureAD -ErrorAction STOP | Out-Null
            }
            catch {
                Write-Warning $_
                CONTINUE
            }
        }

        try {
            $tenantSKUs = Get-AzureADSubscribedSku -ErrorAction STOP
        }
        catch {
            Write-Warning $_
            RETURN
        }

        # Build Skype SKU Variables from Available Licenses in the Tenant
        foreach ($tenantSKU in $tenantSKUs) {
            switch ($tenantSKU.SkuPartNumber) {
                "MCOPSTN1" {$DomesticCallingPlan = $tenantSKU.SkuId; break}
                "MCOPSTN2" {$InternationalCallingPlan = $tenantSKU.SkuId; break}
                "MCOMEETADV" {$AudioConferencing = $tenantSKU.SkuId; break}
                "MCOEV" {$PhoneSystem = $tenantSKU.SkuId; break}
                "ENTERPRISEPREMIUM" {$E5WithPhoneSystem = $tenantSKU.SkuId; break}
                "ENTERPRISEPREMIUM_NOPSTNCONF" {$E5NoAudioConferencing = $tenantSKU.SkuId; break}
                "ENTERPRISEPACK" {$E3 = $tenantSKU.SkuId; break}
                "STANDARDPACK" {$E1 = $tenantSKU.SkuId; break}
                "MCOSTANDARD" {$SkypeStandalonePlan = $tenantSKU.SkuId; break}
                "MCOPSTNC" {$CommunicationsCredit = $tenantSKU.SkuId; break}
                "MCOCAP" {$CommonAreaPhone = $tenantSKU.SkuId; break}
            } # End of switch statement
        } # End of foreach $tenantSKUs
    } # End of BEGIN

    PROCESS {
        foreach ($ID in $Identity) {
            try {
                Get-AzureADUser -ObjectId $ID -ErrorAction STOP | Out-Null
            }
            catch {
                $output = GetActionOutputObject2 -Name $ID -Result "Not a valid user account"
                Write-Output $output
                continue
            }

            # Get user's currently assigned licenses
            $userCurrentLicenses = (Get-AzureADUserLicenseDetail -ObjectId $ID).SkuId

            # Skype Standalone Plan
            if ($AddSkypeStandalone -eq $true) {
                if ($null -ne $SkypeStandalonePlan) {
                    if ($userCurrentLicenses -notcontains $SkypeStandalonePlan) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $SkypeStandalonePlan -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $SkypeStandalonePlan
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP                                
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: Skype Standalone Plan license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign Skype Standalone Plan license: $_"
                        }
                    }
                    else {
                        $output = GetActionOutputObject2 -Name $ID -Result "INFO: User already assigned Skype Standalone Plan"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No Skype Standalone Plan licenses found in tenant"
                }

                Write-Output $output
            }

            # E1
            if ($AddE1 -eq $true) {
                if ($null -ne $E1) {
                    if ($userCurrentLicenses -notcontains $E1) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $E1 -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $E1
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: E1 license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign E1 license: $_"
                        }
                    }
                    else {
                        $output = GetActionOutputObject2 -Name $ID -Result "INFO: User already assigned E1"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No E1 licenses found in tenant"
                }

                Write-Output $output
            }

            # E3
            if ($AddE3 -eq $true) {
                # Verify if E3 licenses exist in tenant
                if ($null -ne $E3) {
                    # Verify if user does not have the license assigned
                    if ($userCurrentLicenses -notcontains $E3) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $E3 -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $E3
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: E3 license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign E3 license: $_"
                        }
                    }
                    else {
                        $output = Get-ActionOutputObject -Name $ID -Result "INFO: User already assigned E3"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No E3 licenses found in tenant"
                }

                Write-Output $output
            }

            # E5 with Phone System
            if ($AddE5 -eq $true) {
                if ($null -ne $E5WithPhoneSystem) {
                    if ($userCurrentLicenses -notcontains $E5WithPhoneSystem) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $E5WithPhoneSystem -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $E5WithPhoneSystem
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: E5 license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign E5 license: $_"
                        }
                    }
                    else {
                        $output = GetActionOutputObject2 -Name $ID -Result "INFO: User already assigned E5"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No E5 licenses found in tenant"
                }

                Write-Output $output
            }

            # E5 No PSTN Conferencing
            if ($AddE5NoAudioConferencing -eq $true) {
                if ($null -ne $E5NoAudioConferencing) {
                    if ($userCurrentLicenses -notcontains $E5NoAudioConferencing) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $E5NoAudioConferencing -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $E5NoAudioConferencing
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: E5 without Audio Conferencing license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign E5 without Audio Conferencing license: $_"
                        }
                    }
                    else {
                        $output = GetActionOutputObject2 -Name $ID -Result "INFO: User already assigned E5 without Audio Conferencing"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No E5 without Audio Conferencing licenses found in tenant"
                }

                Write-Output $output
            }

            # Audio Conferencing Add-On
            if ($AddAudioConferencing -eq $true) {
                # Checking to see if $AudioConferencing exists
                if ($null -ne $AudioConferencing) {
                    if ($userCurrentLicenses -notcontains $AudioConferencing) {
                        # Adding Audio Conferencing SKU to user account
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $AudioConferencing -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $AudioConferencing
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: Audio Conferencing add-on license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign Audio Conferencing add-on license: $_"
                        }
                    }
                    else {
                        $output = GetActionOutputObject2 -Name $ID -Result "INFO: User already assigned Audio Conferencing"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No Audio Conferencing add-on licenses found in tenant"
                }

                # Output results of AudioConferencing assignment
                Write-Output $output
            }

            # Phone System Add-On
            if ($AddPhoneSystem -eq $true) {
                if ($null -ne $PhoneSystem) {
                    if ($userCurrentLicenses -notcontains $PhoneSystem) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $PhoneSystem -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $PhoneSystem
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: Phone System add-on license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign Phone System add-on license: $_"
                        }
                    }
                    else {
                        $output = GetActionOutputObject2 -Name $ID -Result "INFO: User already assigned Phone System add-on"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No Phone System add-on licenses found in tenant"
                }

                Write-Output $output
            }

            # Domestic Calling Plan
            if ($AddDomesticCallingPlan -eq $true) {
                if ($null -ne $DomesticCallingPlan) {
                    if ($userCurrentLicenses -notcontains $DomesticCallingPlan) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $DomesticCallingPlan -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $DomesticCallingPlan
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: Domestic Calling Plan license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign Domestic Calling Plan license: $_"
                        }
                    }
                    else {
                        $output = GetActionOutputObject2 -Name $ID -Result "INFO: User already assigned Domestic Calling Plan"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No Domestic Calling Plan licenses found in tenant"
                }

                Write-Output $output
            }

            # Domestic & International Calling Plan
            if ($AddInternationalCallingPlan -eq $true) {
                if ($null -ne $InternationalCallingPlan) {
                    if ($userCurrentLicenses -notcontains $InternationalCallingPlan) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $InternationalCallingPlan -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $InternationalCallingPlan
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: International Calling Plan license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign International Calling Plan license: $_"
                        }
                    }
                    else {
                        $output = GetActionOutputObject2 -Name $ID -Result "INFO: User already assigned International Calling Plan"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No International Calling Plan licenses found in tenant"
                }

                Write-Output $output
            }

            # Communications Credit
            if ($AddCommunicationsCredit -eq $true) {
                if ($null -ne $CommunicationsCredit) {
                    if ($userCurrentLicenses -notcontains $CommunicationsCredit) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $CommunicationsCredit -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $CommunicationsCredit
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: Communications Credit license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign Communications Credit license: $_"
                        }
                    }
                    else {
                        $output = GetActionOutputObject2 -Name $ID -Result "INFO: User already assigned Communications Credit License"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No Communications Credit licenses found in tenant"
                }

                Write-Output $output
            }

            # Common Area Phone
            if ($AddCommonAreaPhone -eq $true) {
                if ($null -ne $CommonAreaPhone) {
                    if ($userCurrentLicenses -notcontains $CommonAreaPhone) {
                        try {
                            #Set-MsolUserLicense -UserPrincipalName $ID -AddLicenses $CommonAreaPhone -ErrorAction STOP
                            $license = NewLicenseObject -SkuId $CommonAreaPhone
                            Set-AzureADUserLicense -ObjectId $ID -AssignedLicenses $license -ErrorAction STOP
                            $output = GetActionOutputObject2 -Name $ID -Result "SUCCESS: Common Area Phone license assigned"
                        }
                        catch {
                            $output = GetActionOutputObject2 -Name $ID -Result "ERROR: Unable to assign Common Area Phone license: $_"
                        }
                    }
                    else {
                        $output = GetActionOutputObject2 -Name $ID -Result "INFO: User already assigned Common Area Phone License"
                    }
                }
                else {
                    $output = GetActionOutputObject2 -Name $ID -Result "WARNING: No Common Area Phone licenses found in tenant"
                }

                Write-Output $output
            }
        } # End of foreach ($ID in $Identity)
    } # End of PROCESS
} # End of Add-SkypeOnlineUserLicense

function Connect-SkypeOnline
{
<#
.SYNOPSIS
Creates a remote PowerShell session out to Skype for Business Online.

.DESCRIPTION
Connecting to a remote PowerShell session to Skype for Business Online requires several components
and steps. This function consolidates those activities by 1) verifying the SkypeOnlineConnector is
installed and imported, 2) prompting for username and password to make and to import the session.

.PARAMETER UserName
The username or sign-in address to use when making the remote PowerShell session connection.

.EXAMPLE
Connect-SkypeOnline

Example 1 will prompt for the username and password of an administrator with permissions to connect to Skype for Business Online.

.EXAMPLE
Connect-SkypeOnline -UserName admin@contoso.com

Example 2 will prefill the authentication prompt with admin@contoso.com and only ask for the password for the account to connect out to Skype for Business Online.

.NOTES
Requires that the Skype Online Connector PowerShell module be installed.
#>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$UserName         
    )
    
    if ((TestSkypeOnlineModule) -eq $true)
    {
        if ((TestSkypeOnlineConnection) -eq $false)
        {
            $moduleVersion = (Get-Module -Name SkypeOnlineConnector).Version
            if ($moduleVersion.Major -le "6") # Version 6 and lower do not support MFA authentication for Skype Module PowerShell; also allows use of older PSCredential objects
            {
                try
                {
                    $SkypeOnlineSession = New-CsOnlineSession -Credential (Get-Credential $UserName -Message "Enter the sign-in address and password of a Global or Skype for Business Admin") -ErrorAction STOP
                    Import-Module (Import-PSSession -Session $SkypeOnlineSession -AllowClobber -ErrorAction STOP) -Global
                }
                catch
                {
                    $errorMessage = $_
                    if ($errorMessage -like "*Making sure that you have used the correct user name and password*")
                    {
                        Write-Warning -Message "Logon failed. Please try again and make sure that you have used the correct user name and password."
                    }                    
                    elseif ($errorMessage -like "*Please create a new credential object*")
                    {
                        Write-Warning -Message "Logon failed. This may be due to multi-factor being enabled for the user account and not using the latest Skype for Business Online PowerShell module."
                    }
                    else
                    {
                        Write-Warning -Message $_
                    }
                }
            }
            else # This should be all newer version than 6; does not support PSCredential objects but supports MFA
            {
                try
                {
                    if ($PSBoundParameters.ContainsKey("UserName"))
                    {
                        $SkypeOnlineSession = New-CsOnlineSession $UserName -ErrorAction STOP
                    }
                    else
                    {
                        $SkypeOnlineSession = New-CsOnlineSession -ErrorAction STOP
                    }

                    Import-Module (Import-PSSession -Session $SkypeOnlineSession -AllowClobber -ErrorAction STOP) -Global
                }
                catch
                {
                    Write-Warning -Message $_
                }
            } # End of if statement for module version checking
        }
        else
        {
            Write-Warning -Message "A Skype Online PowerShell Sessions already exists. Please run Disconnect-SkypeOnline before attempting this command again."
        } # End checking for existing Skype Online Connection
    }
    else
    {
        Write-Warning -Message "Skype Online PowerShell Connector module is not installed. Please install and try again."
        Write-Warning -Message "The module can be downloaded here: https://www.microsoft.com/en-us/download/details.aspx?id=39366"
    } # End of testing module existence
} # End of Connect-SkypeOnline

# Work in Progress - Currently not in list of exported functions
function Connect-SkypeOnlineMultiForest
{
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$UserName,

        [Parameter()]
        [ValidateSet("APC","AUS","CAN","EUR","IND","JPN","NAM")]
        [string]$Region
    )

    if ((Get-Module).Name -notcontains "SkypeOnlineConnector")
    {
        try
        {
            Import-Module SkypeOnlineConnector -ErrorAction STOP
        }
        catch
        {
            Write-Error -Message "Unable to import SkypeOnlineConnector PowerShell Module : $_"
        }
    }
    
    if ((Get-PsSession).ComputerName -notlike "*.online.lync.com")
    {
        try
        {        
            $SkypeOnlineCredentials = Get-Credential $UserName -Message "Enter the sign-in address and password of an O365 or Skype Online Admin"
            
            if ($Region.Length -gt 0)
            {
                switch ($Region)
                {
                    "APC" {$forestCode = "0F"; break}
                    "AUS" {$forestCode = "AU1"; break}
                    "CAN" {$forestCode = "CA1"; break}
                    "EUR" {$forestCode = "1E"; break}
                    "IND" {$forestCode = "IN1"; break}
                    "JPN" {$forestCode = "JP1"; break}
                    "NAM" {$forestCode = "2A"; break}
                }
                
                $SkypeOnlineSession = New-CsOnlineSession -Credential $SkypeOnlineCredentials -OverridePowershellUri "https://admin$forestCode.online.lync.com/OcsPowershellLiveId" -Verbose -ErrorAction STOP
            }
            else
            {
                $SkypeOnlineSession = New-CsOnlineSession -Credential $SkypeOnlineCredentials -Verbose -ErrorAction STOP
            }

            Import-PSSession -Session $SkypeOnlineSession -AllowClobber -Verbose -ErrorAction STOP
        }
        catch
        {
            Write-Warning -Message $_
        }
    }
    else
    {
        Write-Warning -Message "Existing Skype Online PowerShell Sessions Exists"
    }
}

function Disconnect-SkypeOnline
{
<#
.SYNOPSIS
Disconnects any current Skype for Business Online remote PowerShell sessions and removes any imported modules.

.EXAMPLE
Disconnect-SkypeOnline
Example 1 will remove any current Skype for Business Online remote PowerShell sessions and removes any imported modules.
#>

    [CmdletBinding()]
    param()

    [bool]$sessionFound = $false

    $PSSesssions = Get-PSSession

    foreach ($session in $PSSesssions)
    {
        if ($session.ComputerName -like "*.online.lync.com")
        {
            $sessionFound = $true
            Remove-PSSession $session
        }
    }

    Get-Module | Where-Object {$_.Description -like "*.online.lync.com*"} | Remove-Module

    if ($sessionFound -eq $false)
    {
        Write-Warning -Message "No remote PowerShell sessions to Skype Online currently exist"
    }

} # End of Disconnect-SkypeOnline

function Get-SkypeOnlineConferenceDialInNumbers
{
<#
.SYNOPSIS
Gathers the audio conference dial-in numbers information for a Skype for Business Online tenant.

.DESCRIPTION
This command uses the tenant's conferencing dial-in number web page to gather a "user-readable" list of
the regions, numbers, and available languages where dial-in conferencing numbers are available. This web
page can be access at https://dialin.lync.com/DialInOnline/Dialin.aspx?path=<DOMAIN> replacing "<DOMAIN>"
with the tenant's default domain name (i.e. contoso.com).

.PARAMETER Domain
The Skype for Business Online Tenant domain to gather the conference dial-in numbers.

.EXAMPLE
Get-SkypeOnlineConferenceDialInNumbers -Domain contoso.com
Example 1 will gather the conference dial-in numbers for contoso.com based on their conference dial-in number web page.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,HelpMessage="Enter the domain name to gather the available conference dial-in numbers")]
        [string]$Domain
    )

    try
    {
        $siteContents = Invoke-WebRequest https://webdir1a.online.lync.com/DialinOnline/Dialin.aspx?path=$Domain -ErrorAction STOP
    }
    catch
    {
        Write-Warning -Message "Unable to access that dial-in page. Please check the domain name and try again. Also try to manually navigate to the page using the URL http://dialin.lync.com/DialInOnline/Dialin.aspx?path=$Domain."
        RETURN
    }

    $tables = $siteContents.ParsedHtml.getElementsByTagName("TABLE")
    $table = $tables[0]
    $rows = @($table.rows)

    $output = [PSCustomObject][ordered]@{
        Location = $null
        Number = $null
        Languages = $null
    }

    for ($n = 0; $n -lt $rows.Count; $n += 1)
    {
        if ($rows[$n].innerHTML -like "<TH*")
        {
            $output.Location = $rows[$n].innerText
        }
        else
        {
            $output.Number = $rows[$n].cells[0].innerText
            $output.Languages = $rows[$n].cells[1].innerText
            Write-Output $output
        }
    }
} # End of Get-SkypeOnlineConferenceDialInNumbers

function Get-SkypeOnlineUserLicense
{
<#
.SYNOPSIS
Gathers licenses assigned to a Skype for Business Online user for Cloud PBX and PSTN Calling Plans.

.DESCRIPTION
This script lists the UPN, Name, currently O365 Plan, Calling Plan, Communication Credit, and Audio Conferencing Add-On License

.PARAMETER Identity
The Identity/UPN/sign-in address for the user entered in the format <name>@<domain>.
Aliases include: "UPN","UserPrincipalName","Username"

.EXAMPLE
.\Get-SkypeOnlineLicense.ps1 -Identity John@domain.com

Example 1 will confirm the license for a single user: John@domain.com

.EXAMPLE
.\Get-SkypeOnlineLicense.ps1 -Identity John@domain.com,Jane@domain.com

Example 2 will confirm the licenses for two users: John@domain.com & Jane@domain.com

.EXAMPLE
Import-Csv User.csv | .\Get-SkypeOnlineLicense.ps1

Example 3 will use a CSV as an input file and confirm the licenses for users listed in the file. The input file must
have a single column heading of "Identity" with properly formatted UPNs.

.NOTES
If using a CSV file for pipeline input, the CSV user data file should contain a column name matching each of this script's parameters. Example:

Identity
John@domain.com
Jane@domain.com

Output can be redirected to a file or grid-view.
#>


    [CmdletBinding()]
    param(

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = "Enter the UPN or login name of the user account, typically <user>@<domain>.")]
            [Alias("UPN","UserPrincipalName","Username")]
            [string[]]$Identity
    )

    BEGIN
    {
        if ((TestAzureADModule) -eq $false) {RETURN}

        if ((TestAzureADConnection) -eq $false)
        {
            try
            {
                Connect-AzureAD -ErrorAction STOP
            }
            catch
            {
                Write-Warning $_
                CONTINUE
            }
        }
    } # End of BEGIN

    PROCESS
    {
        foreach ($User in $Identity)
        {
            try
            {
                Get-AzureADUser -ObjectId $User -ErrorAction STOP | Out-Null
            }
            catch
            {
                $output = [PSCustomObject][ordered]@{
                    User = $User
                    License = "Invalid User"
                    CallingPlan = "Invalid User"
                    CommunicationsCreditLicense = "Invalid User"
                    AudioConferencingAddOn = "Invalid User"
                    CommoneAreaPhoneLicense = "Invalid User"
                }

                Write-Output $output
                continue
            }
                        
            $userInformation = Get-AzureADUser -ObjectId $User
            $assignedLicenses = (Get-AzureADUserLicenseDetail -ObjectId $User).SkuPartNumber
            [string]$DisplayName = $userInformation.Surname + ", " + $userInformation.GivenName
            [string]$O365License = $null
            [string]$currentCallingPlan = "Not Assigned"
            [bool]$CommunicationsCreditLicense = $false
            [bool]$AudioConferencingAddOn = $false
            [bool]$CommonAreaPhoneLicense = $false

            if ($null -ne $assignedLicenses)
            {
                foreach ($license in $assignedLicenses)
                {
                    switch -Wildcard ($license)
                    {
                        "DESKLESSPACK" {$O365License += "Kiosk Plan, ";break}
                        "EXCHANGEDESKLESS" {$O365License += "Exchange Kiosk, "; break}
                        "EXCHANGESTANDARD" {$O365License += "Exchange Standard, "; break}
                        "EXCHANGEENTERPRISE" {$O365License += "Exchange Premium, "; break}
                        "MCOSTANDARD" {$O365License += "Skype Plan 2, "; break}
                        "STANDARDPACK" {$O365License += "E1, "; break}
                        "ENTERPRISEPACK" {$O365License += "E3, "; break}
                        "ENTERPRISEPREMIUM" {$O365License += "E5, "; break}
                        "ENTERPRISEPREMIUM_NOPSTNCONF" {$O365License += "E5 (No Audio Conferencing), "; break}
                        "MCOCAP" {$CommonAreaPhoneLicense = $true; break}
                        "MCOPSTN1" {$currentCallingPlan = "Domestic"; break}
                        "MCOPSTN2" {$currentCallingPlan = "Domestic & International"; break}
                        "MCOPSTNC" {$CommunicationsCreditLicense = $true; break}
                        "MCOMEETADV" {$AudioConferencingAddOn = $true; break}
                    }
                }
            }
            else
            {
                $O365License = "No Licenses Assigned"
            }
            
            $output = [PSCustomObject][ordered]@{
                User                        = $User                
                License                     = $O365License.TrimEnd(", ") # Removes any trailing ", " at the end of the string                
                CallingPlan                 = $currentCallingPlan                
                CommunicationsCreditLicense = $CommunicationsCreditLicense
                AudioConferencingAddOn      = $AudioConferencingAddOn
                CommoneAreaPhoneLicense     = $CommonAreaPhoneLicense                
            }

            Write-Output $output
        } # End of foreach ($UserPrincipal in $Identity)
    } # End of PROCESS
} # End of Get-SkypeOnlineUserLicense

function Get-SkypeOnlineTenantLicenses
{
<#
.SYNOPSIS
Displays the Skype individual plans, add-on & grouped license SKUs for a tenant.

.DESCRIPTION
Skype for Business Online services can be provisioned through several different combinations of individual
plans as well as add-on and grouped license SKUs. This command displays these license SKUs in a more friendly
format with descriptive names, active, consumed, remaining, and expiring licenses.

.EXAMPLE
Get-SkypeOnlineTenantLicenses
Example 1 will display all the Skype related licenses for the tenant.

.NOTES
Requires the Azure Active Directory PowerShell module to be installed and authenticated to the tenant's Azure AD instance.
#>

    [CmdletBinding()]
    param()
        
    if ((TestAzureADModule) -eq $false) {RETURN}

    if ((TestAzureADConnection) -eq $false)
    {
        try
        {
            Connect-AzureAD -ErrorAction STOP | Out-Null
        }
        catch
        {
            Write-Warning $_
            CONTINUE
        }
    }

    try
    {
        $tenantSKUs = Get-AzureADSubscribedSku -ErrorAction STOP
    }
    catch
    {
        Write-Warning $_
        RETURN
    }

    foreach ($tenantSKU in $tenantSKUs)
    {
        [string]$skuFriendlyName = $null
        switch ($tenantSKU.SkuPartNumber)
        {
            "MCOPSTN1" {$skuFriendlyName = "Domestic Calling Plan"; break}
            "MCOPSTN2" {$skuFriendlyName = "Domestic and International Calling Plan"; break}
            "MCOPSTNC" {$skuFriendlyName = "Communications Credit Add-On"; break}
            "MCOMEETADV" {$skuFriendlyName = "Audio Conferencing Add-On"; break}
            "MCOEV" {$skuFriendlyName = "Phone System Add-On"; break}
            "MCOCAP" {$skuFriendlyName = "Common Area Phone"; break}
            "ENTERPRISEPREMIUM" {$skuFriendlyName = "Enterprise E5 with Phone System"; break}
            "ENTERPRISEPREMIUM_NOPSTNCONF" {$skuFriendlyName = "Enterprise E5 Without Audio Conferencing"; break}
            "ENTERPRISEPACK" {$skuFriendlyName = "Enterprise E3"; break}
            "STANDARDPACK" {$skuFriendlyName = "Enterprise E1"; break}
            "MCOSTANDARD" {$skuFriendlyName = "Skype for Business Online Standalone Plan 2"; break}
            "O365_BUSINESS_PREMIUM" {$skuFriendlyName = "O365 Business Premium"; break}
            "PHONESYSTEM_VIRTUALUSER" {$skuFriendlyName = "Phone System - Virtual User"; break}
        }
        
        if ($skuFriendlyName.Length -gt 0)
        {
            [PSCustomObject][ordered]@{
                License = $skuFriendlyName
                Available = $tenantSKU.PrepaidUnits.Enabled
                Consumed = $tenantSKU.ConsumedUnits
                Remaining = $($tenantSKU.PrepaidUnits.Enabled - $tenantSKU.ConsumedUnits)
                Expiring = $tenantSKU.PrepaidUnits.Warning
            }
        }    
    } # End of foreach ($tenantSKU in $tenantSKUs}
} # End of Get-SkypeOnlineTenantLicenses

function Remove-SkypeOnlineNormalizationRule
{
<#
.SYNOPSIS
Removes a normalization rule from a tenant dial plan.

.DESCRIPTION
This command will display the normalization rules for a tenant dial plan in a list with
index numbers. After choosing one of the rule index numbers, the rule will be removed from
the tenant dial plan. This command requires a remote PowerShell session to Skype for Business Online.

.PARAMETER DialPlan
This is the name of a valid dial plan for the tenant. To view available tenant dial plans,
use the command Get-CsTenantDialPlan.

.EXAMPLE
Remove-SkypeOnlineNormalizationRule -DialPlan US-OK-OKC-DialPlan
Example 1 will display the availble normalization rules to remove from dial plan US-OK-OKC-DialPlan.

.NOTES
The dial plan rules will display in format similar the example below:

RuleIndex Name            Pattern    Translation
--------- ----            -------    -----------
        0 Intl Dialing    ^011(\d+)$ +$1
        1 Extension Rule  ^(\d{5})$  +155512$1
        2 Long Distance   ^1(\d+)$   +1$1
        3 Default         ^(\d+)$    +1$1
#>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the name of the dial plan to modify the normalization rules.")]
        [string]$DialPlan
    )

    if ((TestSkypeOnlineModule) -eq $true)
    {
        if ((TestSkypeOnlineConnection) -eq $false)
        {
            Write-Warning -Message "You must create a remote PowerShell session to Skype Online before continuing."
            Connect-SkypeOnline
        }
    }
    else
    {
        Write-Warning -Message "Skype Online PowerShell Connector module is not installed. Please install and try again."
        Write-Warning -Message "The module can be downloaded here: https://www.microsoft.com/en-us/download/details.aspx?id=39366"
    }

    $dpInfo = Get-CsTenantDialPlan -Identity $DialPlan -ErrorAction SilentlyContinue

    if ($null -ne $dpInfo)
    {
        $currentNormRules = $dpInfo.NormalizationRules
        [int]$ruleIndex = 0
        [int]$ruleCount = $currentNormRules.Count
        [array]$ruleArray = @()
        [array]$indexArray = @()

        if ($ruleCount -ne 0)
        {
            foreach ($normRule in $dpInfo.NormalizationRules)
            {
                $output = [PSCustomObject][ordered]@{
                    'RuleIndex' = $ruleIndex
                    'Name' = $normRule.Name
                    'Pattern' = $normRule.Pattern
                    'Translation' = $normRule.Translation
                }

                $ruleArray += $output
                $indexArray += $ruleIndex
                $ruleIndex++
            } # End of foreach ($normRule in $dpInfo.NormalizationRules)

            # Displays rules to the screen with RuleIndex added
            $ruleArray | Out-Host

            do
            {
                $indexToRemove = Read-Host -Prompt "Enter the Rule Index of the normalization rule to remove from the dial plan (leave blank to quit without changes)"
                
                if ($indexToRemove -notin $indexArray -and $indexToRemove.Length -ne 0)
                {
                    Write-Warning -Message "That is not a valid Rule Index. Please try again or leave blank to quit."
                }
            } until ($indexToRemove -in $indexArray -or $indexToRemove.Length -eq 0)

            if ($indexToRemove.Length -eq 0) {RETURN}

            # If there is more than 1 rule left, remove the rule and set to new normalization rules
            # If there is only 1 rule left, we have to set -NormalizationRules to $null
            if ($ruleCount -ne 1)
            {
                $newNormRules = $currentNormRules
                $newNormRules.Remove($currentNormRules[$indexToRemove])
                Set-CsTenantDialPlan -Identity $DialPlan -NormalizationRules $newNormRules
            }
            else
            {
                Set-CsTenantDialPlan -Identity $DialPlan -NormalizationRules $null
            }
        }
        else
        {
            Write-Warning -Message "$DialPlan does not contain any normalization rules."
        }
    }
    else
    {
        Write-Warning -Message "$DialPlan is not a valid dial plan for the tenant. Please try again."
    }
} # End of Remove-SkypeOnlineNormalizationRule

function Set-SkypeOnlineUserPolicy
{
<#
.SYNOPSIS
Sets policies on a Skype for Business Online user

.DESCRIPTION
Skype for Business Online offers the assignment of several policies to control client, conferencing,
external access, and mobility options. Typically these are assigned using different commands, but
Set-SkypeOnlineUserPolicy allows settings all these with a single command. One or all policy options can
be used during assignment.

.PARAMETER Identity
This is the sign-in address/User Principal Name of the user to configure.

.PARAMETER ClientPolicy
This is the Client Policy to assign to the user.

.PARAMETER ConferencingPolicy
This is the Conferencing Policy to assign to the user.

.PARAMETER ExternalAccessPolicy
This is the External Access Policy to assign to the user.

.PARAMETER MobilityPolicy
This is the Mobility Policy to assign to the user.

.EXAMPLE
Set-SkypeOnlineUserPolicy -Identity John.Doe@contoso.com -ClientPolicy ClientPolicyNoIMURL
Example 1 will set the user John.Does@contoso.com with a client policy.

.EXAMPLE
Set-SkypeOnlineUserPolicy -Identity John.Doe@contoso.com -ClientPolicy ClientPolicyNoIMURL -ConferencingPolicy BposSAllModalityNoFT
Example 2 will set the user John.Does@contoso.com with a client and conferencing policy.

.EXAMPLE
Set-SkypeOnlineUserPolicy -Identity John.Doe@contoso.com -ClientPolicy ClientPolicyNoIMURL -ConferencingPolicy BposSAllModalityNoFT -ExternalAccessPolicy FederationOnly -MobilityPolicy
Example 3 will set the user John.Does@contoso.com with a client, conferencing, external access, and mobility policy.
#>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Enter the identity for the user to configure")]
        [Alias("UPN","UserPrincipalName","Username")]
        [string[]]$Identity,
        
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$ClientPolicy,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$ConferencingPolicy,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$ExternalAccessPolicy,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$MobilityPolicy
    )

    BEGIN
    {
        if ((TestSkypeOnlineModule) -eq $true)
        {
            if ((TestSkypeOnlineConnection) -eq $false)
            {
                Write-Warning -Message "You must create a remote PowerShell session to Skype Online before continuing."
                Connect-SkypeOnline
            }
        }
        else
        {
            Write-Warning -Message "Skype Online PowerShell Connector module is not installed. Please install and try again."
            Write-Warning -Message "The module can be downloaded here: https://www.microsoft.com/en-us/download/details.aspx?id=39366"
        }

        # Get available policies for tenant
        Write-Verbose -Message "Gathering all policies for tenant"
        $tenantClientPolicies = (Get-CsClientPolicy -WarningAction SilentlyContinue).Identity
        $tenantConferencingPolicies = (Get-CsConferencingPolicy -Include SubscriptionDefaults -WarningAction SilentlyContinue).Identity
        $tenantExternalAccessPolicies = (Get-CsExternalAccessPolicy -WarningAction SilentlyContinue).Identity
        $tenantMobilityPolicies = (Get-CsMobilityPolicy -WarningAction SilentlyContinue).Identity
    } # End of BEGIN

    PROCESS
    {
        foreach ($ID in $Identity)
        {
            # User Validation
            # Validating users in a try/catch block does not catch the error properly and does not allow for custom outputting of an error message
            if ($null -ne (Get-CsOnlineUser -Identity $ID -ErrorAction SilentlyContinue))
            {
                # Client Policy
                if ($PSBoundParameters.ContainsKey("ClientPolicy"))
                {
                    # Verify if $ClientPolicy is a valid policy to assign
                    if ($tenantClientPolicies -icontains "Tag:$ClientPolicy")
                    {
                        try
                        {
                            # Attempt to assign policy
                            Grant-CsClientPolicy -Identity $ID -PolicyName $ClientPolicy -WarningAction SilentlyContinue -ErrorAction STOP
                            $output = GetActionOutputObject3 -Name $ID -Property "Client Policy" -Result "Success: $ClientPolicy"
                        }
                        catch
                        {
                            $errorMessage = $_
                            $output = GetActionOutputObject3 -Name $ID -Property "Client Policy" -Result "Error: $errorMessage"
                        }
                    }
                    else
                    {
                        # Output invalid client policy to error log file
                        $output = GetActionOutputObject3 -Name $ID -Property "Client Policy" -Result "Error: $ClientPolicy is not valid or does not exist"
                    }

                    # Output final ClientPolicy Success or Fail message
                    Write-Output -InputObject $output
                } # End of setting Client Policy

                # Conferencing Policy
                if ($PSBoundParameters.ContainsKey("ConferencingPolicy"))
                {
                    # Verify if $ConferencingPolicy is a valid policy to assign
                    if ($tenantConferencingPolicies -icontains "Tag:$ConferencingPolicy")
                    {
                        try
                        {
                            # Attempt to assign policy
                            Grant-CsConferencingPolicy -Identity $ID -PolicyName $ConferencingPolicy -WarningAction SilentlyContinue -ErrorAction STOP
                            $output = GetActionOutputObject3 -Name $ID -Property "Conferencing Policy" -Result "Success: $ConferencingPolicy"
                        }
                        catch
                        {
                            # Output to error log file on policy assignment error
                            $errorMessage = $_
                            $output = GetActionOutputObject3 -Name $ID -Property "Conferencing Policy" -Result "Error: $errorMessage"
                        }
                    }
                    else
                    {
                        # Output invalid conferencing policy to error log file
                        $output = GetActionOutputObject3 -Name $ID -Property "Conferencing Policy" -Result "Error: $ConferencingPolicy is not valid or does not exist"
                    }

                    # Output final ConferencingPolicy Success or Fail message
                    Write-Output -InputObject $output
                } # End of setting Conferencing Policy
    
                # External Access Policy
                if ($PSBoundParameters.ContainsKey("ExternalAccessPolicy"))
                {
                    # Verify if $ExternalAccessPolicy is a valid policy to assign
                    if ($tenantExternalAccessPolicies -icontains "Tag:$ExternalAccessPolicy")
                    {
                        try
                        {
                            # Attempt to assign policy
                            Grant-CsExternalAccessPolicy -Identity $ID -PolicyName $ExternalAccessPolicy -WarningAction SilentlyContinue -ErrorAction STOP
                            $output = GetActionOutputObject3 -Name $ID -Property "External Access Policy" -Result "Success: $ExternalAccessPolicy"
                        }
                        catch
                        {
                            $errorMessage = $_                            
                            $output = GetActionOutputObject3 -Name $ID -Property "External Access Policy" -Result "Error: $errorMessage"
                        }
                    }
                    else
                    {
                        # Output invalid external access policy to error log file
                        $output = GetActionOutputObject3 -Name $ID -Property "External Access Policy" -Result "Error: $ExternalAccessPolicy is not valid or does not exist"
                    }

                    # Output final ExternalAccessPolicy Success or Fail message
                    Write-Output -InputObject $output
                } # End of setting External Access Policy

                # Mobility Policy
                if ($PSBoundParameters.ContainsKey("MobilityPolicy"))
                {
                    # Verify if $MobilityPolicy is a valid policy to assign
                    if ($tenantMobilityPolicies -icontains "Tag:$MobilityPolicy")
                    {
                        try
                        {
                            # Attempt to assign policy
                            Grant-CsMobilityPolicy -Identity $ID -PolicyName $MobilityPolicy -WarningAction SilentlyContinue -ErrorAction STOP
                            $output = GetActionOutputObject3 -Name $ID -Property "Mobility Policy" -Result "Success: $MobilityPolicy"
                        }
                        catch
                        {
                            $errorMessage = $_                            
                            $output = GetActionOutputObject3 -Name $ID -Property "Mobility Policy" -Result "Error: $errorMessage"
                        }
                    }
                    else
                    {
                        # Output invalid external access policy to error log file
                        $output = GetActionOutputObject3 -Name $ID -Property "Mobility Policy" -Result "Error: $MobilityPolicy is not valid or does not exist"
                    }

                    # Output final MobilityPolicy Success or Fail message
                    Write-Output -InputObject $output
                } # End of setting Mobility Policy
            } # End of setting policies
            else
            {
                $output = GetActionOutputObject3 -Name $ID -Property "User Validation" -Result "Error: Not a valid Skype user account"
                Write-Output -InputObject $output
            }
        } # End of foreach ($ID in $Identity)
    } # End of PROCESS block
} # End of Set-SkypeOnlineUserPolicy

function Test-SkypeOnlineExternalDNS
{
<#
.SYNOPSIS
Tests a domain for the required external DNS records for a Skype for Business Online deployment.

.DESCRIPTION
Skype for Business Online requires the use of several external DNS records for clients and federated
partners to locate services and users. This function will look for the required external DNS records
and display their current values, if they are correctly implemented, and any issues with the records.

.PARAMETER Domain
The domain name to test records. This parameter is required.

.EXAMPLE
Test-SkypeOnlineExternalDNS -Domain contoso.com
Example 1 will test the contoso.com domain for the required external DNS records for Skype for Business Online.
#>

    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, HelpMessage="This is the domain name to test the external DNS Skype Online records.")]
        [string]$Domain
    )

    # VARIABLES
    [string]$federationSRV = "_sipfederationtls._tcp.$Domain"
    [string]$sipSRV = "_sip._tls.$Domain"
    [string]$lyncdiscover = "lyncdiscover.$Domain"
    [string]$sip = "sip.$Domain"

    # Federation SRV Record Check
    $federationSRVResult = Resolve-DnsName -Name "_sipfederationtls._tcp.$Domain" -Type SRV -ErrorAction SilentlyContinue
    $federationOutput = [PSCustomObject][ordered]@{
            Name = $federationSRV
            Type = "SRV"
            Target = $null
            Port = $null
            Correct = "Yes"
            Notes = $null
        }

    if ($null -ne $federationSRVResult)
    {
        $federationOutput.Target = $federationSRVResult.NameTarget
        $federationOutput.Port = $federationSRVResult.Port
        if ($federationOutput.Target -ne "sipfed.online.lync.com")
        {
            $federationOutput.Notes += "Target FQDN is not correct for Skype Online. "
            $federationOutput.Correct = "No"
        }

        if ($federationOutput.Port -ne "5061")
        {
            $federationOutput.Notes += "Port is not set to 5061. "
            $federationOutput.Correct = "No"
        }
    }
    else
    {
        $federationOutput.Notes = "Federation SRV record does not exist. "
        $federationOutput.Correct = "No"
    }

    Write-Output -InputObject $federationOutput
    
    # SIP SRV Record Check
    $sipSRVResult = Resolve-DnsName -Name $sipSRV -Type SRV -ErrorAction SilentlyContinue
    $sipOutput = [PSCustomObject][ordered]@{
            Name = $sipSRV
            Type = "SRV"
            Target = $null
            Port = $null
            Correct = "Yes"
            Notes = $null
        }

    if ($null -ne $sipSRVResult)
    {
        $sipOutput.Target = $sipSRVResult.NameTarget
        $sipOutput.Port = $sipSRVResult.Port
        if ($sipOutput.Target -ne "sipdir.online.lync.com")
        {
            $sipOutput.Notes += "Target FQDN is not correct for Skype Online. "
            $sipOutput.Correct = "No"
        }

        if ($sipOutput.Port -ne "443")
        {
            $sipOutput.Notes += "Port is not set to 443. "
            $sipOutput.Correct = "No"
        }
    }
    else
    {
        $sipOutput.Notes = "SIP SRV record does not exist. "
        $sipOutput.Correct = "No"
    }

    Write-Output -InputObject $sipOutput

    #Lyncdiscover Record Check
    $lyncdiscoverResult = Resolve-DnsName -Name $lyncdiscover -Type CNAME -ErrorAction SilentlyContinue
    $lyncdiscoverOutput = [PSCustomObject][ordered]@{
            Name = $lyncdiscover
            Type = "CNAME"
            Target = $null
            Port = $null
            Correct = "Yes"
            Notes = $null
        }

    if ($null -ne $lyncdiscoverResult)
    {
        $lyncdiscoverOutput.Target = $lyncdiscoverResult.NameHost
        $lyncdiscoverOutput.Port = "----"
        if ($lyncdiscoverOutput.Target -ne "webdir.online.lync.com")
        {
            $lyncdiscoverOutput.Notes += "Target FQDN is not correct for Skype Online. "
            $lyncdiscoverOutput.Correct = "No"
        }
    }
    else
    {
        $lyncdiscoverOutput.Notes = "Lyncdiscover record does not exist. "
        $lyncdiscoverOutput.Correct = "No"
    }

    Write-Output -InputObject $lyncdiscoverOutput

    #SIP Record Check
    $sipResult = Resolve-DnsName -Name $sip -Type CNAME -ErrorAction SilentlyContinue
    $sipOutput = [PSCustomObject][ordered]@{
            Name = $sip
            Type = "CNAME"
            Target = $null
            Port = $null
            Correct = "Yes"
            Notes = $null
        }

    if ($null -ne $sipResult)
    {
        $sipOutput.Target = $sipResult.NameHost
        $sipOutput.Port = "----"
        if ($sipOutput.Target -ne "sipdir.online.lync.com")
        {
            $sipOutput.Notes += "Target FQDN is not correct for Skype Online. "
            $sipOutput.Correct = "No"
        }
    }
    else
    {
        $sipOutput.Notes = "SIP record does not exist. "
        $sipOutput.Correct = "No"
    }

    Write-Output -InputObject $sipOutput
} # End of Test-SkypeOnlineExternalDNS

# *** Non-Exported Helper Functions ***

# 2 Parameter Version
function GetActionOutputObject2
{
    param(
        [Parameter(Mandatory = $true, HelpMessage = "Name of account being modified")]
        [string]$Name,

        [Parameter(Mandatory = $true, HelpMessage = "Result of action being performed")]
        [string]$Result
    )
        
    $outputReturn = [PSCustomObject][ordered]@{
        User = $Name
        Result = $Result
    }

    return $outputReturn
} # End of GetActionOutputObject2

# 3 Parameter Version
function GetActionOutputObject3
{
    param(
        [Parameter(Mandatory = $true, HelpMessage = "Name of account being modified")]
        [string]$Name,

        [Parameter(Mandatory = $true, HelpMessage = "Object/property that is being modified")]
        [string]$Property,

        [Parameter(Mandatory = $true, HelpMessage = "Result of action being performed")]
        [string]$Result
    )
        
    $outputReturn = [PSCustomObject][ordered]@{
        User = $Name
        Property = $Property
        Result = $Result
    }

    return $outputReturn
} # End of GetActionOutputObject3

function NewLicenseObject
{
    param(
        [Parameter(Mandatory = $true, HelpMessage = "SkuId of the license")]
        [string]$SkuId
    )

    $productLicenseObj = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $productLicenseObj.SkuId = $SkuId
    $assignedLicensesObj = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $assignedLicensesObj.AddLicenses = $productLicenseObj
    return $assignedLicensesObj
}

function TestAzureADModule
{
    [CmdletBinding()]
    param()

    Write-Verbose -Message "Verifying if AzureAD module is installed and available"

    if ((Get-Module -ListAvailable).Name -notcontains "AzureAD")
    {
        Write-Warning -Message "Azure Active Directory PowerShell module is not installed. Please install and try again."
        return $false
    }
}

function TestAzureADConnection
{
    [CmdletBinding()]
    param()

    try
    {
        Get-AzureADCurrentSessionInfo -ErrorAction STOP | Out-Null
    }
    catch
    {
        Write-Warning -Message "A connection to AzureAD must be present before continuing"
        return $false
    }
}

function TestSkypeOnlineModule
{
    [CmdletBinding()]
    param()
    
    if ((Get-Module -ListAvailable).Name -notcontains "SkypeOnlineConnector")
    {        
        return $false
    }
    else
    {
        try
        {
            Import-Module -Name SkypeOnlineConnector
            return $true
        }
        catch
        {
            Write-Warning $_
            return $false
        }
    }
}

function TestSkypeOnlineConnection
{
    [CmdletBinding()]
    param()

    if ((Get-PsSession).ComputerName -notlike "*.online.lync.com")
    {
        return $false
    }
    else
    {
        return $true
    }
}

Export-ModuleMember -Function Add-SkypeOnlineUserLicense, Connect-SkypeOnline, Disconnect-SkypeOnline,`
                              Get-SkypeOnlineConferenceDialInNumbers, Get-SkypeOnlineUserLicense, Get-SkypeOnlineTenantLicenses, Set-SkypeOnlineUserPolicy,`
                              Remove-SkypeOnlineNormalizationRule, Test-SkypeOnlineExternalDNS