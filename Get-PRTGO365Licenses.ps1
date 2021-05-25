<#
.SYNOPSIS
Retrieves current license information from Office 365 tenant in PRTG compatible format

.DESCRIPTION
The Get-Office365License.ps1 uses Office 365 Power Shell extensions to retrieve available licensing information for your Office 365 tenant. 
The XML output can be used as PRTG custom sensor.

.PARAMETER O365User
Represents the username for connecting to your Office 365 tenant. See NOTES section for more details.

.PARAMETER O365Pass
Represents the corresponding password for the user connecting to your tenant.

.PARAMETER IncludeSku
If you have more than one license contract with Microsoft, you propably want to monitor only specific of them, or monitor each in seperate.
Use this parameter to limit the license calculation to the provided SKUs.
Can be a single string, comma separated or array of stings. 
To use all SKUs, skip this parameter
If IncludeSku parameter is used, ExcludeSku IS NOT PROCESSED!

.PARAMETER ExcludeSku
If you have more than one license contract with Microsoft, you propably want some of them NOT to be monitored.
Use this parameter to skip the provided SKUs in calculation.
Can be a single string, comma separated or array of stings. 
To use all SKUs, skip this parameter
If IncludeSku parameter is used, ExcludeSku IS NOT PROCESSED!

.PARAMETER ShowMySkus
Use this parameter to get an output (NOT FOR PRTG) of your avaialbe SKUs and their counts. Use it to find out, which SKUs to include/exclude using the parameters IncludeSku and ExcludeSku.
Using this parameter will ignore the parameters IncludeSku and ExcludeSku and is for manual usage only!

.EXAMPLE
Retrieves complete license count of your tenant.
Get-Office365License.ps1 -O365User "User@Tenant.onmicrosoft.com" -O365Pass "TopSecret"

.EXAMPLE
Shows a list of all SKUs in your tenant. Not for use with PRTG. Use this to determine values for parameters IncludeSku and ExcludeSku.
Get-Office365License.ps1 -O365User "User@Tenant.onmicrosoft.com" -O365Pass "TopSecret" -ShowMySkus

.EXAMPLE
Get license count for your EnterprisePack SKU only.
Get-Office365License.ps1 -O365User "User@Tenant.onmicrosoft.com" -O365Pass "TopSecret" -IncludeSku "tenant:ENTERPRISEPACK"

.EXAMPLE
Get license count for all of your SKUs, except for POWERBI.
Get-Office365License.ps1 -O365User "User@Tenant.onmicrosoft.com" -O365Pass "TopSecret" -ExcludeSku "tenant:POWER_BI_STANDARD,tenant:FLOW_FREE"

.NOTES
You need to install the Microsoft Online Services-Sign in assistant (http://go.microsoft.com/fwlink/?LinkID=286152) and Azure Active Directory-Module for Windows PowerShell
(http://go.microsoft.com/fwlink/p/?linkid=236297) on the probe device in order to get this script to work. More details see link below.

Author:  Marc Debold
Version: 1.4
Version History:
    1.4  26.11.2018  Fixed ExcludeSku & IncludeSku issue
                     Removed some license counters for less channels
                     Removed prefix of license channel names
                     Added detection of sync errors
                     Updated channel names
    1.3  31.10.2016  Fixed time offset issue
                     Changed XML output
    1.2  31.10.2016  Code cleanup
                     Changed output to absolute numbers
                     Added consumed units
                     Added monitoring of last DirSync (and password sync) if available
    1.1  29.10.2016  Added workaround for PRTG running PS scripts in x86 environment only
                     Launching separate PowerShell process in x64 environment using sysnative folder
    1.0  16.10.2016  Initial release

Thanks to Ollie for the great idea.

Relevant information in this context:
AAD Basics
http://aka.ms/aadposh

Download: Microsoft Online Services-Anmelde-Assistenten für IT-Experten RTW
http://go.microsoft.com/fwlink/?LinkID=286152

Download: Azure Active Directory-Modul für Windows PowerShell (64-Bit-Version)
http://go.microsoft.com/fwlink/p/?linkid=236297

.LINK
http://www.team-debold.de/2016/11/05/prtg-office-365-lizenzen-im-blick/
#>

[CmdletBinding()] param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $O365User,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string] $O365Pass,
    [string[]] $IncludeSku = $null,
    [string[]] $ExcludeSku = $null,
    [switch] $ShowMySkus
)

#Check for 32 bit environment
if ($env:PROCESSOR_ARCHITECTURE -eq "x86") {
    # Launch script in 64 bit environment
    $ScriptParameter = "-O365User '$O365User' -O365Pass '$O365Pass' "
    if ($IncludeSku -ne $null) {
        $ScriptParameter += "-IncludeSku '$($IncludeSku -join "','")' "
    }
    if ($ExcludeSku -ne $null) {
        $ScriptParameter += "-ExcludeSku '$($ExcludeSku -join "','")' "
    }
    if ($ShowMySkus) {
        $ScriptParameter += "-ShowMySkus "
    }
    # Use Sysnative virtual directory on 64-bit machines
    Invoke-Expression "$env:windir\sysnative\WindowsPowerShell\v1.0\powershell.exe -file '$($MyInvocation.MyCommand.Definition)' $ScriptParameter"
    Exit
}

#Initialize varaibles
$Result = @()
$SyncStats = @{
    TimeSinceLastDirSync = $null;
    TimeSinceLastPassSync = $null
}

if ($ExcludeSku -ne $null) {
    $ExcludeSku = $ExcludeSku -split ","
}

if ($IncludeSku -ne $null) {
    $IncludeSku = $IncludeSku -split ","
}

$CSV = "C:\Scripts\Office_365_User_Licensing.csv"

#Friendly License Name as hash table
#https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
$SkuHashTable = @{
"MCOMEETADV"="Microsoft 365 Audio Conferencing"
"AAD_BASIC"="AZURE ACTIVE DIRECTORY BASIC"
"AAD_PREMIUM"="AZURE ACTIVE DIRECTORY PREMIUM P1"
"AAD_PREMIUM_P2"="AZURE ACTIVE DIRECTORY PREMIUM P2"
"RIGHTSMANAGEMENT"="AZURE INFORMATION PROTECTION PLAN 1"
"MCOCAP"="COMMON AREA PHONE"
"MCOPSTNC"="COMMUNICATIONS CREDITS"
"DYN365_ENTERPRISE_PLAN1"="DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION"
"DYN365_ENTERPRISE_CUSTOMER_SERVICE"="DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION"
"DYN365_FINANCIALS_BUSINESS_SKU"="DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION"
"DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE"="DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION"
"DYN365_ENTERPRISE_SALES"="DYNAMICS 365 FOR SALES ENTERPRISE EDITION"
"DYN365_ENTERPRISE_TEAM_MEMBERS"="DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION"
"DYN365_TEAM_MEMBERS"="DYNAMICS 365 TEAM MEMBERS"
"Dynamics_365_for_Operations"="DYNAMICS 365 UNF OPS PLAN ENT EDITION"
"EMS"="ENTERPRISE MOBILITY + SECURITY E3"
"EMSPREMIUM"="ENTERPRISE MOBILITY + SECURITY E5"
"EXCHANGESTANDARD"="EXCHANGE ONLINE (PLAN 1)"
"EXCHANGEENTERPRISE"="EXCHANGE ONLINE (PLAN 2)"
"EXCHANGEARCHIVE_ADDON"="EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE"
"EXCHANGEARCHIVE"="EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER"
"EXCHANGEESSENTIALS"="EXCHANGE ONLINE ESSENTIALS"
"EXCHANGE_S_ESSENTIALS"="EXCHANGE ONLINE ESSENTIALS"
"EXCHANGEDESKLESS"="EXCHANGE ONLINE KIOSK"
"EXCHANGETELCO"="EXCHANGE ONLINE POP"
"INTUNE_A"="INTUNE"
"M365EDU_A1"="Microsoft 365 A1"
"M365EDU_A3_FACULTY"="MICROSOFT 365 A3 FOR FACULTY"
"M365EDU_A3_STUDENT"="MICROSOFT 365 A3 FOR STUDENTS"
"M365EDU_A5_FACULTY"="MICROSOFT 365 A5 FOR FACULTY"
"M365EDU_A5_STUDENT"="MICROSOFT 365 A5 FOR STUDENTS"
"O365_BUSINESS"="MICROSOFT 365 APPS FOR BUSINESS"
"SMB_BUSINESS"="MICROSOFT 365 APPS FOR BUSINESS"
"OFFICESUBSCRIPTION"="MICROSOFT 365 APPS FOR ENTERPRISE"
"MCOMEETADV_GOC"="MICROSOFT 365 AUDIO CONFERENCING FOR GCC"
"O365_BUSINESS_ESSENTIALS"="MICROSOFT 365 BUSINESS BASIC"
"SMB_BUSINESS_ESSENTIALS"="MICROSOFT 365 BUSINESS BASIC"
"O365_BUSINESS_PREMIUM"="MICROSOFT 365 BUSINESS STANDARD"
"SMB_BUSINESS_PREMIUM"="MICROSOFT 365 BUSINESS STANDARD"
"SPB"="MICROSOFT 365 BUSINESS PREMIUM"
"MCOPSTN_5"="MICROSOFT 365 DOMESTIC CALLING PLAN (120 Minutes)"
"SPE_E3"="MICROSOFT 365 E3"
"SPE_E5"="Microsoft 365 E5"
"SPE_E3_USGOV_DOD"="Microsoft 365 E3_USGOV_DOD"
"SPE_E3_USGOV_GCCHIGH"="Microsoft 365 E3_USGOV_GCCHIGH"
"INFORMATION_PROTECTION_COMPLIANCE"="Microsoft 365 E5 Compliance"
"IDENTITY_THREAT_PROTECTION"="Microsoft 365 E5 Security"
"IDENTITY_THREAT_PROTECTION_FOR_EMS_E5"="Microsoft 365 E5 Security for EMS E5"
"M365_F1"="Microsoft 365 F1"
"SPE_F1"="Microsoft 365 F3"
"FLOW_FREE"="MICROSOFT FLOW FREE"
"M365_G3_GOV"="MICROSOFT 365 GCC G3"
"MCOEV"="MICROSOFT 365 PHONE SYSTEM"
"MCOEV_DOD"="MICROSOFT 365 PHONE SYSTEM FOR DOD"
"MCOEV_FACULTY"="MICROSOFT 365 PHONE SYSTEM FOR FACULTY"
"MCOEV_GOV"="MICROSOFT 365 PHONE SYSTEM FOR GCC"
"MCOEV_GCCHIGH"="MICROSOFT 365 PHONE SYSTEM FOR GCCHIGH"
"MCOEVSMB_1"="MICROSOFT 365 PHONE SYSTEM FOR SMALL AND MEDIUM BUSINESS"
"MCOEV_STUDENT"="MICROSOFT 365 PHONE SYSTEM FOR STUDENTS"
"MCOEV_TELSTRA"="MICROSOFT 365 PHONE SYSTEM FOR TELSTRA"
"MCOEV_USGOV_DOD"="MICROSOFT 365 PHONE SYSTEM_USGOV_DOD"
"MCOEV_USGOV_GCCHIGH"="MICROSOFT 365 PHONE SYSTEM_USGOV_GCCHIGH"
"PHONESYSTEM_VIRTUALUSER"="MICROSOFT 365 PHONE SYSTEM - VIRTUAL USER"
"WIN_DEF_ATP"="Microsoft Defender Advanced Threat Protection"
"CRMPLAN2"="MICROSOFT DYNAMICS CRM ONLINE BASIC"
"CRMSTANDARD"="MICROSOFT DYNAMICS CRM ONLINE"
"IT_ACADEMY_AD"="MS IMAGINE ACADEMY"
"INTUNE_A_D_GOV"="MICROSOFT INTUNE DEVICE for GOVERNMENT"
"POWERAPPS_VIRAL"="MICROSOFT POWER APPS PLAN 2 TRIAL"
"TEAMS_FREE"="MICROSOFT TEAM (FREE)"
"TEAMS_EXPLORATORY"="MICROSOFT TEAMS EXPLORATORY"
"ENTERPRISEPREMIUM_FACULTY"="Office 365 A5 for faculty"
"ENTERPRISEPREMIUM_STUDENT"="Office 365 A5 for students"
"EQUIVIO_ANALYTICS"="Office 365 Advanced Compliance"
"ATP_ENTERPRISE"="Office 365 Advanced Threat Protection (Plan 1)"
"STANDARDPACK"="OFFICE 365 E1"
"STANDARDWOFFPACK"="OFFICE 365 E2"
"ENTERPRISEPACK"="OFFICE 365 E3"
"DEVELOPERPACK"="OFFICE 365 E3 DEVELOPER"
"ENTERPRISEPACK_USGOV_DOD"="Office 365 E3_USGOV_DOD"
"ENTERPRISEPACK_USGOV_GCCHIGH"="Office 365 E3_USGOV_GCCHIGH"
"ENTERPRISEWITHSCAL"="OFFICE 365 E4"
"ENTERPRISEPREMIUM"="OFFICE 365 E5"
"ENTERPRISEPREMIUM_NOPSTNCONF"="OFFICE 365 E5 WITHOUT AUDIO CONFERENCING"
"DESKLESSPACK"="OFFICE 365 F3"
"ENTERPRISEPACK_GOV"="OFFICE 365 GCC G3"
"MIDSIZEPACK"="OFFICE 365 MIDSIZE BUSINESS"
"LITEPACK"="OFFICE 365 SMALL BUSINESS"
"LITEPACK_P2"="OFFICE 365 SMALL BUSINESS PREMIUM"
"WACONEDRIVESTANDARD"="ONEDRIVE FOR BUSINESS (PLAN 1)"
"WACONEDRIVEENTERPRISE"="ONEDRIVE FOR BUSINESS (PLAN 2)"
"POWER_BI_STANDARD"="POWER BI (FREE)"
"POWER_BI_ADDON"="POWER BI FOR OFFICE 365 ADD-ON"
"POWER_BI_PRO"="POWER BI PRO"
"PROJECTCLIENT"="PROJECT FOR OFFICE 365"
"PROJECTESSENTIALS"="PROJECT ONLINE ESSENTIALS"
"PROJECTPREMIUM"="PROJECT ONLINE PREMIUM"
"PROJECTONLINE_PLAN_1"="PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT"
"PROJECTPROFESSIONAL"="PROJECT ONLINE PROFESSIONAL"
"PROJECTONLINE_PLAN_2"="PROJECT ONLINE WITH PROJECT FOR OFFICE 365"
"SHAREPOINTSTANDARD"="SHAREPOINT ONLINE (PLAN 1)"
"SHAREPOINTENTERPRISE"="SHAREPOINT ONLINE (PLAN 2)"
"MCOIMP"="SKYPE FOR BUSINESS ONLINE (PLAN 1)"
"MCOSTANDARD"="SKYPE FOR BUSINESS ONLINE (PLAN 2)"
"MCOPSTN2"="SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING"
"MCOPSTN1"="SKYPE FOR BUSINESS PSTN DOMESTIC CALLING"
"MCOPSTN5"="SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes)"
"TOPIC_EXPERIENCES"="TOPIC EXPERIENCES"
"MCOPSTNEAU2"="TELSTRA CALLING FOR O365"
"VISIOONLINE_PLAN1"="VISIO ONLINE PLAN 1"
"VISIOCLIENT"="VISIO Online Plan 2"
"VISIOCLIENT_GOV"="VISIO PLAN 2 FOR GOV"
"WIN10_PRO_ENT_SUB"="WINDOWS 10 ENTERPRISE E3"
"WIN10_VDA_E3"="WINDOWS 10 ENTERPRISE E3"
"WIN10_VDA_E5"="Windows 10 Enterprise E5"
"WINDOWS_STORE"="WINDOWS STORE FOR BUSINESS"
}

<# Function to raise error in PRTG style and stop script #>
function New-PrtgError {
    [CmdletBinding()] param(
        [Parameter(Position=0)][string] $ErrorText
    )

    Write-Host "<PRTG>
    <Error>1</Error>
    <Text>$ErrorText</Text>
</PRTG>"
    Exit
}

function Out-Prtg {
    [CmdletBinding()] param(
        [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)][array] $MonitoringData
    )
    # Create output for PRTG
    $XmlDocument = New-Object System.XML.XMLDocument
    $XmlRoot = $XmlDocument.CreateElement("PRTG")
    $XmlDocument.appendChild($XmlRoot) | Out-Null
    # Cycle through outer array
    foreach ($Result in $MonitoringData) {
        # Create result-node
        $XmlResult = $XmlRoot.appendChild(
            $XmlDocument.CreateElement("Result")
        )
        # Cycle though inner hashtable
        $Result.GetEnumerator() | ForEach-Object {
            # Use key of hashtable as XML element
            $XmlKey = $XmlDocument.CreateElement($_.key)
            $XmlKey.AppendChild(
                # Use value of hashtable as XML textnode
                $XmlDocument.CreateTextNode($_.value)    
            ) | Out-Null
            $XmlResult.AppendChild($XmlKey) | Out-Null
        }
    }
    # Format XML output
    $StringWriter = New-Object System.IO.StringWriter 
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
    $XmlWriter.Formatting = "indented"
    $XmlWriter.Indentation = 1
    $XmlWriter.IndentChar = "`t" 
    $XmlDocument.WriteContentTo($XmlWriter) 
    $XmlWriter.Flush() 
    $StringWriter.Flush() 
    Return $StringWriter.ToString() 
}

# Create credential object
$O365Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $O365User, ($O365Pass | ConvertTo-SecureString -AsPlainText -Force)

# Load Office 365 module
try {
    Import-Module MSOnline -ErrorAction Stop
} catch {
    New-PrtgError -ErrorText "Could not load PowerShell Module MSOnline"
}

# Connect to Office 365 using provided credential
try {
    Connect-MsolService -Credential $O365Cred -ErrorAction Stop
} catch {
    New-PrtgError -ErrorText "Error connecting to your tenant. Please check credentials"
}

# Get all licenses
$RawSkus = Get-MsolAccountSku

# Write licenses to host, if switch is set
if ($ShowMySkus) {
    $RawSkus | ft -AutoSize
} else {
    # Check for DirSync errors
    $Fails = @()
    $isSyncError = 0
    if (Get-MsolHasObjectsWithDirSyncProvisioningErrors) {
        $Fails = Get-MsolDirSyncProvisioningError
        $isSyncError = 1
    }
    $Result += @{
        Channel = "DirSync: Error count";
        Value = $Fails.Count;
        Unit = "Count";
        LimitMaxWarning = 0.5;
        LimitMode = 1;
        LimitWarningMsg = "DirSync errors encountered for: $($Fails.DisplayName -join ', ')"
    }

    # Check for DirSync statistics
    try {
        $CompanyInfo = Get-MsolCompanyInformation -ErrorAction Stop
    } catch {
        New-PrtgError -ErrorText "Unable to retrieve company information"
    }
    if ($CompanyInfo.DirectorySynchronizationEnabled) {
        $Result += @{
            Channel = "DirSync: Time since last run";
            Value = [math]::Round(((Get-Date).ToUniversalTime() - $CompanyInfo.LastDirSyncTime).TotalHours, 2);
            Unit = "TimeHours";
            Float = 1;
            DecimalMode = 2;
            LimitMaxWarning = 12;
            LimitMode = 1
        }
    }
    if ($CompanyInfo.PasswordSynchronizationEnabled) {
        $Result += @{
            Channel = "DirSync: Time since last password sync";
            Value = [math]::Round(((Get-Date).ToUniversalTime() - $CompanyInfo.LastPasswordSyncTime).TotalHours, 2);
            Unit = "TimeHours";
            Float = 1;
            DecimalMode = 2;
            LimitMaxWarning = 12;
            LimitMode = 1
        }
    }
    # Check license count
    if ($RawSkus.Count -gt 0) {
        # Process includes or excludes of skus
        if ($IncludeSku -ne $null) {
            $Skus = Get-MsolAccountSku | ? {$_.AccountSkuId -in $IncludeSku}
        } else {
            $Skus = Get-MsolAccountSku | ? {$_.AccountSkuId -notin $ExcludeSku}
        }
        # See, if there still are skus to process
        if (($Skus | measure).Count -gt 0) {
            # Fetch stats into array
            foreach ($Sku in $Skus) {

                #Finding $Sku.AccountSkuId in the Hash Table
			    $LicenseItem = $Sku.AccountSkuId -split ":" | Select-Object -Last 1
			    $TextLic = $SkuHashTable.Item("$LicenseItem")

                # fallback to not user friendly name
			    If (!($TextLic))
			    {
				    $TextLic = $LicenseItem

			    }

                $Result += @{
                    Channel = "$($TextLic) - Free Licenses";
                    Value = [int]($Sku.ActiveUnits - $Sku.ConsumedUnits);
                    LimitMinWarning = 5;
                    LimitMinError = 1;
                    LimitMode = 1
                }
                $Result += @{
                    Channel = "$($TextLic) - Total Licenses";
                    Value = [int]$Sku.ActiveUnits
                }

            }
        } else {
            New-PrtgError -ErrorText "No Skus found"
        }
    }

    if ($Result -ne $null) {
        Out-Prtg -MonitoringData $Result
        Start-Sleep 0 # For vscode debugging only
    }

}
