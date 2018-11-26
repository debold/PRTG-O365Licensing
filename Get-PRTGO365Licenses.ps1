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
                $Result += @{
                    Channel = "$($Sku.AccountSkuId.SubString($Sku.AccountSkuId.IndexOf(':')+1)) - Free Licenses";
                    Value = [int]($Sku.ActiveUnits - $Sku.ConsumedUnits);
                    LimitMinWarning = 5;
                    LimitMinError = 1;
                    LimitMode = 1
                }
<#                $Result += @{
                    Channel = "$($Sku.AccountSkuId.SubString($Sku.AccountSkuId.IndexOf(':')+1)) - Warning Licenses";
                    Value = [int]$Sku.WarnungUnits;
                    LimitMaxWarning = 0;
                    LimitMode = 1
                }#>
                $Result += @{
                    Channel = "$($Sku.AccountSkuId.SubString($Sku.AccountSkuId.IndexOf(':')+1)) - Total Licenses";
                    Value = [int]$Sku.ActiveUnits
                }
<#                $Result += @{
                    Channel = "$($Sku.AccountSkuId.SubString($Sku.AccountSkuId.IndexOf(':')+1)) - Consumed Licenses";
                    Value = [int]$Sku.ConsumedUnits
                }#>
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