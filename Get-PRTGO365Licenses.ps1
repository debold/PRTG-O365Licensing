<#
.SYNOPSIS
Retrieves current license information from Office 365 tenant in PRTG compatible format

.DESCRIPTION
The Get-Office365License.ps1 uses Office 365 Power Shell extensions to retrieve available licensing information for your Office 365 tenant. The XML output can be used as PRTG custom sensor.

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
Get-Office365License.ps1 -O365User "User@Tenant.onmicrosoft.com" -O365Pass "TopSecret" -ExcludeSku "tenant:POWER_BI_STANDARD"

.NOTES
You need to install the Microsoft Online Services-Sign in assistant (http://go.microsoft.com/fwlink/?LinkID=286152) and Azure Active Directory-Module for Windows PowerShell
(http://go.microsoft.com/fwlink/p/?linkid=236297) on the probe device in order to get this script to work. More details see link below.

Author:  Marc Debold
Version: 1.1
Version History:
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
http://www.team-debold.de/2016/10/16/prtg-office-365-lizenzen-im-blick/ ‎

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
    $cmd = "$($env:SystemRoot)\sysnative\windowspowershell\v1.0\powershell.exe"
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
    Invoke-Expression "$env:windir\sysnative\WindowsPowerShell\v1.0\powershell.exe -file '$($MyInvocation.MyCommand.Definition)' $ScriptParameter"
    Exit
}

#Initialize varaibles
$IsError = $false
$ErrorText = ""
$Result = $null

<# Function to raise error in PRTG style and stop script #>
function Raise-MyError {
    [CmdletBinding()] param(
        [Parameter(Position=0)][string] $ErrorText
    )

    Write-Host "<prtg>
    <error>1</error>
    <text>$ErrorText</text>
</prtg>"
    Exit
}

<# Function to form XML and print it to host #>
function Out-Prtg {
    [CmdletBinding()] param(
        [Parameter(Position=0, Mandatory=$true)] $MonitoringData
    )

    <# Create output for PRTG #>
    $XmlDocument = New-Object System.XML.XMLDocument
    $XmlRoot = $XmlDocument.CreateElement("prtg")
    $XmlDocument.appendChild($XmlRoot) | Out-Null
    foreach ($Item in $MonitoringData) {
        # Available Count
        $XmlResult = $XmlRoot.appendChild($XmlDocument.CreateElement("result"))
        $XmlKey = $XmlDocument.CreateElement("channel")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode("$($Item.Id) - Available Licenses")
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("value")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode($Item.Available)
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("unit")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode("percent")
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("float")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode(1)
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("limitminwarning")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode(20)
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("limitminerror")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode(10)
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("limitmode")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode(1)
        $XmlKey.AppendChild($XmlValue) | Out-Null
        # Warning Count
        $XmlResult = $XmlRoot.appendChild($XmlDocument.CreateElement("result"))
        $XmlKey = $XmlDocument.CreateElement("channel")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode("$($Item.Id) - Warning Licenses")
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("value")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode($Item.Warning)
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("unit")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode("percent")
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("float")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode(1)
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("limitmaxwarning")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode(0)
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("limitmode")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode(1)
        $XmlKey.AppendChild($XmlValue) | Out-Null
        # Active Count
        $XmlResult = $XmlRoot.appendChild($XmlDocument.CreateElement("result"))
        $XmlKey = $XmlDocument.CreateElement("channel")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode("$($Item.Id) - Active Licenses")
        $XmlKey.AppendChild($XmlValue) | Out-Null
        $XmlKey = $XmlDocument.CreateElement("value")
        $XmlResult.AppendChild($XmlKey) | Out-Null
        $XmlValue = $XmlDocument.CreateTextNode($Item.Active)
        $XmlKey.AppendChild($XmlValue) | Out-Null
    }
    <# Format XML output #>
    $StringWriter = New-Object System.IO.StringWriter 
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
    $XmlWriter.Formatting = “indented” 
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
    Raise-MyError "Could not load PowerShell Module MSOnline"
}

# Connect to Office 365 using provided credential
try {
    Connect-MsolService -Credential $O365Cred -ErrorAction Stop
} catch {
    Raise-MyError "Error connecting to your tenant. Please check credentials"
}

# Get all licenses
$RawSkus = Get-MsolAccountSku

# Write licenses to host, if switch is set
if ($ShowMySkus) {
    $RawSkus | ft -AutoSize
} else {
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
            $Result = @()
            foreach ($Sku in $Skus) {
                $Result += @{
                    Id = [string]$Sku.AccountSkuId; 
                    Active = [int]$Sku.ActiveUnits;
                    Warning = [math]::Round($Sku.WarnungUnits/$Sku.ActiveUnits*100, 2);
                    Available = [math]::Round(($Sku.ActiveUnits - $Sku.ConsumedUnits)/$Sku.ActiveUnits*100, 2)
                }
            }
        } else {
            Raise-MyError "No Skus found"
        }
        if ($Result -ne $null) {
            Out-Prtg -MonitoringData $Result
        }
    }
}