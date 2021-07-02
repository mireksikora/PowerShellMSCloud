<# 
Author:         Mirek Sikora
Date:           2021/07/02 (YYYY/MM/DD)
Description:    Generate Cloud report and save it to CSV file
#>

#install-module PowerShellGet -Force -Scope AllUsers
#import-module PowerShellMS
#install-Module AzureAD -Force -Scope AllUsers
#update-Module AzureAD


#Define parameters
#-------------------
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false)]
    [switch] $SecurityOnly,                 # Security group list
    [switch] $DistributionListOnly,         # Distribition List
    [switch] $MailEnabledSecurityOnly,      # Mail Enabled Security group list
    [switch] $All,                          # All groups list with included group types
    [switch] $MFA,
    [string] $TenantName,
    [string] $UserName,
    [string] $Password
)

# Connect to cloud and define variables 
#---------------------------------------
import-module PowerShellMS
If($MFA.IsPresent) { $MFAEnabled = $true } else { $MFAEnabled = $false}
ConnectToMSOLService $MFAEnabled $UserName $Password

$CSVPath,$ExportCSV = SetOutputPathFilename $TenantName

# Set report result based on input parameters
#-----------------------------------------------

If(($SecurityOnly -eq $false) -and ($DistributionListOnly -eq $false) -and ($MailEnabledSecurityOnly -eq $false) -and ($All -eq $false)){
    $All = $true        #default
}

If($All.IsPresent) {    # List All goups and include group types

    $GroupProperties = "DisplayName","EmailAddress","GroupType","ManagedBy","Description"
    $ResultArray = Get-MsolGroup | select-object -property $GroupProperties
    $count=$ResultArray.count
}

If($SecurityOnly.IsPresent) {    # Security Only goup list

    $GroupProperties = "DisplayName","EmailAddress","GroupType","ManagedBy","Description"
    $ResultArray = Get-MsolGroup | select-object -property $GroupProperties | where-object { $_.GroupType -eq "Security" }
    $count=$ResultArray.count
}

If($DistributionListOnly.IsPresent) {    # Distribution List Only

    $GroupProperties = "DisplayName","EmailAddress","GroupType","ManagedBy","Description"
    $ResultArray = Get-MsolGroup | select-object -property $GroupProperties | where-object { $_.GroupType -eq "DistributionList" }
    $count=$ResultArray.count
}

If($MailEnabledSecurityOnly.IsPresent) {    # Mail Enabled Security Only group list

    $GroupProperties = "DisplayName","EmailAddress","GroupType","ManagedBy","Description"
    $ResultArray = Get-MsolGroup | select-object -property $GroupProperties | where-object { $_.GroupType -eq "MailEnabledSecurity" }
    $count=$ResultArray.count  
} 


# Export to CSV and show result. Disconnect from cloud
#------------------------------------------------------
ExportToCSV $ResultArray $CSVPath $ExportCSV
ShowScriptResult $ExportCSV
Write-Host Group count that matched criteria $count -ForegroundColor Yellow
DisconnectFromMSOLService
#end of script

