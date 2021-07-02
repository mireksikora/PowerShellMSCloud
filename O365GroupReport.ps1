<# 
Author:         Mirek Sikora
Date:           2021/06/20 (YYYY/MM/DD)
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
    [switch] $SecurityOnly,       #DisabledOnly user list
    [switch] $DistributionListOnly,        #EnabledOnly user list
    [switch] $MailEnabledSecurityOnly,       #EnforcedOnly user list
    [switch] $All,                #All user list with MFA status
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

If($All.IsPresent) {    #param All

    $GroupProperties = "DisplayName","EmailAddress","GroupType","ManagedBy","Description"
    $ResultArray = Get-MsolGroup | select-object -property $GroupProperties
    $count=$ResultArray.count
}

If($SecurityOnly.IsPresent) {    #param SecurityOnly

    $GroupProperties = "DisplayName","EmailAddress","GroupType","ManagedBy","Description"
    $ResultArray = Get-MsolGroup | select-object -property $GroupProperties | where-object { $_.GroupType -eq "Security" }
    $count=$ResultArray.count
}

If($DistributionListOnly.IsPresent) {    #param DistributionListOnly

    $GroupProperties = "DisplayName","EmailAddress","GroupType","ManagedBy","Description"
    $ResultArray = Get-MsolGroup | select-object -property $GroupProperties | where-object { $_.GroupType -eq "DistributionList" }
    $count=$ResultArray.count
}

If($MailEnabledSecurityOnly.IsPresent) {    #param ALL

    $GroupProperties = "DisplayName","EmailAddress","GroupType","ManagedBy","Description"
    $ResultArray = Get-MsolGroup | select-object -property $GroupProperties | where-object { $_.GroupType -eq "MailEnabledSecurity" }
    $count=$ResultArray.count  
} 


# Export to CSV and show result. Disconnect from cloud
#------------------------------------------------------
ExportToCSV $ResultArray $CSVPath $ExportCSV
ShowScriptResult $ExportCSV
Write-Host User count that matched criteria $count -ForegroundColor Yellow
DisconnectFromMSOLService
#end of script

