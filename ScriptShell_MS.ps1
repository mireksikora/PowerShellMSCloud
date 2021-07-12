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
    [switch] $param_1,      #optional parameter for report 1
    [switch] $param_2,      #optional parameter for report 2
    [switch] $param_3,      #optional parameter for report 3
    [switch] $param_4,      #optional parameter for report 4
    [switch] $MFA,
    [string] $TenantName,
    [string] $UserName,
    [string] $Password
)

$startTime = Get-Date -format G
Write-Host "The script started on $startTime"

# Connect to cloud and define variables 
#---------------------------------------
import-module PowerShellMS
If($MFA.IsPresent) { $MFAEnabled = $true } else { $MFAEnabled = $false}
ConnectToAzureAD $MFAEnabled $UserName $Password 
# or
# ConnectToMSOLService $MFAEnabled $UserName $Password
# or
# ConnectToExchangeOnline $MFAEnabled $UserName $Password

$CSVPath,$ExportCSV = SetOutputPathFilename $TenantName

# Set report result based on input parameters
#-----------------------------------------------

If(($param_1 -eq $false) -and ($param_2 -eq $false) -and ($param_3 -eq $false) -and ($param_4 -eq $false)){
    $param_1 = $true        #default
}

If($param_1.IsPresent) {    #param_1
    <#
    #>
}

If($param_2.IsPresent) {    #param_2
    <#
    #>
}

If($param_3.IsPresent) {    #param_3
    <#
    #>
}

If($param_4.IsPresent) {    #param_4
    <#
    #>
} 


# Export to CSV and show result. Disconnect from cloud
#------------------------------------------------------
If( -not ([string]::IsNullOrEmpty($ResultArray))) {
    ExportToCSV $ResultArray $CSVPath $ExportCSV
    ShowScriptResult $ExportCSV
}

DisconnectFromAzureAD
# or
# DisconnectFromMSOLService
#or
# DisconnectFromExchangeOnline

$endTime = Get-Date -format G
Write-Host "Script completed at $endTime.`n"
Write-host "Time for the full run was: $( New-TimeSpan $startTime $endTime)."

#end of script

