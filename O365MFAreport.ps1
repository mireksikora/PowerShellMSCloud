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
    [switch] $DisabledOnly,       #DisabledOnly user list
    [switch] $EnabledOnly,        #EnabledOnly user list
    [switch] $EnforcedOnly,       #EnforcedOnly user list
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

If(($DisabledOnly -eq $false) -and ($EnabledOnly -eq $false) -and ($EnforcedOnly -eq $false) -and ($All -eq $false)){
    $EnabledOnly = $true        #default
}

If($DisabledOnly.IsPresent) {    #param DisabledOnly
    $ResultArray = @()
    Get-MsolUser -All: $true | ForEach-Object {
        If($_.StrongAuthenticationRequirements.State -eq $null) {
        $UserObject = New-Object PSCustomObject
        $UserObject | Add-Member -MemberType NoteProperty -name "Display Name" -value $_.DisplayName
        $UserObject | Add-Member -MemberType NoteProperty -name "E-mail" -value $_.UserPrincipalName
        $UserObject | Add-Member -MemberType NoteProperty -name "MFA Status" -value "Disabled"
        $UserObject | Add-Member -MemberType NoteProperty -name "Country" -value $_.Country
        $count++
        $ResultArray += $UserObject
        }
    }
}

If($EnabledOnly.IsPresent) {    #param EnabledOnly

    $ResultArray = @()
    Get-MsolUser -All: $true | ForEach-Object {
        If($_.StrongAuthenticationRequirements.State -eq "Enabled") {
        $UserObject = New-Object PSCustomObject
        $UserObject | Add-Member -MemberType NoteProperty -name "Display Name" -value $_.DisplayName
        $UserObject | Add-Member -MemberType NoteProperty -name "E-mail" -value $_.UserPrincipalName
        $UserObject | Add-Member -MemberType NoteProperty -name "MFA Status" -value $_.StrongAuthenticationRequirements.State
        $UserObject | Add-Member -MemberType NoteProperty -name "Country" -value $_.Country
        $count++
        $ResultArray += $UserObject
        }
    }
}

If($EnforcedOnly.IsPresent) {    #param EnforcedOnly
    $ResultArray = @()
    Get-MsolUser -All: $true | ForEach-Object {
        If($_.StrongAuthenticationRequirements.State -eq "Enforced") {
        $UserObject = New-Object PSCustomObject
        $UserObject | Add-Member -MemberType NoteProperty -name "Display Name" -value $_.DisplayName
        $UserObject | Add-Member -MemberType NoteProperty -name "E-mail" -value $_.UserPrincipalName
        $UserObject | Add-Member -MemberType NoteProperty -name "MFA Status" -value $_.StrongAuthenticationRequirements.State
        $UserObject | Add-Member -MemberType NoteProperty -name "Country" -value $_.Country
        $count++
        $ResultArray += $UserObject
        }
    }
}

If($All.IsPresent) {    #param ALL

    $ResultArray = Get-MsolUser -all: $true | Select-Object -Property "DisplayName","UserPrincipalName",@{Name="MFA Status"; Expression= {
        if( $_.StrongAuthenticationRequirements.State -ne $null) {$_.StrongAuthenticationRequirements.State} 
            else { "Disabled"}}}, Country
    $count = $ResultArray.count         
} 


# Export to CSV and show result. Disconnect from cloud
#------------------------------------------------------
ExportToCSV $ResultArray $CSVPath $ExportCSV
ShowScriptResult $ExportCSV
Write-Host User count that matched criteria $count -ForegroundColor Yellow
DisconnectFromMSOLService
#end of script

