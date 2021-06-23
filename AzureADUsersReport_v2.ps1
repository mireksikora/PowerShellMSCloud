<# 
Author:         Mirek Sikora
Date:           2021/06/07 (YYYY/MM/DD)
Description:    Get AzureAD user attributes and save them to CSV file
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
    [switch] $Basic, 
    [switch] $Account,
    [switch] $Location,
    [switch] $All,         #default
    [switch] $MFA,
    [string] $TenantName,
    [string] $UserName,
    [string] $Password
)

# Connect to cloud and define variables 
#---------------------------------------
import-module PowerShellMS
If($MFA -eq $false) {
    ConnectToAzureAD $UserName $Password
}   Else {
        import-module AzureAD
        Connect-AzureAD
}
$CSVPath,$ExportCSV = SetOutputPathFilename $TenantName

# Set report result based on input parameters
#-----------------------------------------------

If(($Basic -eq $false) -and ($Account -eq $false) -and ($Location -eq $false) -and ($All -eq $false)){
    $All = $true #default
}

If($Basic.IsPresent) {     #basic
    $ReportProperties ="*"    
    $users = Get-AzureADUser -All $true | Select-Object $ReportProperties
    $TotalUsers = $users.count

    $ResultArray=@()
    ForEach($user in $users) {
        Write-Progress -Activity "Processing user $count of $TotalUsers"
        $count += 1

        $UserObject = New-Object PSCustomObject
        $creationdate = (Get-AzureADUserExtension -Objectid $user.Objectid).Get_Item('createdDateTime')
        #$objectdatatype = (Get-AzureADUserExtension -Objectid $user.Objectid).Get_Item('odata.type')

        $UserObject | Add-Member -MemberType NoteProperty -name "Display Name" -value $user.DisplayName
        $UserObject | Add-Member -MemberType NoteProperty -name "UserPrincipalName" -value $user.UserPrincipalName
        $UserObject | Add-Member -MemberType NoteProperty -name "Creation Date" -value $creationdate
        #$UserObject | Add-Member -MemberType NoteProperty -name "Object Data Type" -value $objectdatatype
        $UserObject | Add-Member -MemberType NoteProperty -name "Creation Type" -value $user.CreationType
        $UserObject | Add-Member -MemberType NoteProperty -name "Object Type" -value $user.ObjectType
        $UserObject | Add-Member -MemberType NoteProperty -name "User Type" -value $user.UserType
        $UserObject | Add-Member -MemberType NoteProperty -name "Account Enabled" -value $user.AccountEnabled
        $UserObject | Add-Member -MemberType NoteProperty -name "DirSyncEnabled" -value $user.DirSyncEnabled
        $UserObject | Add-Member -MemberType NoteProperty -name "UserState" -value $user.UserState
        $ResultArray += $UserObject
    }
}

If($Account.IsPresent) {   #account
    $ReportProperties ="*"    
    $users = Get-AzureADUser -All $true | Select-Object $ReportProperties
    $TotalUsers = $users.count

    $ResultArray=@()
    ForEach($user in $users) {
        Write-Progress -Activity "Processing user $count of $TotalUsers"
        $count += 1

        $UserObject = New-Object PSCustomObject
        $licenses = get-AzureADUser -objectid $user.objectid | Select-Object -expandproperty AssignedLicenses
        $LicenseCount = $licenses.count

        $UserObject | Add-Member -MemberType NoteProperty -name "Display Name" -value $user.DisplayName
        $UserObject | Add-Member -MemberType NoteProperty -name "UserPrincipalName" -value $user.UserPrincipalName
        $UserObject | Add-Member -MemberType NoteProperty -name "Account Enabled" -value $user.AccountEnabled
        $UserObject | Add-Member -MemberType NoteProperty -name "DirSyncEnabled" -value $user.DirSyncEnabled
        $UserObject | Add-Member -MemberType NoteProperty -name "Mail" -value $user.Mail
        $UserObject | Add-Member -MemberType NoteProperty -name "User State" -value $user.UserState
        $UserObject | Add-Member -MemberType NoteProperty -name "User State Changed On" -value $user.UserStateChangedOn
        $UserObject | Add-Member -MemberType NoteProperty -name "Number of Licenses Assigned" -value $LicenseCount
        $UserObject | Add-Member -MemberType NoteProperty -name "Deletion Timestamp" -value $user.DeletionTimestamp
        $ResultArray += $UserObject
    }
}

If($Location.IsPresent) {  #location
    $ReportProperties = "DisplayName","UserPrincipalName","CompanyName","PhysicalDeliveryOfficeName","Department","StreetAddress",
    "City","State","PostalCode","Country","TelephoneNumber","FacsimileTelephoneNumber","UsageLocation","PreferredLanguage" #location
    $ResultArray = Get-AzureADUser | Select-Object -Property $ReportProperties
     
}

If($All.IsPresent) {       #all
    $ReportProperties ="DeletionTimestamp","ObjectType","AccountEnabled","AgeGroup","City","CompanyName",
    "ConsentProvidedForMinor","Country","CreationType","Department","DirSyncEnabled","DisplayName","FacsimileTelephoneNumber",
    "GivenName","IsCompromised","ImmutableId","JobTitle","LastDirSyncTime","LegalAgeGroupClassification","Mail","MailNickName",
    "Mobile","OnPremisesSecurityIdentifier","PasswordPolicies","PasswordProfile","PhysicalDeliveryOfficeName","PostalCode",
    "PreferredLanguage","RefreshTokensValidFromDateTime","ShowInAddressList","SipProxyAddress","State","StreetAddress","Surname",
    "TelephoneNumber","UsageLocation","UserPrincipalName","UserState","UserStateChangedOn","UserType" 
    $ResultArray = Get-AzureADUser | Select-Object -Property $ReportProperties
}

# Export to CSV and show result. Disconnect from cloud.
#------------------------------------------------------
ExportToCSV $ResultArray $CSVPath $ExportCSV
ShowScriptResult $ExportCSV
DisconnectFromAzureAD
#end of script

