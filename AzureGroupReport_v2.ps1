<# 
Author:         Mirek Sikora
Date:           2021/06/16 (YYYY/MM/DD)
Description:    Azure Group Reports
#>

#Define parameters
#-------------------
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false)]
    [switch] $GroupTypes, #default
    [switch] $GroupMembers,
    [switch] $GroupOwners,
    [switch] $MFA,
    [string] $TenantName,
    [string] $UserName,
    [string] $Password
)

# Connect to cloud and define variables 
#---------------------------------------
import-module PowerShellMS

If($MFA.IsPresent) { $MFAEnabled = $true } else { $MFAEnabled = $false}
ConnectToAzureAD $MFAEnabled $UserName $Password 
$CSVPath,$ExportCSV = SetOutputPathFilename $TenantName

# Run report based on input criteria
#------------------------------------
If(($GroupTypes -ne $true) -and ($GroupMembers -ne $true) -and ($GroupOwners -ne $true))
    { 
        $GroupTypes = $true #set default report in none selected
    }

    If($GroupTypes.IsPresent) {      # Group types report       
        $GroupProperty = "ObjectID","ObjectType","Description","DisplayName","DirSyncEnabled","LastDirSyncTime","MailEnabled","Mail","MailNickName","SecurityEnabled"
        $groups=Get-AzureADGroup -All $true | Select-Object $GroupProperty
    
        $ResultArray=@()
        ForEach ($group in $groups) {
            $GCount += 1
            $GDisplayName = $group.DisplayName
            $GTotal = $groups.count
            Write-Progress -Activity "Processing group count: $GCount of $GTotal group name: $GDisplayName"
    
            $GID = $group.ObjectID
            $MSG = Get-AzureADMSGroup -Id $GID | Select-Object -Property "GroupTypes", "Visibility"
            $O365Group = $MSG.GroupTypes[0]
            if($null -ne $O365Group) {
                $O365Group = $O365Group.replace('Unified','TRUE')
            }
            else {
                $O365Group = "FALSE"
            }
      
            $UserObject = New-Object PSCustomObject
            $UserObject | Add-Member -MemberType NoteProperty -name "Object Type" -value $group.ObjectType
            $UserObject | Add-Member -MemberType NoteProperty -name "Description" -value $group.Description
            $UserObject | Add-Member -MemberType NoteProperty -name "Group Name" -value $group.DisplayName
            $UserObject | Add-Member -MemberType NoteProperty -name "DirSyncEnabled" -value $group.DirSyncEnabled
            $UserObject | Add-Member -MemberType NoteProperty -name "DirSyncTime" -value $group.LastDirSyncTime
            $UserObject | Add-Member -MemberType NoteProperty -name "Security Enabled" -value $group.SecurityEnabled
            $UserObject | Add-Member -MemberType NoteProperty -name "Mail Enabled" -value $group.MailEnabled
            $UserObject | Add-Member -MemberType NoteProperty -name "O365 Group" -value $O365Group
            $UserObject | Add-Member -MemberType NoteProperty -name "Visibility" -value $MSG.Visibility
            $UserObject | Add-Member -MemberType NoteProperty -name "Email" -value $group.Mail
            $UserObject | Add-Member -MemberType NoteProperty -name "Email NickName" -value $group.MailNickName
            $ResultArray += $UserObject
        }    
    }



If($GroupMembers.IsPresent) {        # Group members report
    $groups=Get-AzureADGroup -All $true
    $ResultArray=@()
    ForEach ($group in $groups) {
        $members = Get-AzureADGroupMember -ObjectId $Group.ObjectID -All $true
        $GCount += 1
        $GDisplayName = $group.DisplayName
        $GTotal = $groups.count

        Write-Progress -Activity "Processing group count: $GCount of $GTotal group name: $GDisplayName"
        foreach ($member in $members) {
                $UserObject = New-Object PSCustomObject
                $UserObject | Add-Member -MemberType NoteProperty -name "Group Name" -value $group.DisplayName
                $UserObject | Add-Member -MemberType NoteProperty -name "Member Name" -value $member.DisplayName
                $UserObject | Add-Member -MemberType NoteProperty -name "Member UPN" -value $member.UserPrincipalName
                $UserObject | Add-Member -MemberType NoteProperty -name "Member ObjectType" -value $member.ObjectType
                $UserObject | Add-Member -MemberType NoteProperty -name "Member UserType" -value $member.UserType
                $ResultArray += $UserObject
        }

    }
}




If($GroupOwners.IsPresent) {         # Group owners report
    $groups=Get-AzureADGroup -All $true
    $ResultArray=@()
    ForEach ($group in $groups) {
        $owners = Get-AzureADGroupOwner -ObjectId $Group.ObjectID -All $true
        $GCount += 1
        $GDisplayName = $group.DisplayName
        $GTotal = $groups.count

        Write-Progress -Activity "Processing group count: $GCount of $GTotal group name: $GDisplayName"
        foreach ($owner in $owners) {
                $UserObject = New-Object PSCustomObject
                $UserObject | Add-Member -MemberType NoteProperty -name "Group Name" -value $group.DisplayName
                $UserObject | Add-Member -MemberType NoteProperty -name "Owner Name" -value $owner.DisplayName
                $UserObject | Add-Member -MemberType NoteProperty -name "Owner UPN" -value $owner.UserPrincipalName
                $UserObject | Add-Member -MemberType NoteProperty -name "Owner ObjectType" -value $owner.ObjectType
                $UserObject | Add-Member -MemberType NoteProperty -name "Owner UserType" -value $owner.UserType
                $ResultArray += $UserObject
        }

    }
}


# Export to CSV and show result. Disconnect from cloud.
#------------------------------------------------------
ExportToCSV $ResultArray $CSVPath $ExportCSV
ShowScriptResult $ExportCSV
DisconnectFromAzureAD
#end of script
