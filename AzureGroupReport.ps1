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
    [string] $TenantName,
    [string] $UserName,
    [string] $Password
)

# Connect to Microsoft Cloud
#----------------------------
$RequiredModule="AzureAD" 
$Module=Get-Module -Name $RequiredModule -ListAvailable
If($Module.count -eq 0)
{
    Write-Host $RequiredModule module is not available -ForegroundColor Yellow
    $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
    If($Confirm -match "[yY]")
    {
        Install-Module $RequiredModule -Force -Scope AllUsers
    }
    else {
        Write-Host $RequiredModule module is required to connect to Microsoft Cloud. Please install module using Install-Module comlet.
        Exit
    }
}

Import-Module $RequiredModule

If(($UserName -ne "") -and ($Password -ne ""))
{
    $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
    $CloudCred = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
}
else {
    $CloudCred = Get-Credential    
}

Connect-AzureAD -Credential $CloudCred

# Set output path and file name
#-------------------------------
If($TenantName -eq "")
{
Write-Host "Exporting tenant Azure report to CSV file"
$TenantName = Read-Host -Prompt "`nEnter tenant name"
}

$CSVpath = ".\CSVreports"
If (!(test-path $CSVpath))
{
    New-Item -ItemType Directory -Path $CSVpath | Out-Null
}

$ScriptName = $MyInvocation.MyCommand.Name
$FileName = $ScriptName.trim(".ps1")
$ExportCSV="$CSVpath\$FileName-$TenantName-$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"

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
            Write-Progress -Activity "Processing group count: $GCount   group name: $GDisplayName"
    
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
        Write-Progress -Activity "Processing group count: $GCount   group name: $GDisplayName"
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
        Write-Progress -Activity "Processing group count: $GCount   group name: $GDisplayName"
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



# Export report to CSV file
#---------------------------

If (test-path $CSVpath)
{
    $ResultArray | Export-Csv -Path $ExportCSV
}
else {
    Write-Host `nFolder $CSVpath not found -ForegroundColor Yellow
}



# Show report path and file name
#--------------------------------
If (test-path $ExportCSV -PathType Leaf)
{
    Write-Host CSV file saved to: $ExportCSV
    Write-Host `nReport generated successfully -ForegroundColor Green
}
else {
    Write-Host `nError: Unable to create file $ExportCSV -ForegroundColor Yellow
}

# Disconnect from Microsoft Azure AD
#------------------------------------
Write-Host "`nDisconnecting from Azure AD`n"
Disconnect-AzureAD | Out-Null

#end of script