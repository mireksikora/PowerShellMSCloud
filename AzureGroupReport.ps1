<# 
Author:         Mirek Sikora
Date:           2021/06/16 (YYYY/MM/DD)
Description:    Get Azure Group Report
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
#------------------------------

If($GroupTypes.IsPresent) {                 #default
    $groups=Get=ZaureADGroup -All $true
}



If($GroupMembers.IsPresent) {
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




If($GroupOwners.IsPresent) {
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
#--------------------------------------
Write-Host "`nDisconnecting from Azure AD`n"
Disconnect-AzureAD | Out-Null

#end of script