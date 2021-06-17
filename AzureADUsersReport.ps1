<# 
Author:         Mirek Sikora
Date:           2021/06/07 (YYYY/MM/DD)
Description:    Get AzureAD user attributes and save them to CSV file
#>

#Install-module PowerShellGet -Force -Scope AllUsers
#Install-Module AzureAD -Force -Scope AllUsers
#Update-Module AzureAD


#Define parameters
#-------------------
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false)]
    [switch] $Basic, #default
    [switch] $Account,
    [switch] $Location,
    [switch] $All,
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
Write-Host "Exporting tenant Azure AD user list to CSV file"
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

# Set report fields based in input parameters
#-----------------------------------------------

$ReportProperties ="DisplayName","UserPrincipalName","ObjectType","UserType","AccountEnabled","DirSyncEnabled","UserState"    #basic

If($Account.IsPresent) {
$ReportProperties ="DisplayName","UserPrincipalName","AccountEnabled","DirSyncEnabled","Mail","UserState","UserStateChangedOn","CreationType","AssignedLicenses","DeletionTimestamp"    #account
}

If($Location.IsPresent) {
$ReportProperties ="DisplayName","UserPrincipalName","CompanyName","PhysicalDeliveryOfficeName","Department","StreetAddress","City","State","PostalCode","Country","TelephoneNumber","FacsimileTelephoneNumber","UsageLocation","PreferredLanguage" #location
}

If($All.IsPresent) {
$ReportProperties ="*" #all
}

# Generate report and export to CSV file
#----------------------------------------
If (test-path $CSVpath)
{
Get-AzureADUser | Select-Object -Property $ReportProperties | Export-Csv -Path $ExportCSV
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

