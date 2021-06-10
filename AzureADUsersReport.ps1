<# 
Author:         Mirek Sikora
Date:           2021/06/07 (YYYY/MM/DD)
Description:    Get AzureAD user attributes and save them to CSV file
#>

#Install-module PowerShellGet -Force -Scope AllUsers
#Install-Module AzureAD -Force -Scope AllUsers
#Update-Module AzureAD

# Connect to Microsoft Azure AD
#---------------------------------
Import-Module AzureAD
$AzureADcred = Get-Credential
Connect-AzureAD -Credential $AzureADcred

# Get tenant name from the user
#---------------------------------
Write-Host "Exporting tenant Azure AD user list to CSV file"
$TenantName = Read-Host -Prompt "`nEnter tenant name"

# Build report path and CSV file name
#--------------------------------------
$CSVpath = ".\CSVreports"

If (!(test-path $CSVpath))
{
    New-Item -ItemType Directory -Path $CSVpath | Out-Null
    Write-Host `nFolder $CSVpath created
}


$ScriptName = $MyInvocation.MyCommand.Name
#Write-Host Script name = $ScriptName
$FileName = $ScriptName.trim(".ps1")
#Write-Host FileName = $FileName
$ScriptPath =  $MyInvocation.MyCommand.Definition 
#Write-Host Script path = $ScriptPath.Trim($ScriptName)

$ExportCSV="$CSVpath\$FileName-$TenantName-$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"

# Define report fields
#-----------------------
$AADproperties ="DisplayName","UserPrincipalName","DirSyncEnabled","ObjectType","UserType","City","State","Country"

# Generate report and export to CSV file
#----------------------------------------
If (test-path $CSVpath)
{
Get-AzureADUser | Select-Object -Property $AADproperties | Export-Csv -Path $ExportCSV
}


# Show generated report name and absolute path location
#-------------------------------------------------------
Write-Host ""
Write-Host Saving CSV file to: $ExportCSV
$AbsolutePath = $ScriptPath.Trim($ScriptName) + $ExportCSV.trim(".\")
If (test-path $ExportCSV -PathType Leaf)
{
    Write-Host Report generated successfully
}
else {
    Write-Host Error: Unable to create file to $ExportCSV
}

# Disconnect from Microsoft Azure AD
#--------------------------------------
Write-Host "`nDisconnecting from Azure AD`n"
Disconnect-AzureAD

#end of script

