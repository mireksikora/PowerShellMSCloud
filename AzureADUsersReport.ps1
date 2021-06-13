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
$RequiredModule=AzureAD 
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
$CloudCred = Get-Credential
Connect-AzureAD -Credential $CloudCred

# Get tenant name from the user
#---------------------------------
Write-Host "Exporting tenant Azure AD user list to CSV file"
$TenantName = Read-Host -Prompt "`nEnter tenant name"

# Set output path and report file name
#--------------------------------------
$CSVpath = ".\CSVreports"

If (!(test-path $CSVpath))
{
    New-Item -ItemType Directory -Path $CSVpath | Out-Null
    #Write-Host `nFolder $CSVpath created
}

$ScriptName = $MyInvocation.MyCommand.Name
$FileName = $ScriptName.trim(".ps1")
#$ScriptPath =  $MyInvocation.MyCommand.Definition 
#Write-Host Script path = $ScriptPath.Trim($ScriptName)

$ExportCSV="$CSVpath\$FileName-$TenantName-$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"

# Define report fields
#-----------------------
$ReportProperties ="DisplayName","UserPrincipalName","DirSyncEnabled","ObjectType","UserType","City","State","Country"

# Generate report and export to CSV file
#----------------------------------------
If (test-path $CSVpath)
{
Get-AzureADUser | Select-Object -Property $ReportProperties | Export-Csv -Path $ExportCSV
}
else {
    Write-Host `nUnable to create report path $CSVpath -ForegroundColor Yellow
}


# Show generated report path and file name
#------------------------------------------
#$AbsolutePath = $ScriptPath.Trim($ScriptName) + $ExportCSV.trim(".\")
If (test-path $ExportCSV -PathType Leaf)
{
    Write-Host $ExportCSV
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

