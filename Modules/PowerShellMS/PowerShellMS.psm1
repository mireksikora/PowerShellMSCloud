# Save this module in your Windows user installed PowerShell location under your profile 
# or
# Windows 10 default module location (global for all users) 
# C:\Windows\System32\WindowsPowerShell\v1.0\Modules\PowerShellMS

function ConnectToMsolService {
#===============================

param
(
    [Parameter(Mandatory = $false)]
    [string] $UserName,
    [string] $Password
)

# Connect to Microsoft Cloud
#----------------------------
$RequiredModule="MSOnline" 
$Module=Get-Module -Name $RequiredModule -ListAvailable
If($Module.count -eq 0)
{
    Write-Host $RequiredModule module is not available -ForegroundColor Yellow
    $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No
    If($Confirm -match "[yY]") {
        Install-Module $RequiredModule -Force -Scope AllUsers
    }
    else {
        Write-Host $RequiredModule module is required to connect to Microsoft Cloud. Please install module using Install-Module comlet.
        Exit
    }
}

Import-Module $RequiredModule

If(($UserName -ne "") -and ($Password -ne "")){
    $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
    $CloudCred = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
    }
    else {
        $CloudCred = Get-Credential    
}

Connect-MsolService -Credential $CloudCred
}

function DisconnectFromMSOLService {
    Write-host "To disconnect from MSOLService close your browser or PowerShell script window"

}

function ConnectToAzureAD {
#===========================

param
(
    [Parameter(Mandatory = $false)]
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
    If($Confirm -match "[yY]") {
        Install-Module $RequiredModule -Force -Scope AllUsers
    }
    else {
        Write-Host $RequiredModule module is required to connect to Microsoft Cloud. Please install module using Install-Module comlet.
        Exit
    }
}

Import-Module $RequiredModule

If(($UserName -ne "") -and ($Password -ne "")){
    $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
    $CloudCred = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
    }
    else {
        $CloudCred = Get-Credential    
}

Connect-AzureAD -Credential $CloudCred
}

function DisconnectFromAzureAD {
#===============================
 
# Disconnect from Microsoft Azure AD
#--------------------------------------
Write-Host "`nDisconnecting from Azure AD`n"
Disconnect-AzureAD | Out-Null
}


function SetOutputPathFilename {
#================================
    param
(
    [Parameter(Mandatory = $false)]
    [string] $TenantName
)

# Set output path and file name
#-------------------------------
If($TenantName -eq "") {
    Write-Host "Exporting tenant Azure AD user list to CSV file"
    $TenantName = Read-Host -Prompt "`nEnter tenant name"
}

$CSVpath = ".\CSVreports"
If (!(test-path $CSVpath)) {
    New-Item -ItemType Directory -Path $CSVpath | Out-Null
}

#$ScriptName = $MyInvocation.ScriptName
$ScriptName = Split-Path $MyInvocation.PSCommandPath -Leaf
$FileName = $ScriptName.trim(".ps1")
$ExportCSV="$CSVpath\$FileName-$TenantName-$((Get-Date -format yyyyMMdd_hhmm).ToString()).csv"

Return $CSVpath, $ExportCSV
}

function ExportToCSV {
#======================
    param (
        [Parameter(Mandatory = $true)]
        [array] $ResultArray,
        [string] $CSVPath,
        [string] $ExportCSV
        
    )
# Export to CSV file
#--------------------
    If (test-path $CSVPath) {
        $ResultArray | Export-Csv -Path $ExportCSV
        }
    else {
        Write-Host `nFolder $CSVPath not found -ForegroundColor Yellow
    }
}


function ShowScriptResult {
#==========================
        param
(
    [Parameter(Mandatory = $true)]
    [string] $Filename
)

# Show report path and file name
#--------------------------------
If (test-path $Filename -PathType Leaf) {
    Write-Host `nCSV file saved to: $Filename
    Write-Host Report generated successfully -ForegroundColor Green
    }
    else {
        Write-Host `nError: Unable to create file $Filename -ForegroundColor Yellow
    }
}