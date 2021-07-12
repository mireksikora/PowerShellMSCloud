<# 
Author:         Mirek Sikora
Date:           2021/07/10 (YYYY/MM/DD)
Description:    Generate Cloud report and save it to CSV file
#>

#Define parameters
#-------------------
[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false)]
    [switch] $All,                  # All Mailboxes
    [switch] $SharedMBOnly,         # Shared Mailboxes only
    [switch] $UserMBOnly,           # User Mailboxes only
    [switch] $InactiveMBOnly,      # Not Active Mailboxes (soft delete or disabled)
    [switch] $ActiveMBOnly,         # Active Mailboxes
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
ConnectToExchangeOnline $MFAEnabled $UserName $Password 
$CSVPath,$ExportCSV = SetOutputPathFilename $TenantName

# Set report result based on input parameters
#-----------------------------------------------

If(($All -eq $false) -and ($SharedMBOnly -eq $false) -and ($UserMBOnly -eq $false) -and
     ($InactiveOnly -eq $false) -and ($ActiveOnly -eq $false)) {
    $All = $true        #default
}

If($All.IsPresent) {                # Parameter All mailboxes
    $ResultArray = @()
    $MailboxProperties="DisplayName","UserPrincipalName","RecipientTypeDetails"
    $mailboxes = Get-EXOMailbox -resultsize unlimited | select-object $MailboxProperties 
    $MBtotal = $mailboxes.count

    Write-host Total mailbox count: $MBtotal

    $UserObject = New-Object PSCustomObject
    $MailboxStatisticProperties = "DisconnectReason","ItemCount","TotalItemSize","DeletedItemCount","TotalDeletedItemSize","MailboxTypeDetail"

    foreach ($mailbox in $mailboxes) {
        $mbstat = Get-EXOMailboxStatistics -userprincipalname $mailbox.UserPrincipalName | select-object $MailboxStatisticProperties
        $MBcount += 1
        $MBDisplayName = $mailbox.DisplayName

        Write-Progress -Activity "Processing mailbox count: $MBcount of $MBtotal  Name: $MBDisplayName"
    
        $UserObject | Add-Member -MemberType NoteProperty -name "Mailbox Name" -value $mailbox.DisplayName -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Email Address" -value $mailbox.UserPrincipalName -force
        
        If([string]::IsNullOrEmpty($mbstat.DisconnectReason)) {   # if field is Null or empty
            $Status = "Active"    
            } else { $Status = $mbstat.DisconnectReason }

        $UserObject | Add-Member -MemberType NoteProperty -name "Status" -value $Status -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Mailbox Type" -value $mailbox.RecipientTypeDetails -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Item Count" -value $mbstat.ItemCount -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Deleted Item Count" -value $mbstat.DeletedItemCount -force
        $ResultArray = $UserObject
        $ResultArray | Export-Csv -Path $ExportCSV -Notype -Append         
    }
}

If($SharedMBOnly.IsPresent) {       # Parameter Shared mailbox only
    
    $ResultArray = @()
    $MailboxProperties="DisplayName","UserPrincipalName","RecipientTypeDetails"
    $mailboxes = Get-EXOMailbox -resultsize unlimited | select-object $MailboxProperties | where-object {
        $_.RecipientTypeDetails -eq "SharedMailbox"
    }
    $MBtotal = $mailboxes.count
    Write-host Shared mailbox count: $MBtotal

    $UserObject = New-Object PSCustomObject
    $MailboxStatisticProperties = "DisconnectReason","ItemCount","TotalItemSize","DeletedItemCount","TotalDeletedItemSize","MailboxTypeDetail"

    foreach ($mailbox in $mailboxes) {
        $mbstat = Get-EXOMailboxStatistics -userprincipalname $mailbox.UserPrincipalName | select-object $MailboxStatisticProperties
        $MBcount += 1
        $MBDisplayName = $mailbox.DisplayName

        Write-Progress -Activity "Processing mailbox count: $MBcount of $MBtotal  Name: $MBDisplayName"
    
        $UserObject | Add-Member -MemberType NoteProperty -name "Mailbox Name" -value $mailbox.DisplayName -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Email Address" -value $mailbox.UserPrincipalName -force

        If([string]::IsNullOrEmpty($mbstat.DisconnectReason)) {   # if field is Null or empty
            $Status = "Active"    
            } else { $Status = $mbstat.DisconnectReason }
            
        $UserObject | Add-Member -MemberType NoteProperty -name "Status" -value $Status -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Mailbox Type" -value $mailbox.RecipientTypeDetails -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Item Count" -value $mbstat.ItemCount -force    
        $UserObject | Add-Member -MemberType NoteProperty -name "Item Size" -value $mbstat.TotalItemSize -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Deleted Item Count" -value $mbstat.DeletedItemCount -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Deleted Item Size" -value $mbstat.TotalDeletedItemSize -force
        $ResultArray = $UserObject
        $ResultArray | Export-Csv -Path $ExportCSV -Notype -Append        
    }
}

If($UserMBOnly.IsPresent) {         # parameter User mailbox only 

    $ResultArray = @()
    $MailboxProperties="DisplayName","UserPrincipalName","RecipientTypeDetails"
    $mailboxes = Get-EXOMailbox -resultsize unlimited | select-object $MailboxProperties | where-object {
        $_.RecipientTypeDetails -eq "UserMailbox"
    }
    $MBtotal = $mailboxes.count
    Write-host User mailbox count: $MBtotal

    $UserObject = New-Object PSCustomObject
    $MailboxStatisticProperties = "DisconnectReason","ItemCount","TotalItemSize","DeletedItemCount","TotalDeletedItemSize","MailboxTypeDetail"

    foreach ($mailbox in $mailboxes) {
        $mbstat = Get-EXOMailboxStatistics -userprincipalname $mailbox.UserPrincipalName | select-object $MailboxStatisticProperties | where-object {
            $_.RecipientTypeDetails -eq "SharedMailbox"
        }
        $MBcount += 1
        $MBDisplayName = $mailbox.DisplayName

        Write-Progress -Activity "Processing mailbox count: $MBcount of $MBtotal  Name: $MBDisplayName"
    
        $UserObject | Add-Member -MemberType NoteProperty -name "Mailbox Name" -value $mailbox.DisplayName -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Email Address" -value $mailbox.UserPrincipalName -force

        If([string]::IsNullOrEmpty($mbstat.DisconnectReason)) {   # if field is Null or empty
            $Status = "Active"    
            } else { $Status = $mbstat.DisconnectReason }
            
        $UserObject | Add-Member -MemberType NoteProperty -name "Status" -value $Status -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Mailbox Type" -value $mailbox.RecipientTypeDetails -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Item Count" -value $mbstat.ItemCount -force
        $UserObject | Add-Member -MemberType NoteProperty -name "Deleted Item Count" -value $mbstat.DeletedItemCount -force
        $ResultArray = $UserObject
        $ResultArray | Export-Csv -Path $ExportCSV -Notype -Append 
        
    }
}

If($InactiveMBOnly.IsPresent) {      # Not Active Mailboxes (soft delete or disabled)

    Write-host "starting query, please wait.."
    $MailboxStatisticProperties = "DisplayName","DisconnectReason","MailboxTypeDetail","ItemCount","DeletedItemCount"
    $ResultArray = Get-EXOMailbox -resultsize unlimited | Get-EXOMailboxStatistics -properties $MailboxStatisticProperties | where-object {
        ( -not ([string]::IsNullOrEmpty($_.DisconnectReason)))
    }

    $InactiveCount = $ResultArray.count

    If($InactiveCount -gt 0) {
        Write-Host "Inactive mailboxes count: $InactiveCount" -ForegroundColor Yellow
        ExportToCSV $ResultArray $CSVPath $ExportCSV
    } else {
        Write-Host "Inactive mailboxes count: $InactiveCount" -ForegroundColor Yellow
    }
} 

If($ActiveMBOnly.IsPresent) {         # Active Mailboxes (soft delete or disabled)

    Write-host "starting query, please wait.."
    $MailboxStatisticProperties = "DisplayName","DisconnectReason","MailboxTypeDetail","ItemCount","DeletedItemCount"
    $ResultArray = Get-EXOMailbox -resultsize unlimited | Get-EXOMailboxStatistics -properties $MailboxStatisticProperties | where-object {
       ([string]::IsNullOrEmpty($_.DisconnectReason))
    }

    $ActiveCount = $ResultArray.count

    If($ActiveCount -gt 0) {
        Write-Host "Active mailboxes count: $ActiveCount" -ForegroundColor Yellow
        ExportToCSV $ResultArray $CSVPath $ExportCSV
    } else {
        Write-Host "Active mailboxes count: $ActiveCount" -ForegroundColor Yellow
    }
} 

# Show result. Disconnect from cloud
#------------------------------------------------------
If( -not ([string]::IsNullOrEmpty($ResultArray))) {
    ShowScriptResult $ExportCSV
}

DisconnectFromExchangeOnline

$endTime = Get-Date -format G
Write-Host "Script completed at $endTime."
Write-host "Time for the full run was: $( New-TimeSpan $startTime $endTime)."

#end of script

