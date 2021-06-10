# PowerShell scripts for Microsoft Cloud
PowerShell scripts to manage Microsoft Cloud - Azure and Office

I intend to start with basic AzureAD and MS Office365 reports.
Eventually, I hope to include scripts that will not only report on, but also have ability to modify AD and Office accounts.


## Repository will be populated on demand as requests come from the clinet and auditors.

First Script:

AzureADUsersReport.ps1

This script assumes that AzureAD module is already installed. It returns one basic report with only few AzureAD user properties. They are:        "DisplayName","UserPrincipalName","DirSyncEnabled","ObjectType","UserType","City","State","Country"

At this time script has no parameters. I am planning to include few parametres in the next version. I suppose that default properties will change as well, once parameters for id/password and properties are added.


