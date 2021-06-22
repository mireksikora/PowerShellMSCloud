# PowerShell scripts for Microsoft Cloud
PowerShell scripts to manage Microsoft Cloud - Azure and Office

I intend to start with basic AzureAD and MS Office365 reports.
Eventually, I hope to include scripts that will not only report on, but also have ability to modify AD and Office accounts.


## Repository will be populated on demand as requests come from the clinet and auditors.

First Script:

AzureADUsersReport.ps1

This script assumes that AzureAD module is already installed. It returns one basic report with only few AzureAD user properties. They are:        "DisplayName","UserPrincipalName","DirSyncEnabled","ObjectType","UserType","City","State","Country"

At this time script has no parameters. I am planning to include few parametres in the next version. I suppose that default properties will change as well, once parameters for id/password and properties are added.

Parameters have been added now.

All scripts require at least three parameters: username, password, tenant name. Parameters can be passed during call or the script will prompt for them during runtime. Additional optional parameter can be passed during call, and it is based on report type associated with PS script. For example you can call AzureADUsersReport.ps1 with -Location as parameter. It will generate user report with fields associated with location, like city, state, zip code, etc. Another report for example AzureGroupReport.ps1 wil have -GroupTypes and -GroupMembers as available parameters. Only one optional parameter (report type) can be passed to a script. Once scripts executes successfully, it will gernerate CSV file in the following format:
".\CSVreports\ScriptName-TenantName-YYYYMMMDD_HHMM.csv"
Example script name: ".\CSVreports\AzureADUsersReport-Oracle-2021_0334.csv"

AzureADUsersReport.ps1
----------------------
Fully functional v1.0 in now completed. Call examples with and without parameters:

a) Call without parameters witll prompt for ID/Password, and tenant name. By detault if no parameters are passted, report it will run showing all fields, similar to running with -ALL parameter

./AzureADUsersReport.ps1

b) Pass parameters for report type, tenant name or id/password

./AzureADUsersReport.ps1 -Basic 
./AzureADUsersReport.ps1 -Account
./AzureADUsersReport.ps1 -Location
./AzureADUsersReport.ps1 -All
./AzureADUsersReport.ps1 -Location -tenantname "Oracle"
./AzureADUsersReport.ps1 -Account -tenantname "Oracle" -UserName adminuser@domain.com -Password XXX




AzureGroupReport.ps1
--------------------
Fully functional v1.0 is now completed. Call examples with and without parameters:

./AzureGroupReport.ps1
./AzureGroupReport.ps1 -GroupTypes (default if no parameters are passed)
./AzureGroupReport.ps1 -GroupMembers
./AzureGroupReport.ps1 -GroupMembers -tenantname "IBM"
./AzureGroupReport.ps1 -GroupOwners -UserName adminuser@domain.com -Password YYY


