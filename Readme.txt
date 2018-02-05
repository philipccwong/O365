Requirement 

1, Windows machine with Powershell enabled
2, Local Admin Account (or account with permission to run Powershell script)
3, Admin account with access to https://<Tenant Name>-admin.sharepoint.com

To Run (Site Extract):

Extract the package to C:\ 

1, Run Powershell in Windows (As Admin)
2, Type > Set-ExecutionPolicy RemoteSigned
3, Run > C:\SharePoint_PowerShell\O365\O365Script.ps1
a, Input admin url (i.e. https://colligo-admin.sharepoint.com)
b, Input O365 admin account (i.e. o365_admin@colligo.com)
c, Input O365 admin password

Result will be locate under C:\SharePoint_PowerShell\O365\O365_OutPut_Timestamp.html

To Run (Termset):

1, Run Powershell in Windows (As Admin)
2, Type > Set-ExecutionPolicy RemoteSigned (can skip)
3, Run > C:\SharePoint_PowerShell\O365\O365_Termset.ps1
a, Input admin url (i.e. https://colligo-admin.sharepoint.com)
b, Input O365 admin account (i.e. o365_admin@colligo.com)
c, Input O365 admin password

Result will be locate under C:\SharePoint_PowerShell\O365\O365_Termset_Timestamp.html