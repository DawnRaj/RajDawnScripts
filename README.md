# RajDawnScripts
Scripts for Active Directory Management.

Remove-SpecialChar.ps1 script can be used to remove certain special characters from the display name attribute in an Active Directory/Exchange enviornment.

This script uses the ActiveRolesManagementShell module which comes by default with Quest Active Roles management product.
https://www.oneidentity.com/products/active-roles/

Standalone older version of this shell can be downloaded from
https://www.powershelladmin.com/wiki/Quest_ActiveRoles_Management_Shell_Download


Input-Example-Remove-SpecialCharacter.csv is a sample "CSV" input file.There are three mandatory values required.
1. DisplayName of the user identity
2. Alias/MailNickName property of the user identity.
3. Domain FQDN of the user identitiy (To be provided during the script runtime)