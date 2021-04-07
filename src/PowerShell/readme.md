# Nicks Teams Membership Synch PowerShell Script

Last update: 2021-04-06
Last updated by: Nicholas Rogoff

This script allows an owner of a Team to add or synchonize all the members of an Azure Active Directory group.

## Requirements

- The person or service principle executing the script must be an Owner of the Team and have access to the Enterprise Azure Active Directory.
- To execute the script you require the following:
  - The latest version of PowerShell
  - Latest version of MSOnline PowerShell module
  - Latest version of MicrosoftTeams PowerShell module

---

## Instructions

### Getting Started - Install PowerShell Modules

You will need to ensure you have installed the required modeules.

1. Open PowerShell with **'Run as Administrator'**
2. Run 

``` powershell
Install-Module -Name MSOnline
```

3. Then run 

``` powershell
Install-Module -Name MicrosoftTeams
```
---

### Logins

You will need to Login twice in your PowerShell session, once into the Azure AD and another for Teams.

1. Run the following 

``` powershell
Connect-MsolService
```

2. This will pop up a login. Use your mydomain user and MFA to login.
3. Then run 

``` powershell
Connect-MicrosoftTeams
```

4. This will pop up another Login screen and login as usual with your mydomain credentials.

### Run the script

The easiest way to run the script is to navigate in PowerShell to the folder where you copied it too.

Then to run the script just use

```
.\Sync-Team-Members-With-AD-Group.ps1 ...
```

whith the options that suit you.

The script has several options as follows:

``` powershell
Sync-Team-Members-With-AD-Group.ps1 [-AzureADGroupName] <string> [[-ADDomain] <string>] [[-TeamDisplayName] <string>] [[-TeamGroupId] <string>] [[-MaxUsers] <int>] [[-RemoveInvalidTeamMembers] <bool>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

---

### Paramters

#### **AzureADGroupName**
This is the Name of the Active Directory group you wish to populate the Team with. The name of the Azure AD Group (not including the @domain part). This needs to be specific enough to ensure that, when combined with the ADDomain, that only one group matches

#### **ADDomain**
[Mandatory] The AD Domain (after the @ symbol). 

#### **TeamDisplayName**
[Optional] The display name of the Team who's membership you wish to alter. Occationally this can results in more than one team being found. 
In those cases the  TeamGroupId parameter must be used instead. One MUST be used, but not both.

#### **TeamGroupId**
[Optional] Do not use this if using the TeamDisplayName. This can be used instead of the TeamDisplayName in cases where an exact match can't be found.  Group Id of the Team can be found in the Teams URL.

#### **MaxUsers**
[Optional] (Default = ALL). Max number of AD Group Users to process. This can be used to test a small batch. Set to 0 to process ALL members of the AD group

#### **RemoveInvalidTeamMembers**
[Optional] (Default = False). This indicates whether you want to remove members of the Team that are no longer in the AD Group. The default is to not remove members and Add-only. 
If you do want sync the membership of the Team exactly, e.g. to remove any Team Members that are not part of the AD group, then set this to True

---

### Examples


This example will add the ALL members of the 'MyADGroup@mydomain.com' AD Group (in no particular order) to the Team whose diaply name matches 'Software Engineering Talent Community'. Only missing members will be added. It will not remove any Team members.

``` powershell
.\Sync-Team-Members-With-AD-Group.ps1  -AzureADGroupName "MyADGroup" -TeamDisplayName "Software Engineering Talent Community"
```

---

This example will add the first 10 members of the 'MyADGroup@accenture.com' AD Group (in no particular order) to the Team whose diaply name matches 'Software Engineering Talent Community'. Only missing members will be added. It will not remove any Team members.

``` powershell
.\Sync-Team-Members-With-AD-Group.ps1  -AzureADGroupName "MyADGroup" -ADDomain "accenture.com"  -TeamDisplayName "Software Engineering Talent Community" -MaxUsers 10 -RemoveInvalidTeamMembers $false
```

---

This example will add the ALL members of the 'MyADGroup@mydomain.com' AD Group (in no particular order) to the Team whose Team Group Id is '679c179b-1234-1234-1234-27acee64b6d8'. Only missing members will be added. It will not remove any Team members. The full verbose output will be output.

``` powershell
.\Sync-Team-Members-With-AD-Group.ps1  -AzureADGroupName "MyADGroup" -TeamGroupId "679c179b-1234-1234-1234-27acee64b6d8" -RemoveInvalidTeamMembers $false -Verbose
```

---

This example will add the ALL members of the 'MyADGroup@mydomain.com' AD Group (in no particular order) to the Team whose Team Group Id is '679c179b-1234-1234-1234-27acee64b6d8'. All team members that are no longer in the AD group will be REMOVED.

``` powershell
.\Sync-Team-Members-With-AD-Group.ps1  -AzureADGroupName "MyADGroup" -TeamGroupId "679c179b-1234-1234-1234-27acee64b6d8" -RemoveInvalidTeamMembers $true
```