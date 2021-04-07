<#PSScriptInfo
.VERSION 1.2
.GUID 21a6ad93-df53-4a1a-82fd-4a902cb57350
.AUTHOR Nicholas Rogoff

.RELEASENOTES
Initial version.
#>
<# 
.SYNOPSIS 
Synchronizes Team membership with an Azure AD Group. 
 
.DESCRIPTION 
Loops through all members of an AD group and adds any missing users to the membership. 
The script then loops through all the existing Team members and removes any that are no longer in the AD Group.

NOTE: This script will NOT remove Owners.

PRE-REQUIREMENT
---------------
Install Teams Module (MS Online)
PS> Install-Module -Name MSOnline

Install Microsoft Teams cmdlets module
PS> Install-Module -Name MicrosoftTeams

Sign in to Teams and Azure AD using
Connect-MicrosoftTeams
Connect-MsolService

.INPUTS
None. You cannot pipe objects to this!

.OUTPUTS
None.

.PARAMETER AzureADGroupName
This is the Name of the Active Directory group you wish to populate the Team with. The name of the Azure AD Group (not including the @domain part). This needs to be specific enough to ensure that, when combined with the ADDomain, that only one group matches

.PARAMETER ADDomain
[Mandatory] The AD Domain (after the @ symbol). 

.PARAMETER TeamDisplayName
[Optional] The display name of the Team who's membership you wish to alter. Occationally this can results in more than one team being found. 
In those cases the  TeamGroupId parameter must be used instead. One MUST be used, but not both.

.PARAMETER TeamGroupId
[Optional] Do not use this if using the TeamDisplayName. This can be used instead of the TeamDisplayName in cases where an exact match can't be found.  Group Id of the Team can be found in the Teams URL.

.PARAMETER MaxUsers
[Optional] (Default = ALL). Max number of AD Group Users to process. This can be used to test a small batch. Set to 0 to process ALL members of the AD group

.PARAMETER RemoveInvalidTeamMembers
[Optional] (Default = False). This indicates whether you want to remove members of the Team that are no longer in the AD Group. The default is to not remove members and Add-only. 
If you do want sync the membership of the Team exactly, e.g. to remove any Team Members that are not part of the AD group, then set this to True

.NOTES
  Version:        1.2
  Author:         Nicholas Rogoff
  Creation Date:  2021-04-06
  Purpose/Change: Added more detailed help
   
.EXAMPLE 
PS> .\Sync-Team-Members-With-AD-Group.ps1 -AzureADGroupName "MyADGroup" -TeamDisplayName "My Team"
This will add all missing members of the AD Group to the Team

.EXAMPLE
PS> .\Sync-Team-Members-With-AD-Group.ps1 -AzureADGroupName "MyADGroup" -TeamDisplayName "My Team" -MaxUsers 10 -Verbose
This will add all missing members of the first 10 members AD Group to the Team and will output verbose details

.EXAMPLE 
PS> .\Sync-Team-Members-With-AD-Group.ps1 -AzureADGroupName "MyADGroup" -TeamDisplayName "My Team" -RemoveInvalidTeamMembers
This will add all missing members of the AD Group to the Team, and REMOVE any members of the Team that are not in the AD Group (Except for Team Owners)

#>
#---------------------------------------------------------[Script Parameters]------------------------------------------------------
[CmdletBinding(SupportsShouldProcess)]
Param(
  [Parameter(Mandatory = $true, HelpMessage = "This is the Name of the Active Directory group you wish to populate the Team with. The name of the Azure AD Group (not including the @domain part). This needs to be specific enough to ensure that, when combined with the ADDomain, that only one group matches", ValueFromPipeline)]
  [String] $AzureADGroupName,
  [Parameter(Mandatory = $true, HelpMessage = "The AD Domain (after the @ symbol)")]
  [String] $ADDomain,
  [Parameter(Mandatory = $false, HelpMessage = "The display name of the Team")]
  [string] $TeamDisplayName = "",
  [Parameter(Mandatory = $false, HelpMessage = "The Group Id of the Team")]
  [string] $TeamGroupId = "",
  [Parameter(Mandatory = $false, HelpMessage = "Max number of AD Group Users to process. Default is 0 (ALL)")]
  [int] $MaxUsers = 0,
  [Parameter(Mandatory = $false, HelpMessage = "Default = False. If you do want to remove any Team Members that are not part of the AD group, then set this to True")]
  [bool] $RemoveInvalidTeamMembers = $false
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Starting synchonisation") -ForegroundColor Blue

# Set Error Action to Silently Continue
$ErrorActionPreference = 'Continue'

if ($TeamDisplayName.Length -lt 1 -and ($TeamGroupId -lt 1)) {
  throw "!! Invalid Parameters !! - You must pass at least a Team Group Name or Team Group Id."
}

#----------------------------------------------------------[Declarations]----------------------------------------------------------

if (Get-Module -ListAvailable -Name MSOnline) {
  Write-Host "MSOnline Module exists"
} 
else {
  throw "MSOnline Module does not exist. You can install it by using 'Install-Module -Name MSOnline'"
}

if (Get-Module -ListAvailable -Name MicrosoftTeams) {
  Write-Host "MicrosoftTeams Module exists"
} 
else {
  throw "MicrosoftTeams Module does not exist. You can install it by using 'Install-Module -Name MicrosoftTeams'"
}

#----------------------------------------------------------[Functions]----------------------------------------------------------

function Add-MissingTeamMembers {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $true, HelpMessage = "This is the Team to add the users to")]
    [Microsoft.TeamsCmdlets.PowerShell.Custom.Model.TeamSettings] $team,
    [Parameter(Mandatory = $true, HelpMessage = "This is the AD Group membership from which to add missing members")]
    [Microsoft.Online.Administration.GroupMember[]] $ADGroupMembers
  )
  $TeamMembersAdded = [System.Collections.ArrayList]@()
  #Add missing members
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Checking membership of Team: " + $team.DisplayName + " ( " + $team.GroupId + " ) ") -ForegroundColor Yellow
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Against AD Group: " + $ADGroup.DisplayName + " ( " + $ADGroup.ObjectId + " ) ") -ForegroundColor Yellow
  Write-Host ("--------------------------------------------------------")
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Team Membership Total: " + $ExistingTeamMembers.count) -ForegroundColor Yellow
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " AD Group Membership Total: " + $ADGroupMembers.count) -ForegroundColor Yellow
  foreach ($groupMember in $ADGroupMembers) {
    #Check if exists in Teams already
    if ((($ExistingTeamMembers | Select-Object User) -Match $groupMember.EmailAddress).Count -eq 0 ) {
      #Add missing member
      Add-TeamUser -GroupId $team.GroupId -User $groupMember.EmailAddress
      Write-Host ("+ Added: " + $groupMember.EmailAddress)
      $TeamMembersAdded.Add($groupMember)
    }
    else {
      Write-Verbose ("| Existed: " + $groupMember.EmailAddress)
    }
  }
  Write-Host ("=====================")
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " " + $TeamMembersAdded.count + " new members added") -ForegroundColor Yellow
  Write-Host ("")
}

function Remove-MissingADUsers {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $true, HelpMessage = "This is the list of existing Team users you want to search. The full type should be Microsoft.TeamsCmdlets.PowerShell.Custom.GetTeamUser+GetTeamUserResponse")]
    [object] $ExistingTeamMembers,
    [Parameter(Mandatory = $true, HelpMessage = "This is the AD Group membership from which to check against for invalid team members")]
    [Microsoft.Online.Administration.GroupMember[]] $ADGroupMembers
  )
  $TeamMembersRemoved = [System.Collections.ArrayList]@()
  # Now check for existing Team members that are no longer AD Group members
  foreach ($teamMember in $ExistingTeamMembers) {
    #Check if exists in Teams already
    if (((($ADGroupMembers | Select-Object EmailAddress) -Match $teamMember.User).Count -eq 0) -and ($teamMember.Role -notmatch "owner") ) {
      #Remove from team
      Remove-TeamUser -GroupId $team.GroupId -User $teamMember.User
      $TeamMembersRemoved.Add($teamMember)
      Write-Verbose (" - Removed: " + $teamMember.User)
    }
    else {
      Write-Verbose (" | Not removed: " + $teamMember.User)
    }
  }
  Write-Host ("---------------------")
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " " + $TeamMembersRemoved.Count + " Team members removed") -ForegroundColor Yellow
}


function Get-AzureADGroupId {
  [CmdletBinding()]
  [OutputType([Guid])]
  Param(
    [Parameter(Mandatory = $true, HelpMessage = "The name of the Azure AD Group", ValueFromPipeline)]
    [String] $AzureADGroupName,
    [Parameter(Mandatory = $true, HelpMessage = "The AD Domain.")]
    [String] $ADDomain
  )

  process {
    $GroupUPN = $AzureADGroupName + "@" + $ADDomain

    $ADGroup = Get-MsolGroup -SearchString $GroupUPN
    $numMatched = $ADGroup.Count
    if ($numMatched -gt 1) {
      throw "!! ERROR !! Ambiguous Group Name. More than one group matched the search $GroupUPN"
    }
    else {
      $ADGroupId = $ADGroup[0].ObjectId
    }
    Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " AAD Group " + $GroupUPN + " has ObjectId " + $ADGroupId ) -ForegroundColor Blue
    return $ADGroupId
  }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Get Team

try {
  if ($TeamDisplayName.Length -gt 0) {
  
    Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Getting the Team..." + $TeamDisplayName) -ForegroundColor Blue
    $team = Get-Team -DisplayName $TeamDisplayName
    $numMatched = $team.Count
    if ($numMatched -gt 1) {
      throw "!! ERROR !! Ambiguous Team Display Name. More than one group matched the filter $TeamDisplayName . You will need to use TeamGroupId instead. You can get this from the Team URL Link address"
    }
    else {
      $team = $team[0]
    }
  }
  else {
    Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Getting the Team..." + $TeamGroupId) -ForegroundColor Blue
    $team = Get-Team -GroupId $TeamGroupId
  }

  # Get the Object ID from Group

  $ADGroupId = Get-AzureADGroupId -AzureADGroupName $AzureADGroupName -ADDomain $ADDomain
  # Get AD / Outlook Group Members
  $ADGroup = Get-MsolGroup -ObjectId $ADGroupId
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Getting the AD Group..." + $ADGroup.DisplayName) -ForegroundColor Blue
  if ($MaxUsers -gt 0) {
    $ADGroupMembers = Get-MsolGroupMember -GroupObjectId $ADGroupId -MaxResults $MaxUsers | Where-Object { $_.GroupMemberType -eq 'User' }
    $ADNestedGroups = Get-MsolGroupMember -GroupObjectId $ADGroupId -MaxResults $MaxUsers | Where-Object { $_.GroupMemberType -ne 'User' }
  }
  else {
    $ADGroupMembers = Get-MsolGroupMember -GroupObjectId $ADGroupId -All | Where-Object { $_.GroupMemberType -eq 'User' }
    $ADNestedGroups = Get-MsolGroupMember -GroupObjectId $ADGroupId -All | Where-Object { $_.GroupMemberType -ne 'User' }
  }
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " " + $ADGroupMembers.Count + " ...AD Group Members fetched (excluidng nested Groups)" ) -ForegroundColor Blue
  #Get existing Team members
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Getting the latest Team members..." + $TeamDisplayName) -ForegroundColor Blue
  $teamGroupId = $team.GroupId
  $ExistingTeamMembers = Get-TeamUser -GroupId $teamGroupId
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " " + $ADGroupMembers.Count + " ...Team Members fetched") -ForegroundColor Blue
  Add-MissingTeamMembers -team $team -ADGroupMembers $ADGroupMembers
  if (($RemoveInvalidTeamMembers) -and ($MaxUsers -eq 0)) {
    Remove-MissingADUsers -ExistingTeamMembers $ExistingTeamMembers -ADGroupMembers $ADGroupMembers
  }

  Write-Host "--------- Nested Groups Not Added ------------"
  foreach ($nestedgroupMember in $ADNestedGroups) {
    Write-Host (" - " + $nestedgroupMember.EmailAddress + ", " + $nestedgroupMember.DisplayName + ", " + $nestedgroupMember.GroupMemberType)
  }
  Write-Host "--------- END Nested Groups Not Added ------------"

  Write-Host ("=====================")
  Write-Host ("****** Completed ******") -ForegroundColor Blue
}
catch {
  throw "$_"
}