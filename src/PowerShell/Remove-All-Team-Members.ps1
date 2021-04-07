<#PSScriptInfo
.VERSION 1.1
.GUID ba4dc84c-90d9-43ed-829a-4701161c152d
.AUTHOR Nicholas Rogoff

.RELEASENOTES
Initial version.
#>
<# 
.SYNOPSIS 
Removes all Team member from a Team except the Owners. 
 
.DESCRIPTION 
Loops through all members of a Team and removes them...except Owners.

PRE-REQUIREMENT
---------------
Install Teams Module (MS Online)
PS> Install-Module -Name MSOnline

Install Microsoft Teams cmdlets module
PS> Install-Module -Name MicrosoftTeams

.INPUTS
  None. You cannot pipe objects to this!

.OUTPUTS
  None.

.PARAMETER TeamDisplayName
The display name of the Team who's membership you wish to alter

.PARAMETER Credential
The credentials (PSCredential object) for an Owner of the Team. Use '$Credential = Get-Credential' to prompt and store the credentials to securely pass

.NOTES
  Version:        1.0
  Author:         Nicholas Rogoff
  Creation Date:  2020-05-15
  Purpose/Change: Initial script development 

  .EXAMPLE 
PS> .\SyncTeamMembership.ps1 -TeamDisplayName "My Team"
PS> .\SyncTeamMembership.ps1 -TeamDisplayName "My Team" -Verbose
#>
#---------------------------------------------------------[Script Parameters]------------------------------------------------------

[CmdletBinding()]
Param(
  [Parameter(Mandatory = $true, HelpMessage = "The display name of the Team")]
  [string] $TeamDisplayName,
  [Parameter(Mandatory=$true, HelpMessage="The credentials for an Owner of the Team")]
  [System.Management.Automation.PSCredential] $Credential
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Starting Team Members removal...") -ForegroundColor Blue

# Set Error Action to Silently Continue
$ErrorActionPreference = 'Continue'

# Signin to Office 365 stuff
Connect-MicrosoftTeams -Credential $Credential
Connect-MsolService -Credential $Credential

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

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Remove-Members {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $true, HelpMessage = "The display name of the Team")]
    [string] $TeamDisplayName,
    [Parameter(Mandatory=$true, HelpMessage="The credentials for an Owner of the Team")]
    [System.Management.Automation.PSCredential] $Credential
  )

  # Get Team
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Getting the Team..." + $TeamDisplayName) -ForegroundColor Blue
  $team = Get-Team -DisplayName $TeamDisplayName

  # Get existing Team members
  $ExistingTeamMembers = Get-TeamUser -GroupId $team.GroupId
  $TeamMembersRemoved = [System.Collections.ArrayList]@()

  # Add missing members
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Checking membership of Team: " + $team.DisplayName + " ( " + $team.GroupId + " ) ") -ForegroundColor Yellow
  Write-Host ("--------------------------------------------------------")
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " Team Membership Total: " + $ExistingTeamMembers.count) -ForegroundColor Yellow
  try {
    foreach ($teamMember in $ExistingTeamMembers) {
      # Check if exists in Teams already
      if ($teamMember.Role -notmatch "owner" ) {
        #Remove from team
        Remove-TeamUser -GroupId $team.GroupId -User $teamMember.User
        $TeamMembersRemoved.Add($teamMember)
        Write-Verbose ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " - Removed: " + $teamMember.User)
      }
      else {
        Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " | Ownner NOT removed: " + $teamMember.User) -ForegroundColor Magenta
      }
    }
  }
  catch {
    Write-Host -BackgroundColor Red ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + "Error: $($_.Exception)") -ForegroundColor Red
    Break
  }
  Write-Host ("---------------------")
  Write-Host ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + " " + $TeamMembersRemoved.count + " Team members removed") -ForegroundColor Yellow
  Write-Host ("=====================")
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------


Remove-Members -TeamDisplayName $TeamDisplayName -Credential $Credential

Write-Host ("****** Completed ******") -ForegroundColor Blue
Write-Host ("=======================") -ForegroundColor Blue
