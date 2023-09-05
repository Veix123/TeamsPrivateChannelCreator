# Copyright [2020] [Veiko Vunder]
# 
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# Credits: https://chrishayward.co.uk/2020/08/24/microsoft-teams-installing-powershell-preview-beta-modules-manage-private-channels/


#########################################################
## This script works in administrative privileges only!
##
## Run the following commented commands to get the preview version of the MicrosoftTeams module with private channel support!
##
## Update the PowerShellGet module
#Install-Module -Name PowerShellGet -Repository PSGallery -Force
#
## RESTART POWERSHELL!!!
#
## Remove any old Teams modules
#Uninstall-Module MicrosoftTeams
#
## Install new MicrosoftTeams preview module
#Install-Module -Name MicrosoftTeams -RequiredVersion 3.1.1 -AllowPrerelease
#
## Some additional commands for debugging
#Get-Module -Name MicrosoftTeams
#Find-Module -Name MicrosoftTeams
#Get-Command -Module MicrosoftTeams
#Get-Help New-TeamChannel


# Creates a new Team and/or gets an id of the existing group
function CreateNewGroup($GroupShortName, $GroupDisplayName, $Owner, [REF]$GroupIdRef)
{
    $Group=$null
    $GroupIdRef.Value=$null
    try
    {
        $Group = New-Team -MailNickname "$GroupShortName" -displayname "$GroupDisplayName" -Visibility "private" -Owner $Owner
        $GroupIdRef.Value = $Group.GroupId
    }
    catch
    {
        # Failed to create new group, try if it already exists
        try
        {
            $Group = Get-Team -MailNickName $GroupShortName
            if($Group)
            {
                $UserResponse= [System.Windows.Forms.MessageBox]::Show("Group $($Group.MailNickName) already exists!!!`nAdd private channels to the existing group?" , "Attention" , [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
                if($UserResponse -eq "OK")
                {
                    $GroupIdRef.Value = $Group.GroupId
                    "Using existing group: '$GroupShortName' with display name '$GroupDisplayName'"
                }
                else
                {
                    return
                }
            }
            else
            {
                return
            }
        }
        catch
        {
            "Failed to create and retreive group"
            return
        }
    }
    "Created new group: '$GroupShortName' with display name '$GroupDisplayName'"
}

function AddStudentToGroup($GroupId, $StudentName, $StudentEmail)
{
    try
    {
        Add-TeamUser -GroupId "$GroupId" -User $StudentEmail
        $StudentTeamsEmail = (Get-TeamUser -GroupId $GroupId | where Name -Like $StudentName).User
        if($StudentTeamsEmail)
        {
            "User '$StudentName' with email '$StudentEmail' added to the group. Teams email is $StudentTeamsEmail"
        }
        else
        {
            Write-Warning "Student '$StudentName' with email '$StudentEmail' was added to the group, but the name does not match with the ones in MSTeam"
        }
    }
    catch
    {
        Write-Warning "Failed to add student '$StudentName' with email '$StudentEmail' to the group"
    }
}

function AddAdminsToGroup($GroupId, [array]$Admins)
{
    foreach ($Admin in ($Admins | select -skip 1))
    {
        try
        {
            Add-TeamUser -GroupId "$GroupId" -User $Admin
            Add-TeamUser -GroupId "$GroupId" -User $Admin -Role Owner
            "Admin '$Admin' added to the group"
        }
        catch
        {
            "ERROR: Failed to add admin '$Admin' to the group"
        }
    }
}


function CreatePrivateChannelForStudent($GroupId, $StudentName, $ChannelPrefix, [array]$Admins)
{
    # Check if student is present in the team
    $StudentTeamsEmail = (Get-TeamUser -GroupId $GroupId | where Name -Like $StudentName).User
    if(!$StudentTeamsEmail)
    {
        "ERROR: Student '$StudentName' not found from the team!"
        return
    }
    
    # Create the channel
    $StudentNameTrimmed = $StudentName.Trim("'")
    $ChannelName = "$ChannelPrefix$StudentNameTrimmed"
    "Creating private channel '$ChannelName'"
    try
    {
        New-TeamChannel -GroupId $GroupId -DisplayName "$ChannelName" -Description "Private channel for communicating with supervisors" -MembershipType Private -Owner $Admins[0]
    }
    catch
    {
        Write-Warning "ERROR: Failed to create a private channel '$ChannelName' (already exists?)"
    }
}

function ChangePrivateChannelName($Group, $StudentName, $ChannelPrefix)
{
    # Check if student is present in the team
    $StudentTeamsEmail = (Get-TeamUser -GroupId $GroupId | where Name -Like $StudentName).User
    if(!$StudentTeamsEmail)
    {
        "ERROR: Student '$StudentName' not found from the team!"
        return
    }

    $CurrentChannelName = "$ChannelPrefix$StudentName"
    $NewChannelName = "WED 10 $StudentName"
    #try
    #{
        Set-TeamChannel -GroupId $GroupId -CurrentDisplayName "$CurrentChannelName" -NewDisplayName "$NewChannelName"
    #}
   # catch
    #{
    #    Write-Warning "ERROR: Failed to change a private channel '$ChannelName' name"
   # }
}

function AddStudentAndAdminsToPrivateChannel($GroupId, $StudentName, $ChannelPrefix, [array]$Admins)
{
    # Check if student is present in the team
    $StudentTeamsEmail = (Get-TeamUser -GroupId $GroupId | where Name -Like $StudentName).User
    if(!$StudentTeamsEmail)
    {
        "ERROR: Student '$StudentName' not found from the team!"
        return
    }

    $StudentNameTrimmed = $StudentName.Trim("'")
    $ChannelName = "$ChannelPrefix$StudentNameTrimmed"

    # Add the student to the channel
    try
    {
        Add-TeamChannelUser -GroupId $GroupId -DisplayName "$ChannelName" -User $StudentTeamsEmail
    }
    catch
    {
        Write-Warning "ERROR: Failed to add student to the private channel"
    }
    
    # Add the rest of the admins
    foreach ($Admin in ($Admins | select -skip 1))
    {
        try
        {
            Add-TeamChannelUser -GroupId $GroupId -DisplayName "$ChannelName" -User $Admin
            Add-TeamChannelUser -GroupId $GroupId -DisplayName "$ChannelName" -User $Admin -Role Owner
            "Admin '$Admin' added to the private channel '$ChannelName'"
        }
        catch
        {
            Write-Warning "Failed to add admins to the private channel '$ChannelName'"
        }
    } 
}


########################## ENTER TEAM DETAILS HERE #####################################

# Graphical programming course team 1
#$GroupDisplayName = "Visuaalprogrammeerimine 22/23K (LOTI.05.086)"
#$GroupShortName = "VisProg23K"

# Graphical programming course team 2
#$GroupDisplayName = "Visuaalprogrammeerimine 22/23K (LOTI.05.086) N14"
#$GroupShortName = "VisProg23KLisa"

# Data acquisition course
$GroupDisplayName = "2023 DAQSP (LOTI.05.052)"
$GroupShortName = "DAQSP2023Fall"


# Global variable that will hold an ID of the team
$global:GroupId = $null
#$global:GroupId = "3a9165a2-731f-44a9-9507-20265b75b7af"

# Let's go
Connect-MicrosoftTeams

# Read the admins Teams email addresses from a text file, each address in a separate line.
$Admins = [array](Get-Content -Path "C:\Users\Veiko\Desktop\DAQSP\daqadmins.csv")
$Admins | ft

CreateNewGroup -GroupShortName "$GroupShortName" -GroupDisplayName "$GroupDisplayName" -Owner $Admins[0] -GroupIdRef ([REF]$global:GroupId)
"Using GroupId: " + $global:GroupId

AddAdminsToGroup -GroupId "$global:GroupId" -Admins $Admins

# Read the registered students from a CSV file.
# Use semicolon as a separator amd the following format with a header row (FirstName; LastName; Group; Student Email)
$Students = [array](Import-Csv -Delimiter ";" -Path "C:\Users\Veiko\Desktop\DAQSP\daqstudents.csv") 
"Students total: " + $Students.Count
$Students | ft

# Get a subset of students to fit into max 30 private group limit.
#$students = $students |  Where-Object {$_.Group -eq "K12" -or $_.Group -eq "N12"}
#$Students = $Students |  Where-Object {$_.FirstName -eq "X" -or $_.LastName -eq "Y" -or $_.LastName -eq "Z"}
#"Students subset: " + $Students.Count
#$Students | ft


# Add students to the team
foreach($line in $Students)
{
    if($line)
    {
        AddStudentToGroup -GroupId "$global:GroupId" -StudentName "$($line.FirstName) $($line.LastName)" -StudentEmail "$($line.Email)"
    }
}

# Prepare the private channel (first admin will be set as owner at this point)
foreach($line in $Students)
{
    if($line)
    {
        CreatePrivateChannelForStudent -GroupId "$global:GroupId" -StudentName "$($line.FirstName) $($line.LastName)" -ChannelPrefix "$($line.Group) " -Admins $Admins
#        CreatePrivateChannelForStudent -GroupId "$global:GroupId" -StudentName "$($line.FirstName) $($line.LastName)" -Admins $Admins
    }
}

# Add student and the rest of the admins to the private channel
foreach($line in $Students)
{
    if($line)
    {
        AddStudentAndAdminsToPrivateChannel -GroupId "$global:GroupId" -StudentName "$($line.FirstName) $($line.LastName)" -ChannelPrefix "$($line.Group) " -Admins $Admins
#        AddStudentAndAdminsToPrivateChannel -GroupId "$global:GroupId" -StudentName "$($line.FirstName) $($line.LastName)" -Admins $Admins
    }
}


foreach($line in $Students)
{

    if($line)
    {
        ChangePrivateChannelName -GroupId "$global:GroupId" -StudentName "$($line.FirstName) $($line.LastName)" -ChannelPrefix "$($line.Group) "
    }
}

# The separated foreach blocks above is necessary because it takes some time for a private channel to become available. Trying to add members to the channel that was just created will fail!