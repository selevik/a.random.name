<#
.Synopsis
   A function for syncing a AAD-group members with a Microsoft Team
.DESCRIPTION
   The default behavior is to syncronize (add and remove) members from the AAD-group to the Team group. Note that this is a function and need to be imported first (.\Sync-AADGroupMembersWithTeams.ps1)
.EXAMPLE
   Sync-AADGroupMembersWithTeams -TeamName 'A Random Name' -AADGroupName 'aRandomName' 
.EXAMPLE
   Sync-AADGroupMembersWithTeams -TeamName 'A Random Name' -AADGroupName 'aRandomName' -OnlyAdd -Cred $Credentials
.NOTES
   To use this funcetion script make sure to have both the Teams Module and MSOnline Module installed.
#>
function Sync-AADGroupMembersWithTeam
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'https://a.random.name')]
    Param
    (
        # The name of the Team you want to sync members to
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $TeamName,

        # The name of the group you want to sync members from
        [Parameter(Mandatory=$true, 
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true, 
            ValueFromRemainingArguments=$false, 
            Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $AADGroupName,
        
        # Provide a PSCredential-object to the function (for automation etc).
        [PSCredential]
        $Cred = (Get-Credential -Message "Provide Admin-UPN for MsOnline"),

        # Use -OnlyAdd if you only want to add new members and not remove any.
        [switch]
        $OnlyAdd,
        
        # Use -TestRun to only test what would happend (with only Write-host output). 
        [switch]
        $TestRun
    )

    Begin
    {
        function Compare-Arrays ($Reference, $Difference)
        {
            $returnObject = New-Object PSObject
            [System.Collections.Generic.HashSet[String]]$FirstHashSet = $Reference
            [System.Collections.Generic.HashSet[String]]$SecondHashSet = $Difference
            if ([string]::IsNullOrEmpty($FirstHashSet)) 
            {
                Add-Member -InputObject $returnObject -MemberType NoteProperty -Name AddToReference -Value $SecondHashSet
            }
            else 
            {
                $FirstHashSet.ExceptWith($SecondHashSet)
                Add-Member -InputObject $returnObject -MemberType NoteProperty -Name AddToReference -Value $FirstHashSet
            }

            [System.Collections.Generic.HashSet[String]]$FirstHashSet = $Reference
            [System.Collections.Generic.HashSet[String]]$SecondHashSet = $Difference
            if ([string]::IsNullOrEmpty($SecondHashSet)) 
            {
                 Add-Member -InputObject $returnObject -MemberType NoteProperty -Name RemoveFromReference -Value $FirstHashSet
            }
            else 
            {
                $SecondHashSet.ExceptWith($FirstHashSet)
                Add-Member -InputObject $returnObject -MemberType NoteProperty -Name RemoveFromReference -Value $SecondHashSet
            }
            return $returnObject
        }

        try 
        {
            Connect-MsolService -Credential $Cred
            Connect-MicrosoftTeams -Credential $Cred | Out-Null           
        }
        catch
        {
            throw "Something went wrong when connecting to Online. Wrong Credentials? Missing Module? Error:`n$_"
        }      
    }
    Process
    {
        #TeamGroup
        $TeamGroups= Get-Msolgroup -SearchString $TeamName -GroupType DistributionList
        if (!$TeamGroups)
        {
            throw "Could not find $TeamName in Azure! Error:`n$_" 
        }
        if ($TeamGroups.Count -gt 1) #More than 1 group found, trying to exact match
        {
            $TeamGroupID =  $TeamGroups | ? DisplayName -eq $TeamName | Select ObjectID
            if (!$TeamGroupID)
            {
                throw "More than one Azure Group fond for $TeamName and could not get an exact match"
            }
        }
        else
        {
            $TeamGroupID =  $TeamGroups.ObjectId
        }

        #AAD-Group
        $AADGroups= Get-Msolgroup -SearchString $AADGroupName #| ? {$_.GroupType -ne "DistributionList"}
        if (!$AADGroups)
        {
            throw "Could not find $AADGroupName in Azure! Error:`n$_" 
        }
        if ($AADGroups.Count -gt 1) #More than 1 group found, trying to exact match
        {
            $AADGroupID =  $AADGroups | ? DisplayName -eq $AADGroupName | Select ObjectID
            if (!$AADGroupID)
            {
                throw "More than one Azure Group fond for $AADGroupName and could not get an exact match"
            }
        }
        else
        {
            $AADGroupID =  $AADGroups.ObjectId
        }
        
        #Members
        $AADGroupMembers = Get-MsolGroupMember -GroupObjectId $AADGroupID
        $TeamGroupMembers = Get-MsolGroupMember -GroupObjectId $TeamGroupID

        $Comparison = Compare-Arrays -Reference $AADGroupMembers.ObjectId -Difference $TeamGroupMembers.ObjectId
        if ($Comparison.AddToReference.count -gt 0)
        {
            foreach ($objectId in $Comparison.AddToReference)
            {
                try
                {
                    $upn = Get-MsolUser -ObjectId $objectId | select -ExpandProperty userPrincipalName 
                    if ($TestRun)
                    {
                        Write-Host "Would have: Add-TeamUser -GroupId $TeamGroupID -User $upn -Role Member" -ForegroundColor Gray 
                    }
                    else
                    {
                        Add-TeamUser -GroupId $TeamGroupID -User $upn -Role Member
                    }
                    #implement logging here
                    Write-Host "Sucessfully added $upn to $TeamName"
                }
                catch
                {
                    throw "Could not add user $upn to the team $TeamName, error: $_"
                }
            }
        
            if (!$OnlyAdd) #then perform removes aswell. Bare in mind that you cannot remove the last administrator. 
            {
                foreach ($objectId in $Comparison.RemoveFromReference)
                {
                    try
                    {
                        $upn = Get-MsolUser -ObjectId $objectId | select -ExpandProperty userPrincipalName 
                        if ($TestRun)
                        {
                            Write-Host "Would have: Remove-TeamUser -GroupId $TeamGroupID -User $upn" -ForegroundColor Gray 
                        }
                        else
                        {
                            Remove-TeamUser -GroupId $TeamGroupID -User $upn
                        }
                        #implement logging here
                        Write-Host "Sucessfully removed $upn from $TeamName"
                    }
                    catch
                    {
                        throw "Could not remove user $upn from team $TeamName, error: $_"
                    }
                }
            }   
        }
        else
        {
            Write-host "Nothing to do, the Group and Team is already up to date and synced"
        }    
    }
    End
    {
        #Some nice ending ;)
    }
}
