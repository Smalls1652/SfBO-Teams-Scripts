
<#PSScriptInfo

.VERSION 2020.04.02

.GUID a72501cc-9ebb-4f67-8583-b0a2c0ef6f25

.AUTHOR Tim Small

.COMPANYNAME Smalls.Online

.COPYRIGHT 2020

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<#
.SYNOPSIS
    Get numbers assigned to users in Teams.
.DESCRIPTION
    Get the subscriber numbers in the Skype for Business/Teams system and match them with the users they are assigned to.
.PARAMETER SearchFilter
    Apply a search filter for users.
.EXAMPLE
    PS C:\> $csOnline = New-CsOnlineSession -UserName "jdoe1@contoso.com"
    PS C:\> Import-PSSesion -Session $csOnline
    PS C:\> Get-TeamsAssignedNumbers.ps1

    Connect to the Skype for Business Online session, import the cmdlets to the session, and then get the number assignments.
.EXAMPLE
    PS C:\> $csOnline = New-CsOnlineSession -UserName "jdoe1@contoso.com"
    PS C:\> Import-PSSesion -Session $csOnline
    PS C:\> Get-TeamsAssignedNumbers.ps1 -SearchFilter { UserPrincipalName -like "*@sub.contoso.com" }

    Connect to the Skype for Business Online session, import the cmdlets to the session, and then get the number assignments for users that have a UserPrincipalName like '*@sub.contoso.com'.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$SearchFilter
)

begin {

    <#
    Before processing, run checks to ensure that:
        1. An active session is available to Skype for Business Online
        2. That the session has been imported into the current console session.
    #>

    #Check for an active SfBO session. If a session is not found in an open state, throw a terminating error.
    $sfbSessions = Get-PSSession | Where-Object { ($PSItem.Name -like "SfBPowerShellSession_*") -and ($PSItem.State -eq "Opened") }
    switch (($null -eq $sfbSessions)) {
        $false {
            Write-Verbose "An active SfB Online PowerShell session was found."
            break
        }

        Default {
            $PSCmdlet.ThrowTerminatingError(
                [System.Management.Automation.ErrorRecord]::new(
                    [System.Exception]::new("An active SfB PowerShell session was not found."),
                    "SfB.NoActiveSession",
                    [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                    $sfbSessions
                )
            )
            break
        }
    }

    #Check to ensure that the SfBO session has been imported. If the two used commands in this script aren't found, then throw a terminating error.
    try {
        $null = Get-Command -Name @("Get-CsOnlineTelephoneNumber", "Get-CsOnlineUser") -ErrorAction "Stop"
    }
    catch [System.Exception] {
        $ErrorDetails = $PSItem

        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.Exception]::new("Could not find cmdlets for SfB Online. Ensure that the session has been imported into your console."),
                "SfB.ModuleNotImported",
                [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                $ErrorDetails
            )
        )
    }

    #Splat variables for the main and sub progress bars. This is to simplify the cmdlets for 'Write-Progress'.
    $MainProgressSplat = @{
        "Activity" = "Get Assigned Phone Numbers";
        "Status"   = "Progress->";
        "Id"       = 0;
    }

    $SubProgressSplat = @{
        "Activity" = "Processing numbers";
        "Status"   = "Progress->";
        "Id"       = 1;
        "ParentId" = 0;
    }
}

process {
    #Get all of the 'Subscriber' type phone numbers.
    Write-Progress @MainProgressSplat -PercentComplete 0 -CurrentOperation "Getting phone numbers"
    $csAssignedNumbers = Get-CsOnlineTelephoneNumber -InventoryType Subscriber

    switch ([string]::IsNullOrEmpty($SearchFilter)) {
        #If '$SearchFilter' is not null or empty, then apply the search filter to get users.
        $false {
            Write-Progress @MainProgressSplat -PercentComplete 25 -CurrentOperation "Getting all users with filter '$($SearchFilter)'"
            $csUsers = Get-CsOnlineUser -Filter $SearchFilter -ErrorAction Stop
            break
        }

        #If '$SearchFilter' is null or empty, then get all users.
        Default {
            Write-Progress @MainProgressSplat -PercentComplete 25 -CurrentOperation "Getting all users"
            $csUsers = Get-CsOnlineUser
        }
    }
    
    $n = 1 #Init loop counter
    $TotalNumberOfNumbers = ($csAssignedNumbers | Measure-Object).Count #Get the total number of returned user objects.

    $returnObj = [System.Collections.Generic.List[pscustomobject]]::new() #Init the returnObj as a List<PSCustomObject> object.

    #Loop through each number and match to the correct user.
    Write-Progress @MainProgressSplat -PercentComplete 50 -CurrentOperation "Matching numbers to users"
    Write-Progress @SubProgressSplat
    foreach ($number in $csAssignedNumbers) {
        #Calculate the percentage complete for the sub-progress bar.
        $percentComplete = (($n / $TotalNumberOfNumbers) * 100)
        Write-Progress @SubProgressSplat -PercentComplete $percentComplete -CurrentOperation "Searching for user with '$($number.ID)'"

        #Filter for the user with the number.
        $csUser = $csUsers | Where-Object { $PSItem.LineURI -like "*$($number.Id)" }   
        
        switch (($null -eq $csUser)) {
            #If the number matched
            $false {
                $upn = $csUser.UserPrincipalName
                break
            }

            #If the number didn't match
            Default {
                $upn = "Not Assigned"
                break
            }
        }
        
        #Build the object and add it to the returnObj.
        $returnObj.Add(
            [pscustomobject]@{
                "UserPrincipalName" = $upn;
                "AssignedNumber"    = ($number.Id -replace "(\d)(\d{3})(\d{3})(\d{4})", "+`$1 (`$2) `$3-`$4")
            }
        )

        $n++ #Increment the loop counter
    }

    #Mark the progress bars as completed.
    Write-Progress @SubProgressSplat -Completed
    Write-Progress @MainProgressSplat -Completed
}

end {
    return $returnObj
}