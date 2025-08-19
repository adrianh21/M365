#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
Imports group memberships and ownership from a CSV into various Microsoft 365 groups.

.DESCRIPTION
Reads a CSV file where each row defines a group and its members/managers/owners. The script determines the target group's
type in Microsoft 365 (M365 Group, Distribution List, or Mail-Enabled Security Group) and adds users accordingly.
- For M365 Groups: Adds members and owners based on separate CSV columns.
- For DLs/MESGs: Adds all specified users (members, managers, owners) as standard members.

.PARAMETER CsvPath
The full path to the input CSV file.

.PARAMETER CsvHeaderGroupEmail
The CSV header for the group's email address.

.PARAMETER CsvHeaderManagers
The CSV header for the space-separated list of manager emails.

.PARAMETER CsvHeaderMembers
The CSV header for the space-separated list of member emails.

.PARAMETER CsvHeaderOwners
The CSV header for the space-separated list of owner emails.

.PARAMETER WhatIf
Shows what actions would be taken without making changes.

.PARAMETER Confirm
Prompts for confirmation before performing actions that modify groups.

.EXAMPLE
.\Import-M365GroupMemberships.ps1 -CsvPath "C:\temp\groups.csv" -Verbose

.EXAMPLE
# Run in WhatIf mode with custom headers
.\Import-M365GroupMemberships.ps1 -CsvPath "C:\data\groups.csv" -CsvHeaderGroupEmail "GroupAddress" -CsvHeaderManagers "MgrList" -WhatIf

.NOTES
- Assumes group and user email addresses are identical between the source (e.g., Google) and Microsoft 365.
- For DLs and MESGs, source 'Managers' and 'Owners' are added as MEMBERS. The 'ManagedBy' property is not modified.
- You must be connected to Exchange Online via `Connect-ExchangeOnline`.
- Requires appropriate permissions (e.g., Exchange Administrator, Groups Administrator).
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$CsvHeaderGroupEmail = "email",

    [Parameter(Mandatory = $false)]
    [string]$CsvHeaderManagers = "Managers",

    [Parameter(Mandatory = $false)]
    [string]$CsvHeaderMembers = "Members",

    [Parameter(Mandatory = $false)]
    [string]$CsvHeaderOwners = "Owners"
)

# --- Helper Functions ---

function Get-UpnsFromCsvColumn {
    param(
        [string]$EmailString
    )
    # This helper function parses a space-separated string of emails,
    # validates each one, and returns a unique, non-empty array of UPNs.
    $upnSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if (-not [string]::IsNullOrWhiteSpace($EmailString)) {
        foreach ($email in $EmailString.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)) {
            $trimmedEmail = $email.Trim()
            if ($trimmedEmail -like '*@*') {
                $upnSet.Add($trimmedEmail) | Out-Null
            }
            else {
                Write-Warning "Skipping invalid user email format: '$trimmedEmail'"
            }
        }
    }
    return [string[]]$upnSet
}

# --- Script Initialization ---

# Check connection
try {
    Get-ConnectionInformation | Out-Null
    Write-Verbose "Exchange Online connection verified."
}
catch {
    Write-Error "Not connected to Exchange Online. Please run Connect-ExchangeOnline and try again."
    return
}

# Import CSV
try {
    $importData = Import-Csv -Path $CsvPath
    if ($null -eq $importData) { throw }
    Write-Verbose "Successfully imported $($importData.Count) rows from '$CsvPath'."
}
catch {
    Write-Error "Failed to import or read CSV file at '$CsvPath'. Error: $($_.Exception.Message)"
    return
}

# Verify headers
$requiredHeaders = @($CsvHeaderGroupEmail, $CsvHeaderManagers, $CsvHeaderMembers, $CsvHeaderOwners)
$actualHeaders = $importData[0].PSObject.Properties.Name
$missingHeaders = $requiredHeaders | Where-Object { $actualHeaders -notcontains $_ }
if ($missingHeaders.Count -gt 0) {
    Write-Error "CSV is missing required header(s): '$($missingHeaders -join "', '")'. Please check the CSV file or the -CsvHeader* parameters."
    return
}
Write-Verbose "CSV headers verified."

# --- Main Processing Loop ---

$totalRows = $importData.Count
$processedCount = 0
$summary = @{
    TotalRows = $totalRows
    GroupsProcessed = 0
    GroupsSkipped = 0
    MembershipErrors = 0
}

Write-Host "Starting group membership processing for $totalRows groups."

foreach ($row in $importData) {
    $processedCount++
    $googleGroupEmail = $row.$($CsvHeaderGroupEmail).Trim()

    Write-Progress -Activity "Processing Groups from CSV" -Status "Group $processedCount of ${totalRows}: $googleGroupEmail" -PercentComplete (($processedCount / $totalRows) * 100)

    Write-Host ("-" * 50)
    Write-Host "Processing Row ${processedCount}: $googleGroupEmail"

    if ([string]::IsNullOrWhiteSpace($googleGroupEmail)) {
        Write-Warning "Skipping row $processedCount because the group email is empty."
        $summary.GroupsSkipped++
        continue
    }

    # 1. Identify Target Recipient
    try {
        $targetRecipient = Get-Recipient -Identity $googleGroupEmail -ErrorAction Stop
        Write-Verbose "Found target '$googleGroupEmail' with type '$($targetRecipient.RecipientTypeDetails)'."
    }
    catch {
        Write-Warning "Could not find or access '$googleGroupEmail' in Microsoft 365. Skipping. Error: $($_.Exception.Message)"
        $summary.GroupsSkipped++
        continue
    }

    # 2. Prepare Member and Owner UPN lists using the helper function
    $ownersToAdd = Get-UpnsFromCsvColumn -EmailString ($row.$($CsvHeaderManagers) + " " + $row.$($CsvHeaderOwners))
    $membersToAdd = Get-UpnsFromCsvColumn -EmailString $row.$($CsvHeaderMembers)
    Write-Verbose "Found $($ownersToAdd.Count) potential owners and $($membersToAdd.Count) potential members in CSV row."

    $summary.GroupsProcessed++

    # 3. Process based on Target Recipient Type
    switch -Wildcard ($targetRecipient.RecipientTypeDetails) {
        # --- Microsoft 365 Group ---
        "GroupMailbox" {
            Write-Host "Target is a Microsoft 365 Group. Processing owners and members."

            # Optimize: Remove any users from the member list who are already designated as owners.
            $ownerSet = [System.Collections.Generic.HashSet[string]]::new($ownersToAdd, [System.StringComparer]::OrdinalIgnoreCase)
            $membersOnly = $membersToAdd | Where-Object { -not $ownerSet.Contains($_) }
            if ($membersToAdd.Count -ne $membersOnly.Count) {
                Write-Verbose "Optimization: $($membersToAdd.Count - $membersOnly.Count) user(s) removed from member list as they are already being added as owners."
            }

            # Add Owners
            if ($ownersToAdd.Count -gt 0) {
                if ($PSCmdlet.ShouldProcess($googleGroupEmail, "Add $($ownersToAdd.Count) Owner(s)")) {
                    try {
                        Add-UnifiedGroupLinks -Identity $googleGroupEmail -LinkType Owners -Links $ownersToAdd -ErrorAction Stop
                        Write-Host "Successfully added $($ownersToAdd.Count) owner(s)."
                    }
                    catch {
                        Write-Warning "Failed to add owners to '$googleGroupEmail'. Error: $($_.Exception.Message)"
                        $summary.MembershipErrors++
                    }
                }
            }

            # Add Members
            if ($membersOnly.Count -gt 0) {
                if ($PSCmdlet.ShouldProcess($googleGroupEmail, "Add $($membersOnly.Count) Member(s)")) {
                    try {
                        Add-UnifiedGroupLinks -Identity $googleGroupEmail -LinkType Members -Links $membersOnly -ErrorAction Stop
                        Write-Host "Successfully added $($membersOnly.Count) member(s)."
                    }
                    catch {
                        Write-Warning "Failed to add members to '$googleGroupEmail'. Error: $($_.Exception.Message)"
                        $summary.MembershipErrors++
                    }
                }
            }
        }

        # --- Distribution List or Mail-Enabled Security Group ---
        "*DistributionGroup" {
            Write-Warning "Target '$googleGroupEmail' is a DL or MESG. All users from CSV (Owners, Managers, Members) will be added as Members."

            # Combine all users into a single, unique list.
            $allUsersSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $ownersToAdd | ForEach-Object { $allUsersSet.Add($_) | Out-Null }
            $membersToAdd | ForEach-Object { $allUsersSet.Add($_) | Out-Null }
            $allMembersToAdd = [string[]]$allUsersSet

            if ($allMembersToAdd.Count -eq 0) {
                Write-Verbose "No new members to add."
                continue
            }
            
            Write-Host "Attempting to add $($allMembersToAdd.Count) member(s) to '$googleGroupEmail'."
            if ($PSCmdlet.ShouldProcess($googleGroupEmail, "Add $($allMembersToAdd.Count) Member(s) individually")) {
                foreach ($member in $allMembersToAdd) {
                    try {
                        Add-DistributionGroupMember -Identity $googleGroupEmail -Member $member -BypassSecurityGroupManagerCheck -ErrorAction Stop
                        Write-Verbose " -> Successfully added member: $member"
                    }
                    catch {
                        Write-Warning " -> Failed to add member '$member'. Error: $($_.Exception.Message)"
                        $summary.MembershipErrors++
                    }
                }
            }
        }

        # --- Default Case for unhandled types ---
        default {
            Write-Warning "Target '$googleGroupEmail' is a '$($targetRecipient.RecipientTypeDetails)', which is not supported by this script. Skipping."
            $summary.GroupsSkipped++
            $summary.GroupsProcessed-- # Decrement since we didn't actually process it
        }
    } # End Switch
} # End foreach row

# --- Final Summary ---
Write-Progress -Activity "Processing Groups from CSV" -Completed
Write-Host ("=" * 50)
Write-Host "Script Finished. Summary:"
Write-Host " - Total Rows in CSV:      $($summary.TotalRows)"
Write-Host " - Groups Processed:         $($summary.GroupsProcessed)"
Write-Host " - Groups Skipped:           $($summary.GroupsSkipped) (Not found, empty email, or unsupported type)"
Write-Host " - Membership Add Failures:  $($summary.MembershipErrors) (See warnings above for details)"
Write-Host ("=" * 50)