#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
Imports group memberships and ownership into various Microsoft 365 group types (M365 Groups, DLs, MESGs)
from a CSV where each row defines a group and its members/managers/owners in specific columns.

.DESCRIPTION
Reads a CSV file (list format: one row per group). Determines the target group type in M365 (M365 Group, DL, MESG).
- For M365 Groups: Adds Google Members as Members, Google Managers/Owners as Owners.
- For DLs/MESGs: Adds ALL Google Members/Managers/Owners as Members. (See NOTES).
Handles identity/group mapping via customizable functions (simplified based on prior user input).

.PARAMETER CsvPath
The full path to the input CSV file (list format).

.PARAMETER CsvHeaderGroupEmail
The header name in the CSV for the column containing the group's email address. Defaults to "email".

.PARAMETER CsvHeaderManagers
The header name in the CSV for the column containing the space-separated list of manager emails. Defaults to "Managers".

.PARAMETER CsvHeaderMembers
The header name in the CSV for the column containing the space-separated list of member emails. Defaults to "Members".

.PARAMETER CsvHeaderOwners
The header name in the CSV for the column containing the space-separated list of owner emails. Defaults to "Owners".

.PARAMETER WhatIf
Shows what actions would be taken without actually making changes.

.PARAMETER Confirm
Prompts for confirmation before performing actions that modify groups.

.EXAMPLE
# Run against a CSV, handling mixed group types, using default headers
.\Import-GoogleGroupMembershipLists-MultiType.ps1 -CsvPath "C:\temp\groups_list_format.csv" -Verbose

.EXAMPLE
# Run in WhatIf mode with custom headers
.\Import-GoogleGroupMembershipLists-MultiType.ps1 -CsvPath "C:\Users\Adrian\Downloads\groups.csv" -CsvHeaderGroupEmail "GroupAddress" -CsvHeaderManagers "MgrList" -WhatIf

.NOTES
- Assumes Group emails and User emails/UPNs are IDENTICAL between Google and M365, based on prior user confirmation. Functions reflect this.
- For Distribution Lists (DLs) and Mail-Enabled Security Groups (MESGs), Google Managers/Owners are added as MEMBERS, not owners/managers in M365, as the concepts differ. The 'ManagedBy' property is NOT modified by this script.
- Ensure you are connected to Exchange Online (`Connect-ExchangeOnline`).
- Requires appropriate permissions (e.g., Global Admin, Exchange Admin, Groups Admin) in Microsoft 365.
- Target Microsoft 365 groups/lists must exist beforehand.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    # Note: $GoogleDomain / $M365Domain are removed as parameters as the conversion functions
    # below are simplified based on prior confirmation of identical emails/UPNs.
    # If that assumption changes, re-add parameters and revert function logic.

    [Parameter(Mandatory = $false)]
    [string]$CsvHeaderGroupEmail = "email",

    [Parameter(Mandatory = $false)]
    [string]$CsvHeaderManagers = "Managers",

    [Parameter(Mandatory = $false)]
    [string]$CsvHeaderMembers = "Members",

    [Parameter(Mandatory = $false)]
    [string]$CsvHeaderOwners = "Owners"
)

# --- Function Definitions (Simplified based on prior user confirmation) ---

function Convert-GoogleEmailToM365UPN {
    param(
        [string]$GoogleEmail
    )
    # Simplified: Assumes Google User Email == M365 UPN
    if ($GoogleEmail -like '*@*') {
        return $GoogleEmail.Trim()
    } else {
        Write-Warning "Input '$GoogleEmail' provided for user conversion does not appear to be a valid email format."
        return $null
    }
}

function Get-TargetM365GroupEmail {
    param(
        [string]$GoogleGroupEmail
    )
    # Simplified: Assumes Google Group Email == M365 Group Email
     if ($GoogleGroupEmail -like '*@*') {
        return $GoogleGroupEmail.Trim()
    } else {
        Write-Warning "Input '$GoogleGroupEmail' does not look like a valid email address for group identification."
        return $null
    }
}


# --- Main Script Logic ---

# Check connection
try { Get-ConnectionInformation | Out-Null; Write-Verbose "Connection verified." } catch { Write-Error "Not connected. Run Connect-ExchangeOnline."; return }

# Validate CSV Path
if (-not (Test-Path -Path $CsvPath -PathType Leaf)) { Write-Error "CSV not found: $CsvPath"; return }

# Import CSV
try { $importData = Import-Csv -Path $CsvPath } catch { Write-Error "Failed import: $CsvPath. Error: $($_.Exception.Message)"; return }
if ($null -eq $importData -or $importData.Count -eq 0) { Write-Warning "CSV '$CsvPath' empty/unreadable."; return }

# Verify headers
$requiredHeaders = @($CsvHeaderGroupEmail, $CsvHeaderManagers, $CsvHeaderMembers, $CsvHeaderOwners)
$actualHeaders = $importData[0].PSObject.Properties.Name
$missingHeaders = $requiredHeaders | Where-Object { $actualHeaders -notcontains $_ }
if ($missingHeaders.Count -gt 0) { Write-Error "CSV missing header(s): '$($missingHeaders -join "', '")'. Check CSV or parameters."; return }
Write-Verbose "CSV headers verified."

Write-Host "Starting processing of $($importData.Count) groups from CSV."

# Process each row (group)
foreach ($row in $importData) {
    $googleGroupEmail = $row.$($CsvHeaderGroupEmail).Trim()

    Write-Host ("-"*40)
    Write-Host "Processing Group Row for: $googleGroupEmail"

    if ([string]::IsNullOrWhiteSpace($googleGroupEmail)){ Write-Warning "Skipping row (index $($importData.IndexOf($row))) due to empty Group Email ('$CsvHeaderGroupEmail')."; continue }

    # 1. Determine Target M365 Email (Simplified)
    $m365TargetEmail = Get-TargetM365GroupEmail -GoogleGroupEmail $googleGroupEmail
    if (-not $m365TargetEmail) { Write-Warning "Skipping group '$googleGroupEmail' - could not determine target M365 email."; continue }
    Write-Verbose "Target M365 Email/Identity: $m365TargetEmail"

    # 2. Identify Target Recipient Type in M365
    $targetRecipient = $null
    try {
        $targetRecipient = Get-Recipient -Identity $m365TargetEmail -ErrorAction Stop
        Write-Verbose "Found target recipient '$m365TargetEmail'. Type: $($targetRecipient.RecipientTypeDetails)."
    }
    catch [System.Management.Automation.ItemNotFoundException] { Write-Warning "Target '$m365TargetEmail' not found in Microsoft 365. Skipping."; continue }
    catch { Write-Warning "Error accessing target '$m365TargetEmail'. Skipping. Error: $($_.Exception.Message)"; continue }

    # 3. Prepare lists of potential members and owners from CSV row
    $potentialOwners = [System.Collections.Generic.List[string]]::new()
    $potentialMembers = [System.Collections.Generic.List[string]]::new()
    $processedOwnerUPNs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $processedMemberUPNs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    # Parse Owners/Managers from CSV -> potentialOwners list
    $managerString = $row.$($CsvHeaderManagers); $ownerString = $row.$($CsvHeaderOwners)
    $ownerManagerEmailsFromCsv = @()
    if (-not [string]::IsNullOrWhiteSpace($managerString)) { $ownerManagerEmailsFromCsv += $managerString.Split(' ',[System.StringSplitOptions]::RemoveEmptyEntries) }
    if (-not [string]::IsNullOrWhiteSpace($ownerString)) { $ownerManagerEmailsFromCsv += $ownerString.Split(' ',[System.StringSplitOptions]::RemoveEmptyEntries) }

    foreach ($ggEmail in $ownerManagerEmailsFromCsv) {
         $m365UPN = Convert-GoogleEmailToM365UPN -GoogleEmail $ggEmail
         if ($m365UPN -and $processedOwnerUPNs.Add($m365UPN)) { $potentialOwners.Add($m365UPN) }
         # Warnings for failed mapping handled in function
    }
    Write-Verbose "Prepared $($potentialOwners.Count) potential owner UPNs."

    # Parse Members from CSV -> potentialMembers list
    $memberString = $row.$($CsvHeaderMembers); $memberEmailsFromCsv = @()
    if (-not [string]::IsNullOrWhiteSpace($memberString)) { $memberEmailsFromCsv = $memberString.Split(' ',[System.StringSplitOptions]::RemoveEmptyEntries) }

    foreach ($ggEmail in $memberEmailsFromCsv) {
         $m365UPN = Convert-GoogleEmailToM365UPN -GoogleEmail $ggEmail
         if ($m365UPN -and $processedMemberUPNs.Add($m365UPN)) { $potentialMembers.Add($m365UPN) }
         # Warnings for failed mapping handled in function
    }
    Write-Verbose "Prepared $($potentialMembers.Count) potential member UPNs."

    # 4. Process based on Target Recipient Type
    switch ($targetRecipient.RecipientTypeDetails) {
        # --- Microsoft 365 Group ---
        "GroupMailbox" {
            Write-Host "Target is a Microsoft 365 Group. Processing Owners and Members separately."

            # Optimize: Remove members who are also owners
            $membersToAddOptimized = $potentialMembers | Where-Object { -not $processedOwnerUPNs.Contains($_) }
            $removedCount = $potentialMembers.Count - $membersToAddOptimized.Count
            if ($removedCount -gt 0) { Write-Verbose "Optimization: Removed $removedCount member(s) already queued as owners."}

            # Add Owners
            if ($potentialOwners.Count -gt 0) {
                Write-Host "Attempting to add $($potentialOwners.Count) owner(s)..."
                if ($PSCmdlet.ShouldProcess($m365TargetEmail, "Add Owners (UnifiedGroupLinks): $($potentialOwners -join ', ')")) {
                    try { Add-UnifiedGroupLinks -Identity $m365TargetEmail -LinkType Owners -Links $potentialOwners -ErrorAction Stop; Write-Host "Owner add success."} catch { Write-Warning "Owner add failed for '$m365TargetEmail'. Error: $($_.Exception.Message)" }
                }
            } else { Write-Verbose "No new owners to add." }

            # Add Members (Optimized List)
            if ($membersToAddOptimized.Count -gt 0) {
                Write-Host "Attempting to add $($membersToAddOptimized.Count) member(s)..."
                if ($PSCmdlet.ShouldProcess($m365TargetEmail, "Add Members (UnifiedGroupLinks): $($membersToAddOptimized -join ', ')")) {
                    try { Add-UnifiedGroupLinks -Identity $m365TargetEmail -LinkType Members -Links $membersToAddOptimized -ErrorAction Stop; Write-Host "Member add success." } catch { Write-Warning "Member add failed for '$m365TargetEmail'. Error: $($_.Exception.Message)" }
                }
            } else { Write-Verbose "No new members (excluding owners) to add." }
        }

        # --- Distribution List ---
        "MailUniversalDistributionGroup" { Write-Warning "Target '$m365TargetEmail' is a Distribution List. Google Owners/Managers will be added as Members."; $isDLOrMESG = $true }
        "MailNonUniversalDistributionGroup" { Write-Warning "Target '$m365TargetEmail' is a Distribution List. Google Owners/Managers will be added as Members."; $isDLOrMESG = $true }

        # --- Mail-Enabled Security Group ---
        "MailUniversalSecurityGroup" { Write-Warning "Target '$m365TargetEmail' is a Mail-Enabled Security Group. Google Owners/Managers will be added as Members."; $isDLOrMESG = $true }

        # --- Default Case ---
        default {
            Write-Warning "Target '$m365TargetEmail' is of type '$($targetRecipient.RecipientTypeDetails)', which is not handled by this script's membership logic. Skipping additions."
            $isDLOrMESG = $false # Ensure DL/MESG logic doesn't run
        }
    } # End Switch

    # --- Logic for DLs and MESGs (if $isDLOrMESG was set to $true) ---
    if ($isDLOrMESG) {
        # Combine all potential owners and members into one list for DLs/MESGs
        $allMembersToAddSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $potentialOwners | ForEach-Object { $allMembersToAddSet.Add($_) | Out-Null }
        $potentialMembers | ForEach-Object { $allMembersToAddSet.Add($_) | Out-Null }
        $allMembersToAddList = [System.Collections.Generic.List[string]]($allMembersToAddSet)

        if ($allMembersToAddList.Count -gt 0) {
            Write-Host "Attempting to add $($allMembersToAddList.Count) members (incl. Google Owners/Mgrs) to DL/MESG '$m365TargetEmail'..."
            if ($PSCmdlet.ShouldProcess($m365TargetEmail, "Add Members (DistributionGroupMember): $($allMembersToAddList -join ', ')")) {
                # Add-DistributionGroupMember adds one member at a time, need to loop
                $failedMembers = [System.Collections.Generic.List[string]]::new()
                foreach ($memberUPN in $allMembersToAddList) {
                    try {
                        Add-DistributionGroupMember -Identity $m365TargetEmail -Member $memberUPN -BypassSecurityGroupManagerCheck -ErrorAction Stop
                        Write-Verbose "  Successfully added member $memberUPN"
                    } catch {
                        Write-Warning "  Failed to add member '$memberUPN' to '$m365TargetEmail'. Error: $($_.Exception.Message)"
                        $failedMembers.Add($memberUPN)
                    }
                }
                if ($failedMembers.Count -eq 0) { Write-Host "Finished adding members to DL/MESG." }
                else { Write-Warning "Finished adding members to DL/MESG, but $($failedMembers.Count) addition(s) failed."}
            }
        } else {
            Write-Verbose "No new members identified/mapped to add to DL/MESG '$m365TargetEmail'."
        }
        $isDLOrMESG = $false # Reset flag for next loop iteration
    }

} # End foreach row

Write-Host ("-"*40)
Write-Host "Script finished."