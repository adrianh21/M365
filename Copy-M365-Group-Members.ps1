# --- Define Source and Destination Groups ---
$SourceGroupEmail = ""
$DestinationGroupEmail = ""

# --- Get Members from the Source Group ---
Write-Host "Attempting to retrieve members from '$SourceGroupEmail'..." -ForegroundColor Cyan
Try {
    # Get the full member objects first
    $sourceMemberObjects = Get-UnifiedGroupLinks -Identity $SourceGroupEmail -LinkType Members -ResultSize Unlimited -ErrorAction Stop
    # Extract the PrimarySmtpAddress for easier processing
    $sourceMembers = $sourceMemberObjects | Select-Object -ExpandProperty PrimarySmtpAddress
    Write-Host "Successfully found $($sourceMembers.Count) members in '$SourceGroupEmail'." -ForegroundColor Green
} Catch {
    Write-Error "Failed to retrieve members from '$SourceGroupEmail'. Error: $($_.Exception.Message). Please check the group email and your permissions."
    Return # Stop script execution
}

# --- Get Members from the Destination Group FOR CHECKING ---
Write-Host "Attempting to retrieve existing members from '$DestinationGroupEmail' for comparison..." -ForegroundColor Cyan
$destinationMemberSet = $null # Ensure it's null initially

Try {
    # Initialize the HashSet with the case-insensitive comparer FIRST
    $destinationMemberSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    # Get the member objects' email addresses
    $destinationMemberEmails = $null # Reset before retrieving
    $destinationMemberEmails = (Get-UnifiedGroupLinks -Identity $DestinationGroupEmail -LinkType Members -ResultSize Unlimited -ErrorAction Stop).PrimarySmtpAddress

    # Populate the HashSet IF members were found
    if ($null -ne $destinationMemberEmails) {
        # Handle PowerShell potentially returning a single string OR an array
        if ($destinationMemberEmails -is [array]) {
             # Multiple members found
             foreach ($email in $destinationMemberEmails) {
                # Add check for null/empty strings within the list and Trim whitespace
                if (-not [string]::IsNullOrEmpty($email)) {
                   $destinationMemberSet.Add($email.Trim()) | Out-Null # Add returns true/false, pipe to Out-Null to suppress output
                }
             }
        } elseif (-not [string]::IsNullOrEmpty($destinationMemberEmails)) {
             # Only one member was returned, handle as single string
             $destinationMemberSet.Add($destinationMemberEmails.Trim()) | Out-Null
        }
    }
    # If $destinationMemberEmails was $null or empty, the set remains empty, which is correct.

    Write-Host "Found $($destinationMemberSet.Count) existing members in '$DestinationGroupEmail'." -ForegroundColor Green

} Catch {
    Write-Warning "Could not fully retrieve or process members from '$DestinationGroupEmail' to check for existing memberships. Accuracy of 'already exists' check might be affected. Error: $($_.Exception.Message)"
    # Ensure the set is initialized as empty if *any* error occurred, so the later '.Contains' check doesn't fail
    if ($null -eq $destinationMemberSet) {
        $destinationMemberSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
    # Note: The script will proceed, but the pre-check might not be accurate if this block was hit.
}

# --- Add Members to the Destination Group ---

If ($sourceMembers) {
    Write-Host "Starting to process members to add/sync to '$DestinationGroupEmail'..." -ForegroundColor Cyan
     $addedCount = 0
     $skippedCount = 0
     $errorCount = 0

     ForEach ($memberIdentity in $sourceMembers) {
         # Trim whitespace just in case
         $memberIdentity = $memberIdentity.Trim()

         # Skip if the identity is empty/null
         if ([string]::IsNullOrEmpty($memberIdentity)) {
             Write-Warning "Skipping an empty or null member identity found from the source group."
             $skippedCount++
             continue
         }

         Write-Host "Processing member: $memberIdentity"

         if ($destinationMemberSet.Contains($memberIdentity)) {
             # Member already exists in the destination group
             Write-Host "$memberIdentity is already a member of '$DestinationGroupEmail'." -ForegroundColor Yellow
             $skippedCount++
         } else {
            # Member does NOT exist, attempt to add
            Try {
                 Add-UnifiedGroupLinks -Identity $DestinationGroupEmail -LinkType Members -Links $memberIdentity -Confirm:$false -ErrorAction Stop
                 Write-Host "Successfully added $memberIdentity to '$DestinationGroupEmail'." -ForegroundColor Green
                 $addedCount++
             } Catch {
                 Write-Warning "Could not add member $memberIdentity to '$DestinationGroupEmail'. Error: $($_.Exception.Message)"
                 $errorCount++
             }
         }
     }
     Write-Host "`n--- Sync Summary ---" -ForegroundColor Yellow
     Write-Host "Processed $($sourceMembers.Count) source members."
     Write-Host "Newly Added: $addedCount" -ForegroundColor Green
     Write-Host "Already Existed / Skipped: $skippedCount" -ForegroundColor Yellow
     Write-Host "Errors during add attempt: $errorCount" -ForegroundColor Red
     Write-Host "--------------------`n"
     Write-Host "Finished processing members." -ForegroundColor Cyan


} Else {
    Write-Host "No members found in '$SourceGroupEmail' to copy." -ForegroundColor Yellow
}
