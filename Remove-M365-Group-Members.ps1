<#
.SYNOPSIS
Synchronizes Microsoft 365 Group membership based on a CSV file using hardcoded values.
Removes members from the specified group who are not listed in the CSV.

.DESCRIPTION
This script connects to Exchange Online, reads a list of required members (by UserPrincipalName/Email)
from a CSV file, gets the current members of a specified Microsoft 365 Group,
and removes any current members who are NOT found in the CSV file.
The Group Email and CSV Path are hardcoded in the script.

.EXAMPLE
.\Sync-GroupMembership-Hardcoded.ps1
(No parameters needed as they are set inside the script)

.NOTES
- Requires the ExchangeOnlineManagement PowerShell module. Consider running 'Update-Module ExchangeOnlineManagement' for the latest features/fixes.
- Run PowerShell as Administrator.
- The account running the script needs permissions to manage group memberships (e.g., Global Admin, Exchange Admin, Groups Admin).
#>
param(
    [string]$GroupEmail = "", # Define the email address of the group you need to remove members of here.
    [string]$CsvPath = "" # Define the path to the CSV list of users here.
)

#region Prerequisites and Connection
Write-Host "Using Group Email: $GroupEmail"
Write-Host "Using CSV Path:    $CsvPath"

Write-Host "Checking for ExchangeOnlineManagement module..." -ForegroundColor Gray
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Warning "ExchangeOnlineManagement module not found. Please install it using 'Install-Module ExchangeOnlineManagement'"
    exit 1
} else {
     Write-Host "ExchangeOnlineManagement module found." -ForegroundColor Gray
     Write-Host "Note: If you encounter connection issues, consider updating with: Update-Module ExchangeOnlineManagement" -ForegroundColor Cyan
}

Write-Host "Connecting to Exchange Online..." -ForegroundColor Gray
try {
    # Simplified connection - relies solely on Connect-ExchangeOnline
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host "Successfully connected/verified connection to Exchange Online." -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Exchange Online. Please check your credentials, permissions, and ensure MFA prompts (if any) are handled. Error: $($_.Exception.Message)"
    exit 1
}
#endregion

#region Parameter Validation
Write-Host "Validating parameters..." -ForegroundColor Gray
if (-not (Test-Path -Path $CsvPath -PathType Leaf)) {
    Write-Error "CSV file not found at path: $CsvPath"
    exit 1
} else {
    Write-Host "CSV file found." -ForegroundColor Green
}

Write-Host "Checking target group '$GroupEmail'..." -ForegroundColor Gray
$group = Get-UnifiedGroup -Identity $GroupEmail -ErrorAction SilentlyContinue
if (-not $group)
{
    Write-Error "Microsoft 365 Group '$GroupEmail' not found. Ensure it's a Microsoft 365 Group (Unified Group)."
    exit 1
}
Write-Host "Group '$($group.DisplayName)' found." -ForegroundColor Green
#endregion

#region Main Logic
Write-Host "Reading members to keep from CSV: $CsvPath" -ForegroundColor Gray
$membersToKeep = @() # Initialize as empty array.
try {
    # Import the CSV content.
    $csvContent = Import-Csv -Path $CsvPath -Encoding UTF8 -ErrorAction SilentlyContinue

    if ($null -ne $csvContent) {
        # Find a header like 'Email' or 'UserPrincipalName' case-insensitively - ** VERIFIED CLEAN LINE **
        $emailHeader = $csvContent | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match '^Email$' -or $_.Name -match '^UserPrincipalName$' } | Select-Object -First 1 -ExpandProperty Name

        if ($emailHeader) {
            $membersToKeep = $csvContent | Select-Object -ExpandProperty $emailHeader
        } else {
             # If specific headers not found, try using the first column header.
             $firstHeader = ($csvContent | Get-Member -MemberType NoteProperty)[0].Name
             if ($firstHeader) { # Ensure a header was actually found.
                Write-Warning "Could not find 'Email' or 'UserPrincipalName' header. Using first column header '$firstHeader' instead."
                $membersToKeep = $csvContent | Select-Object -ExpandProperty $firstHeader
             } else {
                 Write-Warning "Could not detect any headers in the CSV file '$CsvPath'."
                 $membersToKeep = @() # Treat as empty if no headers.
             }
        }
    } else {
        # Handle empty file case or file read error.
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -Path $CsvPath -PathType Leaf)) {
             Write-Warning "Failed to import CSV content from '$CsvPath'. It might be empty, locked, or improperly formatted. Error: $($Error[0].Exception.Message)"
        } else {
             Write-Warning "CSV file '$CsvPath' appears to be empty."
        }
        $membersToKeep = @()
    }

    # Ensure $membersToKeep is an array even if only one row/column was read directly
    if ($null -ne $membersToKeep -and $membersToKeep -isnot [array]) {
        $membersToKeep = @($membersToKeep)
    }

    if ($null -eq $membersToKeep) {
        # Fallback if it's somehow still null
        $membersToKeep = @()
    }

    $membersToKeep = $membersToKeep | ForEach-Object { $_.ToLowerInvariant().Trim() } | Where-Object { $_ -ne '' } | Sort-Object -Unique

    if ($membersToKeep.Count -eq 0) {
        Write-Warning "After processing, CSV file '$CsvPath' provided no valid email addresses."
        Write-Warning "WARNING: Proceeding will attempt to remove ALL members from the group '$GroupEmail' as the source list is empty."
        $confirmation = Read-Host "Type 'CONFIRM REMOVE ALL' to continue, any other key to exit."
        if ($confirmation -ne 'CONFIRM REMOVE ALL') {
            Write-Host "Operation cancelled by user."
            exit 1
        }
         Write-Host "Proceeding with removal of all members as confirmed by user." -ForegroundColor Yellow
    } else {
         Write-Host "Found $($membersToKeep.Count) unique members to keep listed in the CSV." -ForegroundColor Cyan
    }
}
catch {
    Write-Error "An unexpected error occurred while reading or processing CSV file '$CsvPath'. Error: $($_.Exception.Message)"
    exit 1
}

Write-Host "Getting current members of group '$GroupEmail'..." -ForegroundColor Gray
$currentMembers = @() # Initialize as empty array.
try {
    $currentMembersResult = Get-UnifiedGroupLinks -Identity $GroupEmail -LinkType Members -ResultSize Unlimited

    if ($null -eq $currentMembersResult) {
        Write-Host "Group '$GroupEmail' has no members." -ForegroundColor Cyan
        $currentMembers = @()
    } elseif ($currentMembersResult -is [array]) {
         # Filter out potential null/empty entries from the command output.
         $currentMembers = $currentMembersResult | Select-Object -ExpandProperty PrimarySmtpAddress | Where-Object { $null -ne $_ -and $_ -ne '' }
    } elseif ($null -ne $currentMembersResult -and $currentMembersResult.PrimarySmtpAddress) { # Single member case, ensure PrimarySmtpAddress exists and is not null
        $currentMembers = @($currentMembersResult.PrimarySmtpAddress)
    } else {
        # Handle unexpected single result format or null PrimarySmtpAddress.
        Write-Warning "Received unexpected result format or null PrimarySmtpAddress from Get-UnifiedGroupLinks for a single member result."
        $currentMembers = @()
    }

    # Normalize emails.
    $currentMembers = $currentMembers | ForEach-Object { $_.ToLowerInvariant().Trim() } | Where-Object { $_ -ne '' } | Sort-Object -Unique
    if ($currentMembers.Count -gt 0) {
         Write-Host "Found $($currentMembers.Count) current members in the group after processing." -ForegroundColor Cyan
    }

}
catch {
    Write-Error "Failed to get or process current members for group '$GroupEmail'. The error was: $($_.Exception.Message)"
    exit 1
}

# Calculate members to remove.
Write-Host "Calculating members to remove..." -ForegroundColor Gray
$membersToKeepHashTable = @{}
$membersToKeep | ForEach-Object { $membersToKeepHashTable[$_] = $true }

$membersToRemove = $currentMembers | Where-Object { -not $membersToKeepHashTable.ContainsKey($_) }

if ($membersToRemove.Count -gt 0) {
    Write-Host "$($membersToRemove.Count) members identified for removal (currently in group but not in CSV):" -ForegroundColor Yellow
    ($membersToRemove | Sort-Object) | ForEach-Object { Write-Host "- $_" }
    Write-Host "Attempting removal using Remove-UnifiedGroupLinks..." -ForegroundColor Gray
    try {
        $batchSize = 100
        $startIndex = 0
        while($startIndex -lt $membersToRemove.Count){
            $batch = $membersToRemove[$startIndex..([System.Math]::Min($startIndex + $batchSize - 1, $membersToRemove.Count - 1))]
            Write-Host "Processing removal batch (StartIndex: $startIndex, Count: $($batch.Count))..." -ForegroundColor Gray
            Remove-UnifiedGroupLinks -Identity $GroupEmail -LinkType Members -Links $batch -Confirm:$false -ErrorAction Stop
            $startIndex += $batchSize
        }
        Write-Host "Removal complete. Review the output above." -ForegroundColor Green
    }
    catch {
         Write-Error "An error occurred during the removal process. Error: $($_.Exception.Message)"
    }
}
else {
    Write-Host "No members need to be removed. The group membership already matches the CSV list." -ForegroundColor Green
}
#endregion

#region Cleanup
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Gray
# Use Confirm:$false to disconnect without prompting
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Script finished."
#endregion