<#
.SYNOPSIS
Imports a list of email addresses from a CSV file into a specified Microsoft 365 Group.

.DESCRIPTION
This script reads email addresses from a CSV file (which must have a header named 'Email')
and adds each email address as a member to the specified Microsoft 365 Group (Unified Group).
It connects to Exchange Online if not already connected and requires appropriate admin permissions.

.PARAMETER CsvPath
The full path to the CSV file containing the email addresses.
Example: "C:\temp\members.csv"

.PARAMETER GroupEmail
The primary email address of the target Microsoft 365 Group.
Example: "my-project-group@yourdomain.com"

.EXAMPLE
.\Import-EmailsToM365Group.ps1 -CsvPath "C:\data\new_hires.csv" -GroupEmail "all.staff@yourcompany.com"

.NOTES
- Requires installation of the ExchangeOnlineManagement module: Install-Module ExchangeOnlineManagement
- You will be prompted to log in to Microsoft 365 if not already connected via Connect-ExchangeOnline.
- The account running the script needs permissions to modify group membership.
- Errors during addition (e.g., user not found, already a member) will be displayed as warnings.
#>
param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})] # Basic check that the file exists
    [string]$CsvPath,

    [Parameter(Mandatory=$true)]
    [string]$GroupEmail
)

# --- Configuration ---
$ErrorActionPreference = "Stop" # Stop script on most errors, allows Try/Catch to work reliably

# --- Check and Connect to Exchange Online ---
Write-Host "Checking Exchange Online connection..."
try {
    # Check if already connected
    Get-ConnectionInformation | Where-Object {$_.ConnectionType -eq 'ExchangeOnline'} | Out-Null
    Write-Host "Already connected to Exchange Online." -ForegroundColor Cyan
} catch {
    # If not connected, attempt to connect
    Write-Host "Not connected. Attempting to connect to Exchange Online..."
    try {
        # Suppress the banner, connect, and check the connection again
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Get-ConnectionInformation | Where-Object {$_.ConnectionType -eq 'ExchangeOnline'} | Out-Null
        Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Exchange Online. Please ensure the ExchangeOnlineManagement module is installed and you have internet access. Error: $($_.Exception.Message)"
        # Exit the script if connection fails
        return
    }
}

# --- Validate Group Existence ---
Write-Host "Verifying the target group '$GroupEmail' exists..."
try {
    $targetGroup = Get-UnifiedGroup -Identity $GroupEmail
    Write-Host "Group '$($targetGroup.DisplayName)' found." -ForegroundColor Green
} catch {
    Write-Error "Could not find the Microsoft 365 Group with email '$GroupEmail'. Please verify the group email address. Error: $($_.Exception.Message)"
    return # Exit if group not found
}

# --- Import CSV Data ---
Write-Host "Importing email addresses from '$CsvPath'..."
try {
    $importData = Import-Csv -Path $CsvPath
    # Verify the 'Email' header exists
    if (-not ($importData | Get-Member -Name 'Email')) {
        Write-Error "The CSV file '$CsvPath' does not contain the required header column named 'Email'."
        return
    }
    Write-Host "Found $($importData.Count) records in the CSV file."
} catch {
    Write-Error "Failed to import or read the CSV file '$CsvPath'. Error: $($_.Exception.Message)"
    return
}

# --- Process Members ---
$addedCount = 0
$skippedCount = 0
$errorCount = 0

Write-Host "Starting to add members to group '$GroupEmail'..."

foreach ($row in $importData) {
    # Get the email address, trim whitespace
    $emailToAdd = $row.Email.Trim()

    # Skip blank entries
    if ([string]::IsNullOrWhiteSpace($emailToAdd)) {
        Write-Warning "Skipping blank email address in CSV row number $($importData.IndexOf($row) + 2)." # +2 for 1-based index and header row
        $skippedCount++
        continue # Move to the next row
    }

    # Attempt to add the member
    try {
        Write-Host "Attempting to add member: '$emailToAdd'..."
        # Use -ErrorAction SilentlyContinue within the loop's try block to catch specific errors
        # Use Write-Host for success message inside the try
        Add-UnifiedGroupLinks -Identity $GroupEmail -LinkType Members -Links $emailToAdd -Confirm:$false -ErrorAction Stop
        Write-Host "Successfully added '$emailToAdd'." -ForegroundColor Green
        $addedCount++
    } catch {
        # Handle common, non-blocking errors gracefully
        $errorMessage = $_.Exception.Message
        if ($errorMessage -like "*The specified user*is already a group member*") {
            Write-Warning "User '$emailToAdd' is already a member of the group."
            $skippedCount++
        } elseif ($errorMessage -like "*Recipient not found*" -or $errorMessage -like "*couldn't find object*") {
            Write-Warning "Could not add '$emailToAdd'. Reason: User/Recipient not found in the directory."
            $errorCount++
        } else {
            # Log other unexpected errors
            Write-Warning "Failed to add '$emailToAdd'. Error: $($errorMessage)"
            $errorCount++
        }
        # Continue processing next email even if one fails
        continue
    }
}

# --- Summary ---
Write-Host "`n--- Import Summary ---" -ForegroundColor Yellow
Write-Host "Processed $($importData.Count) rows from CSV."
Write-Host "Successfully added: $addedCount" -ForegroundColor Green
Write-Host "Skipped (blank or already member): $skippedCount" -ForegroundColor Cyan
Write-Host "Errors (user not found, other issues): $errorCount" -ForegroundColor Red
Write-Host "--------------------`n"

Write-Host "Script finished."

# Note: Disconnection is often handled automatically by the module, but you can uncomment below if needed.
# Write-Host "Disconnecting from Exchange Online..."
# Disconnect-ExchangeOnline -Confirm:$false