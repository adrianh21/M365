<#
.SYNOPSIS
    Updates extensionAttribute2 for all Microsoft 365 users with their manager's display name.

.DESCRIPTION
    This script connects to Microsoft Graph to iterate through all member users (excluding guests) 
    in the tenant. For each user, it makes a specific call to find their manager. If a manager is found,
    it updates the user's extensionAttribute2 with the manager's display name.

.NOTES
    Prerequisites:
    1.  PowerShell 5.1 or later.
    2.  The Microsoft.Graph PowerShell module. To install, run:
        Install-Module Microsoft.Graph -Scope CurrentUser
    3.  An account with sufficient permissions. You will be prompted to consent to
        the required scopes upon connection.
#>

#region --- Connection and Authentication ---

try {
    # Install the module if it's not already installed
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Host "Microsoft Graph module not found. Installing..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
    }

    # Define the required permission scopes, including Directory.Read.All for manager lookup.
    $requiredScopes = @("User.ReadWrite.All", "User.Read.All", "Directory.Read.All")

    # Connect to Microsoft Graph
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $requiredScopes
    
    Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph. Please check your connection and permissions. Error: $_"
    # Stop the script if connection fails
    return
}

#endregion

#region --- Main Processing Logic ---

try {
    # Get all member users, specifically selecting the onPremisesExtensionAttributes property.
    # The -Filter "userType eq 'Member'" ensures the exclusion of guests and deactivated users.
    Write-Host "Fetching all member users from the tenant. This may take a moment for large organizations..." -ForegroundColor Cyan
    $allUsers = Get-MgUser -All -Filter "userType eq 'Member' and accountEnabled eq true" -Property Id, DisplayName, UserPrincipalName, onPremisesExtensionAttributes

    Write-Host "Found $($allUsers.Count) member users. Starting update process..." -ForegroundColor Cyan
    Write-Host "------------------------------------------------------------"

    # Loop through each user
    foreach ($user in $allUsers) {
        Write-Host "Processing user: $($user.DisplayName) ($($user.UserPrincipalName))"
        
        $manager = $null # Reset manager variable for each iteration
        $hasManager = $true # Assume user has a manager until proven otherwise

        try {
            # Step 1: Get the manager reference object, which contains the manager's ID.
            $managerRef = Get-MgUserManager -UserId $user.Id -ErrorAction Stop
            
            # Step 2: Use the manager's ID to get their full user object, specifically requesting DisplayName.
            $manager = Get-MgUser -UserId $managerRef.Id -Property "displayName"
        }
        catch {
            # This block will catch any error from the lookup process.
            if ($_.Exception.Message -like "*Request_ResourceNotFound*") {
                # This is the expected error for a user with no manager.
                $hasManager = $false
                Write-Host "  -> No manager assigned to this user. Skipping." -ForegroundColor Yellow
            } else {
                # If it's a different error, log it for debugging.
                $hasManager = $false
                Write-Warning "  -> An unexpected error occurred while fetching manager. Full Error: $($_.ToString())"
            }
        }

        # This block only runs if the try{} block succeeded and the catch{} block was skipped.
        if ($hasManager) {
            if ($null -ne $manager -and -not [string]::IsNullOrEmpty($manager.DisplayName)) {
                try {
                    $managerName = $manager.DisplayName
                    Write-Host "  -> Manager found: '$managerName'." -ForegroundColor White

                    # Correctly access the current attribute value from the OnPremisesExtensionAttributes object.
                    $currentAttributeValue = $user.OnPremisesExtensionAttributes.ExtensionAttribute2
                    
                    if ($currentAttributeValue -eq $managerName) {
                        Write-Host "  -> extensionAttribute2 is already correctly set. No update needed." -ForegroundColor Gray
                    } else {
                        Write-Host "  -> Updating extensionAttribute2..." -ForegroundColor White
                        
                        # Prepare the parameters for updating the user.
                        $updateParams = @{
                            OnPremisesExtensionAttributes = @{
                                ExtensionAttribute2 = $managerName
                            }
                        }
                        Update-MgUser -UserId $user.Id -BodyParameter $updateParams
                        # This success message only runs if the update command does not throw an error.
                        Write-Host "  -> Successfully updated extensionAttribute2 to '$managerName'." -ForegroundColor Green
                    }
                }
                catch {
                    # This inner catch handles potential issues during the attribute update process itself.
                    Write-Warning "  -> An error occurred while updating the attribute for user $($user.DisplayName). Error: $($_.Exception.Message)"
                }
            } else {
                # This handles the case where a manager is assigned, but their details could not be retrieved.
                Write-Warning "  -> Manager is assigned, but their details (DisplayName) could not be retrieved."
            }
        }
        Write-Host "------------------------------------------------------------"
    }
}
catch {
    Write-Error "An unexpected error occurred during user processing. Error: $_"
}
finally {
    # Disconnect from Microsoft Graph
    Write-Host "Processing complete. Disconnecting from Microsoft Graph." -ForegroundColor Cyan
    Disconnect-MgGraph
}

#endregion
